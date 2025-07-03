#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import gspread
from google.oauth2.service_account import Credentials
import time
import logging
from datetime import datetime
import subprocess
import sys
import os

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SheetsMonitor:
    def __init__(self, credentials_path: str, spreadsheet_id: str):
        """
        Google Sheetsモニター
        
        Args:
            credentials_path: Google認証用JSONファイルのパス
            spreadsheet_id: スプレッドシートのID
        """
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.last_command = ""
        self.last_keyword = ""
        
        # Google Sheets接続
        try:
            scope = ['https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
            self.gc = gspread.authorize(creds)
            self.workbook = self.gc.open_by_key(spreadsheet_id)
            logger.info("Google Sheetsに接続しました")
        except Exception as e:
            logger.error(f"Google Sheets接続エラー: {e}")
            raise
    
    def setup_control_sheet(self):
        """
        コントロールシートの初期設定
        """
        try:
            # 'Control' シートを作成または取得
            try:
                control_sheet = self.workbook.worksheet('Control')
                logger.info("既存のControlシートを使用します")
            except gspread.exceptions.WorksheetNotFound:
                control_sheet = self.workbook.add_worksheet(title='Control', rows=10, cols=5)
                logger.info("新しいControlシートを作成しました")
            
            # ヘッダーを設定
            headers = [
                ['キーワード', '実行コマンド', 'ステータス', '最終実行時刻', '結果メッセージ'],
                ['', '', '待機中', '', '下記に使い方を記載'],
                ['', '', '', '', ''],
                ['【使い方】', '', '', '', ''],
                ['1. A列にキーワードを入力', '', '', '', ''],
                ['2. B列に「実行」と入力', '', '', '', ''],
                ['3. 自動でスクレイピングが開始されます', '', '', '', ''],
                ['4. 結果は別シートに保存されます', '', '', '', '']
            ]
            
            # 既存のデータがあるかチェック
            existing_keyword = control_sheet.cell(2, 1).value
            if not existing_keyword:
                control_sheet.update('A1:E8', headers)
                logger.info("Controlシートの初期設定が完了しました")
            
            return control_sheet
            
        except Exception as e:
            logger.error(f"Controlシート設定エラー: {e}")
            raise
    
    def check_for_commands(self, control_sheet):
        """
        新しいコマンドをチェック
        
        Returns:
            tuple: (keyword, command) または (None, None)
        """
        try:
            # B2セルの値を取得（実行コマンド）
            command = control_sheet.cell(2, 2).value
            keyword = control_sheet.cell(2, 1).value
            
            if not command or not keyword:
                return None, None
            
            command = str(command).strip().lower()
            keyword = str(keyword).strip()
            
            # 新しいコマンドかチェック
            current_state = f"{keyword}:{command}"
            if current_state != self.last_command and command == "実行":
                self.last_command = current_state
                return keyword, command
            
            return None, None
            
        except Exception as e:
            logger.error(f"コマンドチェックエラー: {e}")
            return None, None
    
    def update_status(self, control_sheet, status: str, message: str = ""):
        """
        ステータスを更新
        """
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            control_sheet.update('C2', status)
            control_sheet.update('D2', current_time)
            if message:
                control_sheet.update('E2', message)
            logger.info(f"ステータス更新: {status}")
        except Exception as e:
            logger.error(f"ステータス更新エラー: {e}")
    
    def execute_scraping(self, keyword: str, control_sheet):
        """
        スクレイピングを実行
        """
        try:
            self.update_status(control_sheet, "実行中", f"キーワード「{keyword}」で検索中...")
            
            # メインスクリプトを実行
            script_path = os.path.join(os.path.dirname(__file__), 'prtimes_corrected_scraper.py')
            cmd = [sys.executable, script_path, '--keyword', keyword]
            
            logger.info(f"実行中: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)  # 30分タイムアウト
            
            if result.returncode == 0:
                self.update_status(control_sheet, "完了", f"成功: {keyword}の検索が完了しました")
                # 実行コマンドをクリア
                control_sheet.update('B2', '')
                logger.info("スクレイピングが正常に完了しました")
            else:
                error_msg = result.stderr[:200] if result.stderr else "不明なエラー"
                self.update_status(control_sheet, "エラー", f"失敗: {error_msg}")
                logger.error(f"スクレイピングエラー: {error_msg}")
                
        except subprocess.TimeoutExpired:
            self.update_status(control_sheet, "タイムアウト", "30分でタイムアウトしました")
            logger.error("スクレイピングがタイムアウトしました")
        except Exception as e:
            self.update_status(control_sheet, "エラー", f"実行エラー: {str(e)[:200]}")
            logger.error(f"スクレイピング実行エラー: {e}")
    
    def run_monitor(self, check_interval: int = 10):
        """
        メイン監視ループ
        
        Args:
            check_interval: チェック間隔（秒）
        """
        logger.info("Google Sheetsモニタリングを開始します...")
        
        # コントロールシートの設定
        control_sheet = self.setup_control_sheet()
        self.update_status(control_sheet, "監視中", "コマンド待機中...")
        
        while True:
            try:
                # 新しいコマンドをチェック
                keyword, command = self.check_for_commands(control_sheet)
                
                if keyword and command:
                    logger.info(f"新しいコマンドを検出: キーワード='{keyword}', コマンド='{command}'")
                    self.execute_scraping(keyword, control_sheet)
                    # 実行後はステータスを監視中に戻す
                    time.sleep(5)  # 少し待ってからステータス更新
                    self.update_status(control_sheet, "監視中", "次のコマンド待機中...")
                
                time.sleep(check_interval)
                
            except KeyboardInterrupt:
                logger.info("監視を停止します...")
                self.update_status(control_sheet, "停止", "監視が停止されました")
                break
            except Exception as e:
                logger.error(f"監視ループエラー: {e}")
                time.sleep(30)  # エラー時は長めに待機

def main():
    """
    メイン関数
    """
    try:
        import config
        credentials_path = config.GOOGLE_CREDENTIALS_PATH
        spreadsheet_id = config.SPREADSHEET_ID
    except ImportError:
        logger.error("config.py が見つかりません")
        credentials_path = 'credentials.json'
        spreadsheet_id = 'your_spreadsheet_id'
        
    if not os.path.exists(credentials_path):
        logger.error(f"認証ファイルが見つかりません: {credentials_path}")
        return
    
    monitor = SheetsMonitor(credentials_path, spreadsheet_id)
    monitor.run_monitor()

if __name__ == '__main__':
    main()