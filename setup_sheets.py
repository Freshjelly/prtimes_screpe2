#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import gspread
from google.oauth2.service_account import Credentials
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_google_sheets():
    """
    Google Sheetsの初期設定を行う
    """
    try:
        import config
        credentials_path = config.GOOGLE_CREDENTIALS_PATH
        spreadsheet_id = config.SPREADSHEET_ID
    except ImportError:
        logger.error("config.py が見つかりません")
        return False
    
    try:
        # Google Sheets接続
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
        gc = gspread.authorize(creds)
        workbook = gc.open_by_key(spreadsheet_id)
        
        # Controlシートを作成または取得
        try:
            control_sheet = workbook.worksheet('Control')
            logger.info("既存のControlシートを使用します")
        except gspread.exceptions.WorksheetNotFound:
            control_sheet = workbook.add_worksheet(title='Control', rows=15, cols=6)
            logger.info("新しいControlシートを作成しました")
        
        # データをクリア
        control_sheet.clear()
        
        # ヘッダーと説明を設定
        setup_data = [
            ['🔍 PR Times スクレイパー コントロールパネル', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['キーワード', '実行コマンド', 'ステータス', '最終実行時刻', '結果メッセージ', ''],
            ['', '', '待機中', '', 'ここにキーワードを入力 →', ''],
            ['', '', '', '', '', ''],
            ['【📝 使い方】', '', '', '', '', ''],
            ['1️⃣ A4セルにキーワードを入力してください', '', '', '', '', ''],
            ['   （例：美容、健康食品、AI、化粧品など）', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['2️⃣ B4セルに「実行」と入力してください', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['3️⃣ 自動でスクレイピングが開始されます', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['4️⃣ 結果は「prtimes_sc2」シートに保存されます', '', '', '', '', ''],
            ['', '', '', '', '', '']
        ]
        
        # データを書き込み
        control_sheet.update('A1:F15', setup_data)
        
        # セルの書式設定（タイトル行を太字にする）
        try:
            # タイトル行の書式設定
            control_sheet.format('A1:F1', {
                'textFormat': {'bold': True, 'fontSize': 14},
                'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 1.0}
            })
            
            # ヘッダー行の書式設定
            control_sheet.format('A3:F3', {
                'textFormat': {'bold': True},
                'backgroundColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95}
            })
            
            # 入力エリアの書式設定
            control_sheet.format('A4:B4', {
                'backgroundColor': {'red': 1.0, 'green': 1.0, 'blue': 0.9}
            })
            
            logger.info("セルの書式設定が完了しました")
        except Exception as e:
            logger.warning(f"書式設定に失敗しましたが、続行します: {e}")
        
        logger.info("Google Sheetsの初期設定が完了しました！")
        logger.info(f"Google Sheetsを開いて 'Control' シートを確認してください")
        logger.info(f"URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
        
        return True
        
    except Exception as e:
        logger.error(f"Google Sheets設定エラー: {e}")
        return False

if __name__ == '__main__':
    setup_google_sheets()