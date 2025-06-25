#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import logging
from typing import List, Dict, Optional

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PRTimesScraper:
    def __init__(self, email: str, password: str, credentials_path: str):
        """
        PR Timesスクレイパーの初期化
        
        Args:
            email: PR Timesログイン用メールアドレス
            password: PR Timesログイン用パスワード
            credentials_path: Google認証用JSONファイルのパス
        """
        self.email = email
        self.password = password
        self.credentials_path = credentials_path
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        
    def login(self) -> bool:
        """
        PR Timesにログイン
        
        Returns:
            bool: ログイン成功時True、失敗時False
        """
        try:
            # ログインページにアクセスしてCSRFトークンを取得
            login_page_url = 'https://prtimes.jp/login'
            response = self.session.get(login_page_url)
            response.encoding = 'utf-8'
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # CSRFトークンを探す（複数のパターンに対応）
            csrf_token = None
            
            # パターン1: <input type="hidden" name="csrf_token" value="...">
            csrf_input = soup.find('input', {'name': 'csrf_token'})
            if csrf_input:
                csrf_token = csrf_input.get('value')
            
            # パターン2: <meta name="csrf-token" content="...">
            if not csrf_token:
                csrf_meta = soup.find('meta', {'name': 'csrf-token'})
                if csrf_meta:
                    csrf_token = csrf_meta.get('content')
            
            # ログインPOSTリクエスト
            login_url = 'https://prtimes.jp/login'
            login_data = {
                'email': self.email,
                'password': self.password
            }
            
            if csrf_token:
                login_data['csrf_token'] = csrf_token
                
            response = self.session.post(login_url, data=login_data)
            
            # ログイン成功の確認
            if response.status_code == 200 and 'logout' in response.text:
                logger.info("ログインに成功しました")
                return True
            else:
                logger.error("ログインに失敗しました")
                return False
                
        except Exception as e:
            logger.error(f"ログイン中にエラーが発生しました: {e}")
            return False
    
    def fetch_articles(self, keyword: str, max_articles: int = 50) -> List[str]:
        """
        キーワードで検索し、記事URLを収集
        
        Args:
            keyword: 検索キーワード
            max_articles: 収集する最大記事数
            
        Returns:
            List[str]: 記事URLのリスト
        """
        article_urls = []
        page = 1
        
        while len(article_urls) < max_articles:
            search_url = f'https://prtimes.jp/main/html/searchrlp/search_result.html?search_word={keyword}&page={page}'
            
            try:
                response = self.session.get(search_url)
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # 記事リンクを探す（複数のパターンに対応）
                # パターン1: <a class="link-thumbnail">
                links = soup.find_all('a', class_='link-thumbnail')
                
                # パターン2: <h3>タグ内のリンク
                if not links:
                    h3_tags = soup.find_all('h3')
                    links = []
                    for h3 in h3_tags:
                        link = h3.find('a')
                        if link:
                            links.append(link)
                
                if not links:
                    logger.warning(f"ページ {page} で記事が見つかりませんでした")
                    break
                
                for link in links:
                    href = link.get('href')
                    if href:
                        # 相対URLを絶対URLに変換
                        if href.startswith('/'):
                            href = f'https://prtimes.jp{href}'
                        article_urls.append(href)
                        
                        if len(article_urls) >= max_articles:
                            break
                
                logger.info(f"ページ {page} から {len(links)} 件の記事を取得")
                page += 1
                
                # ページネーションの終了チェック
                if page > 3:  # 最大3ページまで
                    break
                    
                time.sleep(1)  # サーバーへの負荷を考慮
                
            except Exception as e:
                logger.error(f"記事の取得中にエラーが発生しました: {e}")
                break
        
        return article_urls[:max_articles]
    
    def extract_info(self, article_url: str) -> Dict[str, str]:
        """
        記事ページからメディア関係者限定情報を抽出
        
        Args:
            article_url: 記事のURL
            
        Returns:
            Dict[str, str]: 抽出した情報
        """
        info = {
            '記事URL': article_url,
            '会社名': '',
            '担当者名': '',
            'メールアドレス': '',
            '電話番号': ''
        }
        
        try:
            response = self.session.get(article_url)
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # メディア関係者限定情報セクションを探す
            media_section = None
            
            # パターン1: class="media-info"
            media_section = soup.find('div', class_='media-info')
            
            # パターン2: "メディア関係者" を含むセクション
            if not media_section:
                for div in soup.find_all('div'):
                    if 'メディア関係者' in str(div):
                        media_section = div
                        break
            
            # 会社名の抽出
            company_pattern = r'(?:株式会社|有限会社|合同会社|合資会社|合名会社)[\w\s]+|[\w\s]+(?:株式会社|有限会社|合同会社|合資会社|合名会社)'
            company_match = re.search(company_pattern, response.text)
            if company_match:
                info['会社名'] = company_match.group().strip()
            
            # 担当者名の抽出（日本人の名前パターン）
            name_pattern = r'[一-龥]{1,4}[\s　]*[一-龥]{1,4}'
            if media_section:
                name_match = re.search(name_pattern, str(media_section))
                if name_match:
                    info['担当者名'] = name_match.group().strip()
            
            # メールアドレスの抽出
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            email_match = re.search(email_pattern, response.text)
            if email_match:
                info['メールアドレス'] = email_match.group()
            
            # 電話番号の抽出
            phone_pattern = r'(?:0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4})|(?:\d{2,4}-\d{2,4}-\d{4})'
            phone_match = re.search(phone_pattern, response.text)
            if phone_match:
                info['電話番号'] = phone_match.group()
            
            logger.info(f"記事から情報を抽出: {article_url}")
            
        except Exception as e:
            logger.error(f"情報抽出中にエラーが発生しました ({article_url}): {e}")
        
        return info
    
    def write_to_sheets(self, data: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
        """
        Google Sheetsにデータを書き込む
        
        Args:
            data: 書き込むデータフレーム
            spreadsheet_id: スプレッドシートのID
            sheet_name: シート名
        """
        try:
            # Google Sheets APIの認証
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_path, scope)
            client = gspread.authorize(creds)
            
            # スプレッドシートを開く
            spreadsheet = client.open_by_key(spreadsheet_id)
            
            # シートを取得または作成
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
            
            # 既存のデータをクリア
            worksheet.clear()
            
            # ヘッダーとデータを書き込む
            header = data.columns.tolist()
            values = data.values.tolist()
            
            # ヘッダーを書き込む
            worksheet.update('A1', [header])
            
            # データを書き込む
            if values:
                worksheet.update(f'A2:E{len(values)+1}', values)
            
            logger.info(f"Google Sheetsへの書き込みが完了しました: {len(values)}件")
            
        except Exception as e:
            logger.error(f"Google Sheetsへの書き込み中にエラーが発生しました: {e}")

def main():
    # 設定をconfig.pyから読み込む
    try:
        import config
        EMAIL = config.PRTIMES_EMAIL
        PASSWORD = config.PRTIMES_PASSWORD
        CREDENTIALS_PATH = config.GOOGLE_CREDENTIALS_PATH
        SPREADSHEET_ID = config.SPREADSHEET_ID
        SHEET_NAME = config.SHEET_NAME
        SEARCH_KEYWORD = config.DEFAULT_SEARCH_KEYWORD
    except ImportError:
        # config.pyが見つからない場合はデフォルト値を使用
        EMAIL = 'your_email@example.com'
        PASSWORD = 'your_password'
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_Data'
        SEARCH_KEYWORD = 'サプリ'
    
    # スクレイパーの初期化
    scraper = PRTimesScraper(EMAIL, PASSWORD, CREDENTIALS_PATH)
    
    # ログイン
    if not scraper.login():
        logger.error("ログインに失敗しました。認証情報を確認してください。")
        return
    
    # 記事URLの収集
    logger.info(f"キーワード '{SEARCH_KEYWORD}' で検索を開始します")
    article_urls = scraper.fetch_articles(SEARCH_KEYWORD, max_articles=50)
    logger.info(f"{len(article_urls)} 件の記事URLを収集しました")
    
    # 各記事から情報を抽出
    results = []
    for i, url in enumerate(article_urls, 1):
        logger.info(f"処理中: {i}/{len(article_urls)}")
        info = scraper.extract_info(url)
        results.append(info)
        time.sleep(1)  # サーバーへの負荷を考慮
    
    # データフレームに変換
    df = pd.DataFrame(results)
    
    # Google Sheetsに書き込み
    scraper.write_to_sheets(df, SPREADSHEET_ID, SHEET_NAME)
    
    # ローカルにも保存（バックアップ）
    df.to_csv('prtimes_data.csv', index=False, encoding='utf-8-sig')
    logger.info("処理が完了しました")

if __name__ == '__main__':
    main()