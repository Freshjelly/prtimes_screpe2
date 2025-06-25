#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PR TIMES自動ログイン・スクレイピング・Google Sheets連携スクリプト
"""

import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import logging
from typing import Dict, List, Optional, Tuple
import os
from urllib.parse import urljoin

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)
logger = logging.getLogger(__name__)


class PRTimesScreaper:
    def __init__(self, username: str, password: str):
        self.session = requests.Session()
        self.username = username
        self.password = password
        self.base_url = "https://prtimes.jp"
        self.login_url = "https://prtimes.jp/login"
        self.search_url = "https://prtimes.jp/main/html/searchrlp/search_result.html"
        
        # ヘッダー設定
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        }
        self.session.headers.update(self.headers)
    
    def login(self) -> bool:
        """PR TIMESにログインする"""
        try:
            # ログインページを取得してCSRFトークンを探す
            logger.info("ログインページにアクセス中...")
            login_page = self.session.get(self.login_url)
            login_page.raise_for_status()
            
            soup = BeautifulSoup(login_page.content, 'html.parser')
            
            # CSRFトークンを探す（複数のパターンで試行）
            csrf_token = None
            csrf_patterns = [
                ('input', {'name': 'csrf_token'}),
                ('input', {'name': '_csrf_token'}),
                ('input', {'name': 'authenticity_token'}),
                ('meta', {'name': 'csrf-token'})
            ]
            
            for tag, attrs in csrf_patterns:
                element = soup.find(tag, attrs)
                if element:
                    csrf_token = element.get('value') or element.get('content')
                    if csrf_token:
                        logger.info(f"CSRFトークンを取得: {csrf_token[:20]}...")
                        break
            
            # ログインフォームのアクション属性を取得
            login_form = soup.find('form', {'id': 'login-form'}) or soup.find('form', {'action': re.compile('login')})
            if login_form:
                action_url = login_form.get('action', self.login_url)
                if not action_url.startswith('http'):
                    action_url = urljoin(self.base_url, action_url)
            else:
                action_url = self.login_url
            
            # ログインデータの準備
            login_data = {
                'username': self.username,
                'password': self.password,
                'email': self.username,  # emailフィールドの場合もある
                'login_id': self.username,  # login_idフィールドの場合もある
            }
            
            # CSRFトークンがある場合は追加
            if csrf_token:
                login_data['csrf_token'] = csrf_token
                login_data['_csrf_token'] = csrf_token
                login_data['authenticity_token'] = csrf_token
            
            # ログインPOSTリクエスト
            logger.info("ログイン試行中...")
            response = self.session.post(
                action_url,
                data=login_data,
                allow_redirects=True
            )
            response.raise_for_status()
            
            # ログイン成功の確認
            if 'logout' in response.text.lower() or 'ログアウト' in response.text:
                logger.info("ログインに成功しました")
                return True
            else:
                logger.warning("ログインに失敗した可能性があります")
                return False
                
        except Exception as e:
            logger.error(f"ログインエラー: {str(e)}")
            return False
    
    def fetch_article_urls(self, keyword: str, max_urls: int = 50) -> List[str]:
        """検索結果から記事URLを収集する"""
        article_urls = []
        max_pages = 3
        
        try:
            for page in range(1, max_pages + 1):
                if len(article_urls) >= max_urls:
                    break
                    
                # 検索URLの構築
                params = {
                    'search_word': keyword,
                    'page': page
                }
                
                logger.info(f"検索ページ {page} を取得中...")
                response = self.session.get(self.search_url, params=params)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.content, 'html.parser')
                
                # 記事URLを抽出（複数のパターンで試行）
                url_patterns = [
                    r'/main/html/rd/p/\d+\.html',
                    r'href="(/main/html/rd/p/\d+\.html)"',
                    r'href=\'(/main/html/rd/p/\d+\.html)\''
                ]
                
                for pattern in url_patterns:
                    matches = re.findall(pattern, response.text)
                    if matches:
                        for match in matches:
                            if len(article_urls) < max_urls:
                                full_url = urljoin(self.base_url, match)
                                if full_url not in article_urls:
                                    article_urls.append(full_url)
                        break
                
                logger.info(f"ページ {page}: {len(article_urls)} 件の記事URLを収集")
                
                # レート制限対策
                time.sleep(1)
                
        except Exception as e:
            logger.error(f"記事URL収集エラー: {str(e)}")
        
        return article_urls[:max_urls]
    
    def extract_media_info(self, article_url: str) -> Dict[str, str]:
        """記事ページからメディア関係者情報を抽出する"""
        info = {
            'url': article_url,
            'company': '',
            'contact': '',
            'email': '',
            'phone': ''
        }
        
        try:
            logger.info(f"記事を取得中: {article_url}")
            response = self.session.get(article_url)
            response.raise_for_status()
            response.encoding = 'utf-8'
            
            soup = BeautifulSoup(response.content, 'html.parser')
            text = soup.get_text()
            
            # メディア関係者セクションを探す
            media_section = None
            media_keywords = ['メディア関係者', '報道関係者', 'お問い合わせ先', 'プレスリリース']
            
            for keyword in media_keywords:
                section = soup.find(text=re.compile(keyword))
                if section:
                    media_section = section.parent
                    while media_section and media_section.name not in ['div', 'section', 'article']:
                        media_section = media_section.parent
                    if media_section:
                        break
            
            # セクションが見つかった場合はその中から、なければ全体から検索
            search_text = media_section.get_text() if media_section else text
            
            # 会社名の抽出
            company_patterns = [
                r'(?:会社名|企業名|社名)[：:：\s]*([^\s\n\r、。；;]+(?:株式会社|有限会社|合同会社|Inc\.|LLC|Corporation))',
                r'(株式会社[^\s\n\r、。；;]{1,30})',
                r'([^\s\n\r、。；;]{1,30}株式会社)',
                r'((?:株式会社|有限会社|合同会社)[^\s\n\r、。；;]{1,30})'
            ]
            
            for pattern in company_patterns:
                match = re.search(pattern, search_text)
                if match:
                    info['company'] = match.group(1).strip().replace('\u3000', ' ')
                    break
            
            # 担当者名の抽出
            contact_patterns = [
                r'(?:担当者?|連絡先|広報|PR)[：:：\s]*([^\s\n\r、。；;]{2,10}(?:様|さん)?)',
                r'(?:お問い合わせ先?|問合せ先?)[：:：\s]*([^\s\n\r、。；;]{2,10})',
                r'(?:氏名|お名前)[：:：\s]*([^\s\n\r、。；;]{2,10})'
            ]
            
            for pattern in contact_patterns:
                match = re.search(pattern, search_text)
                if match:
                    contact = match.group(1).strip()
                    # 会社名や部署名を除外
                    if not any(word in contact for word in ['会社', '部', '課', '室']):
                        info['contact'] = contact.replace('様', '').replace('さん', '')
                        break
            
            # メールアドレスの抽出
            email_pattern = r'[a-zA-Z0-9][a-zA-Z0-9._%+-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,}'
            email_matches = re.findall(email_pattern, search_text)
            if email_matches:
                # PR関連のメールアドレスを優先
                pr_emails = [email for email in email_matches 
                           if any(keyword in email.lower() for keyword in ['pr', 'press', 'media', 'info'])]
                info['email'] = pr_emails[0] if pr_emails else email_matches[0]
            
            # 電話番号の抽出
            phone_patterns = [
                r'(?:TEL|Tel|tel|電話|℡)[：:：\s]*([0-9０-９]{2,4}[-－ー\s]?[0-9０-９]{2,4}[-－ー\s]?[0-9０-９]{3,4})',
                r'(0[0-9]{1,4}[-－ー\s]?[0-9]{1,4}[-－ー\s]?[0-9]{3,4})'
            ]
            
            for pattern in phone_patterns:
                match = re.search(pattern, search_text)
                if match:
                    # 全角数字を半角に変換
                    phone = match.group(1)
                    phone = re.sub(r'[０-９]', lambda x: chr(ord(x.group(0)) - 0xFEE0), phone)
                    phone = phone.replace('－', '-').replace('ー', '-').replace(' ', '')
                    if re.match(r'^\d{2,4}-?\d{2,4}-?\d{3,4}$', phone):
                        info['phone'] = phone
                        break
            
        except Exception as e:
            logger.error(f"情報抽出エラー ({article_url}): {str(e)}")
        
        return info
    
    def scrape(self, keyword: str, max_articles: int = 50) -> pd.DataFrame:
        """メイン処理: ログイン→URL収集→情報抽出"""
        # ログイン
        if not self.login():
            logger.error("ログインに失敗しました")
            return pd.DataFrame()
        
        # 記事URL収集
        article_urls = self.fetch_article_urls(keyword, max_articles)
        logger.info(f"{len(article_urls)} 件の記事URLを収集しました")
        
        # 各記事から情報抽出
        results = []
        for i, url in enumerate(article_urls, 1):
            logger.info(f"処理中: {i}/{len(article_urls)}")
            info = self.extract_media_info(url)
            results.append(info)
            
            # レート制限対策
            time.sleep(1.5)
        
        # DataFrameに変換
        df = pd.DataFrame(results)
        return df


class GoogleSheetsWriter:
    def __init__(self, credentials_file: str):
        self.credentials_file = credentials_file
        self.client = None
        self._authenticate()
    
    def _authenticate(self):
        """Google Sheetsの認証"""
        try:
            scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_file, scope)
            self.client = gspread.authorize(creds)
            logger.info("Google Sheets認証成功")
            
        except Exception as e:
            logger.error(f"Google Sheets認証エラー: {str(e)}")
            raise
    
    def write_to_sheet(self, df: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
        """DataFrameをGoogle Sheetsに書き込む"""
        try:
            # スプレッドシートを開く
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            
            # シートを取得（なければ作成）
            try:
                sheet = spreadsheet.worksheet(sheet_name)
                sheet.clear()  # 既存データをクリア
            except gspread.WorksheetNotFound:
                sheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
            
            # ヘッダーを書き込む
            headers = ['記事URL', '会社名', '担当者名', 'メールアドレス', '電話番号']
            sheet.update('A1:E1', [headers])
            
            # データを書き込む
            if not df.empty:
                # DataFrameをリストに変換
                data = df[['url', 'company', 'contact', 'email', 'phone']].values.tolist()
                
                # バッチで書き込む（効率化）
                if data:
                    cell_range = f'A2:E{len(data) + 1}'
                    sheet.update(cell_range, data)
            
            logger.info(f"{len(df)} 件のデータをGoogle Sheetsに書き込みました")
            
        except Exception as e:
            logger.error(f"Google Sheets書き込みエラー: {str(e)}")
            raise


def main():
    """メイン実行関数"""
    # 設定読み込み
    try:
        from config import (
            PRTIMES_USERNAME,
            PRTIMES_PASSWORD,
            GOOGLE_CREDENTIALS_FILE,
            SPREADSHEET_ID,
            SHEET_NAME,
            SEARCH_KEYWORD,
            MAX_ARTICLES
        )
    except ImportError:
        logger.error("config.pyが見つかりません。config.pyを作成してください。")
        return
    
    try:
        # PR TIMESスクレイパーの初期化と実行
        scraper = PRTimesScreaper(PRTIMES_USERNAME, PRTIMES_PASSWORD)
        df = scraper.scrape(SEARCH_KEYWORD, MAX_ARTICLES)
        
        if df.empty:
            logger.warning("データが取得できませんでした")
            return
        
        # 結果をCSVにも保存（バックアップ）
        csv_filename = f"prtimes_results_{SEARCH_KEYWORD}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        logger.info(f"結果をCSVファイルに保存: {csv_filename}")
        
        # Google Sheetsに書き込む
        if os.path.exists(GOOGLE_CREDENTIALS_FILE):
            writer = GoogleSheetsWriter(GOOGLE_CREDENTIALS_FILE)
            writer.write_to_sheet(df, SPREADSHEET_ID, SHEET_NAME)
        else:
            logger.warning(f"認証ファイルが見つかりません: {GOOGLE_CREDENTIALS_FILE}")
            logger.info("CSVファイルのみ保存されました")
        
    except Exception as e:
        logger.error(f"実行エラー: {str(e)}")
        raise


if __name__ == "__main__":
    main()