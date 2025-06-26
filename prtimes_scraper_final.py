#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import logging
from typing import List, Dict
import csv
import urllib.parse

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PRTimesScraper:
    def __init__(self, credentials_path: str):
        """
        PR Timesスクレイパーの初期化
        
        Args:
            credentials_path: Google認証用JSONファイルのパス
        """
        self.credentials_path = credentials_path
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        })
    
    def search_articles(self, keyword: str, max_articles: int = 50) -> List[str]:
        """
        キーワードで検索し、記事URLを収集
        
        Args:
            keyword: 検索キーワード
            max_articles: 収集する最大記事数
            
        Returns:
            List[str]: 記事URLのリスト
        """
        article_urls = []
        encoded_keyword = urllib.parse.quote(keyword)
        
        # 最初のページから開始
        base_url = f'https://prtimes.jp/main/action.php?run=html&page=searchkey&search_word={encoded_keyword}'
        
        page = 0
        while len(article_urls) < max_articles and page < 5:  # 最大5ページまで
            if page == 0:
                url = base_url
            else:
                url = f'{base_url}&search_page={page}'
            
            try:
                logger.info(f"検索中: {url}")
                response = self.session.get(url)
                
                if response.status_code != 200:
                    logger.warning(f"ステータスコード {response.status_code}: {url}")
                    break
                
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # 記事リンクを探す（/main/html/rd/p/ を含むリンク）
                links = soup.select('a[href*="/main/html/rd/p/"]')
                
                if not links:
                    logger.info("これ以上記事が見つかりません")
                    break
                
                for link in links:
                    href = link.get('href', '')
                    if href:
                        # 相対URLを絶対URLに変換
                        if href.startswith('/'):
                            href = f'https://prtimes.jp{href}'
                        elif not href.startswith('http'):
                            href = f'https://prtimes.jp/{href}'
                        
                        if href not in article_urls:
                            article_urls.append(href)
                            
                            if len(article_urls) >= max_articles:
                                break
                
                logger.info(f"ページ {page + 1} から {len(links)} 件の記事を発見（累計: {len(article_urls)}件）")
                
                # 次のページへのリンクがあるか確認
                next_link = soup.find('a', string='次へ')
                if not next_link:
                    # または「次の40件」などのリンクを探す
                    next_patterns = ['次の', '次へ', 'Next']
                    found_next = False
                    for pattern in next_patterns:
                        next_link = soup.find('a', string=re.compile(pattern))
                        if next_link:
                            found_next = True
                            break
                    
                    if not found_next:
                        logger.info("次のページが見つかりません")
                        break
                
                page += 1
                time.sleep(1)  # サーバーへの負荷を考慮
                
            except Exception as e:
                logger.error(f"検索エラー: {e}")
                break
        
        logger.info(f"合計 {len(article_urls)} 件の記事URLを収集しました")
        return article_urls[:max_articles]
    
    def extract_info(self, article_url: str) -> Dict[str, str]:
        """
        記事ページから情報を抽出
        
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
            
            # 会社名の抽出
            # 1. 発行元企業の情報から
            company_elem = soup.find('div', {'class': 'release-company'})
            if not company_elem:
                company_elem = soup.find('a', {'class': 'link-to-company'})
            if not company_elem:
                # classにcompanyを含む要素を探す
                company_elem = soup.find(['div', 'span', 'a'], class_=re.compile('company'))
            
            if company_elem:
                info['会社名'] = company_elem.text.strip()
            
            # お問い合わせ情報を探す
            # 「お問い合わせ」「本件に関する」「プレスリリース詳細」などのセクションを探す
            contact_keywords = ['お問い合わせ', '問い合わせ', '連絡先', '本件に関する', 'プレスリリース詳細', 'Contact']
            
            contact_section = None
            for keyword in contact_keywords:
                # 見出しを探す
                heading = soup.find(['h2', 'h3', 'h4', 'div'], string=re.compile(keyword))
                if heading:
                    # 見出しの次の要素または親要素を取得
                    contact_section = heading.find_next_sibling() or heading.parent
                    break
            
            # プレスリリース詳細のテーブルを探す
            if not contact_section:
                tables = soup.find_all('table')
                for table in tables:
                    if 'お問い合わせ' in table.text or '担当' in table.text:
                        contact_section = table
                        break
            
            # 全体のテキストから情報を抽出
            full_text = response.text
            
            # 担当者名の抽出
            name_patterns = [
                r'(?:担当|広報)[\s:：]*([一-龥ぁ-んァ-ヶー]{2,5}[\s　]*[一-龥ぁ-んァ-ヶー]{2,5})',
                r'([一-龥]{1,4}[\s　]+[一-龥]{1,4})[\s　]*(?:まで|宛)',
                r'(?:氏名|お名前)[\s:：]*([一-龥ぁ-んァ-ヶー\s　]{2,10})'
            ]
            
            for pattern in name_patterns:
                matches = re.findall(pattern, full_text)
                if matches:
                    # 最も可能性の高い名前を選択（会社名を除外）
                    for match in matches:
                        name = match.strip()
                        if len(name) >= 2 and len(name) <= 10 and '会社' not in name and '株式' not in name:
                            info['担当者名'] = name
                            break
                    if info['担当者名']:
                        break
            
            # メールアドレスの抽出
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            email_matches = re.findall(email_pattern, full_text)
            if email_matches:
                # PRTimes関連のメールアドレスを除外
                for email in email_matches:
                    if 'prtimes' not in email.lower():
                        info['メールアドレス'] = email
                        break
            
            # 電話番号の抽出
            phone_patterns = [
                r'(?:TEL|Tel|tel|電話|℡)[\s:：]*([0-9０-９\-－\s\(\)]{10,20})',
                r'(0[0-9]{1,4}[\-－\s]?[0-9]{1,4}[\-－\s]?[0-9]{3,4})',
                r'(０[０-９]{1,4}[\-－\s]?[０-９]{1,4}[\-－\s]?[０-９]{3,4})'
            ]
            
            for pattern in phone_patterns:
                phone_matches = re.findall(pattern, full_text)
                if phone_matches:
                    phone = phone_matches[0]
                    if isinstance(phone, tuple):
                        phone = phone[0]
                    # 全角数字を半角に変換
                    trans_table = str.maketrans('０１２３４５６７８９－（）　', '0123456789-() ')
                    phone = phone.translate(trans_table).strip()
                    # 不要なスペースを削除
                    phone = re.sub(r'\s+', '', phone)
                    if len(phone) >= 10:  # 電話番号として妥当な長さ
                        info['電話番号'] = phone
                        break
            
            logger.info(f"記事から情報を抽出: {article_url}")
            
        except Exception as e:
            logger.error(f"情報抽出中にエラーが発生しました ({article_url}): {e}")
        
        return info
    
    def write_to_sheets(self, data: List[Dict[str, str]], spreadsheet_id: str, sheet_name: str):
        """
        Google Sheetsにデータを書き込む
        
        Args:
            data: 書き込むデータのリスト
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
            if data:
                header = list(data[0].keys())
                values = []
                for row in data:
                    values.append([row.get(h, '') for h in header])
                
                # ヘッダーを書き込む
                worksheet.update('A1', [header])
                
                # データを書き込む
                if values:
                    worksheet.update(f'A2:E{len(values)+1}', values)
            
            logger.info(f"Google Sheetsへの書き込みが完了しました: {len(data)}件")
            
        except Exception as e:
            logger.error(f"Google Sheetsへの書き込み中にエラーが発生しました: {e}")

def main():
    # 設定をconfig.pyから読み込む
    try:
        import config
        CREDENTIALS_PATH = config.GOOGLE_CREDENTIALS_PATH
        SPREADSHEET_ID = config.SPREADSHEET_ID
        SHEET_NAME = config.SHEET_NAME
        SEARCH_KEYWORD = config.DEFAULT_SEARCH_KEYWORD
    except ImportError:
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_Data'
        SEARCH_KEYWORD = 'サプリ'
    
    # スクレイパーの初期化
    scraper = PRTimesScraper(CREDENTIALS_PATH)
    
    # 記事URLの収集
    logger.info(f"キーワード '{SEARCH_KEYWORD}' で検索を開始します")
    article_urls = scraper.search_articles(SEARCH_KEYWORD, max_articles=50)
    
    if not article_urls:
        logger.error("記事が見つかりませんでした")
        return
    
    # 各記事から情報を抽出
    results = []
    for i, url in enumerate(article_urls, 1):
        logger.info(f"処理中: {i}/{len(article_urls)}")
        info = scraper.extract_info(url)
        results.append(info)
        
        # 結果をリアルタイムで表示（最初の5件）
        if i <= 5:
            logger.info(f"  会社名: {info['会社名']}")
            if info['メールアドレス']:
                logger.info(f"  Email: {info['メールアドレス']}")
            if info['電話番号']:
                logger.info(f"  TEL: {info['電話番号']}")
        
        # 5件ごとに少し長めの待機
        if i % 5 == 0:
            time.sleep(2)
        else:
            time.sleep(0.5)
    
    # Google Sheetsに書き込み
    scraper.write_to_sheets(results, SPREADSHEET_ID, SHEET_NAME)
    
    # ローカルにも保存（バックアップ）
    with open('prtimes_data.csv', 'w', newline='', encoding='utf-8-sig') as f:
        if results:
            fieldnames = list(results[0].keys())
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
    
    logger.info(f"\n処理が完了しました。")
    logger.info(f"収集した記事数: {len(results)}件")
    
    # 統計情報
    email_count = sum(1 for r in results if r['メールアドレス'])
    phone_count = sum(1 for r in results if r['電話番号'])
    logger.info(f"メールアドレス取得数: {email_count}件")
    logger.info(f"電話番号取得数: {phone_count}件")

if __name__ == '__main__':
    main()