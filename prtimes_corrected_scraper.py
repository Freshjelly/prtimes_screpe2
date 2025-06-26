#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import logging
from typing import List, Dict, Optional
import csv
import urllib.parse

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PRTimesCorrectedScraper:
    def __init__(self, email: str, password: str, credentials_path: str):
        """
        PR Timesスクレイパーの修正版
        
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
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        })
        self.logged_in = False
    
    def find_login_url(self) -> Optional[str]:
        """
        正確なログインURLを動的に探索
        
        Returns:
            Optional[str]: 見つかったログインURL、見つからない場合はNone
        """
        logger.info("ログインURLを探索中...")
        
        # 複数のパターンでログインURLを探す
        potential_urls = [
            'https://media.prtimes.jp/login',
            'https://account.prtimes.jp/login',
            'https://login.prtimes.jp/',
            'https://prtimes.jp/main/action.php?run=html&page=login',
            'https://prtimes.jp/main/action.php?run=html&page=corp_login',
            'https://prtimes.jp/login',
            'https://prtimes.jp/main/login'
        ]
        
        for url in potential_urls:
            try:
                response = self.session.get(url, timeout=10)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # ログインフォームがあるかチェック
                    forms = soup.find_all('form')
                    for form in forms:
                        has_email = form.find('input', {'type': 'email'}) or form.find('input', {'name': re.compile('email', re.I)})
                        has_password = form.find('input', {'type': 'password'})
                        
                        if has_email and has_password:
                            logger.info(f"ログインフォームが見つかりました: {url}")
                            return url
                            
            except Exception as e:
                logger.debug(f"URL {url} のチェック中にエラー: {e}")
                continue
        
        # トップページからログインリンクを探す
        try:
            response = self.session.get('https://prtimes.jp/')
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # ログインリンクを探す
            login_links = soup.find_all('a', string=re.compile(r'ログイン|login', re.I))
            for link in login_links:
                href = link.get('href', '')
                if href:
                    if href.startswith('/'):
                        href = f'https://prtimes.jp{href}'
                    elif not href.startswith('http'):
                        href = f'https://prtimes.jp/{href}'
                    
                    # このリンクをチェック
                    try:
                        test_response = self.session.get(href, timeout=10)
                        if test_response.status_code == 200:
                            test_soup = BeautifulSoup(test_response.text, 'html.parser')
                            if test_soup.find('input', {'type': 'password'}):
                                logger.info(f"ログインフォームが見つかりました: {href}")
                                return href
                    except:
                        continue
                        
        except Exception as e:
            logger.debug(f"トップページの探索中にエラー: {e}")
        
        return None
    
    def extract_csrf_token(self, soup: BeautifulSoup) -> Dict[str, str]:
        """
        HTMLからCSRFトークンを抽出
        
        Args:
            soup: BeautifulSoupオブジェクト
            
        Returns:
            Dict[str, str]: CSRFトークンのフィールド名と値
        """
        csrf_data = {}
        
        # 複数のパターンでCSRFトークンを探す
        csrf_patterns = [
            # input type="hidden"
            ('input', {'type': 'hidden', 'name': re.compile(r'csrf|token|_token', re.I)}),
            # meta tag
            ('meta', {'name': re.compile(r'csrf-token|_token', re.I)}),
            # input name with csrf/token
            ('input', {'name': re.compile(r'csrf|token|authenticity', re.I)})
        ]
        
        for tag_name, attrs in csrf_patterns:
            elements = soup.find_all(tag_name, attrs)
            for elem in elements:
                if tag_name == 'meta':
                    name = elem.get('name', '')
                    value = elem.get('content', '')
                    if value:
                        csrf_data[name] = value
                        logger.info(f"CSRFトークン発見 (meta): {name}")
                else:
                    name = elem.get('name', '')
                    value = elem.get('value', '')
                    if name and value:
                        csrf_data[name] = value
                        logger.info(f"CSRFトークン発見 (input): {name}")
        
        return csrf_data
    
    def login(self) -> bool:
        """
        PR Timesにログイン（修正版）
        
        Returns:
            bool: ログイン成功時True、失敗時False
        """
        try:
            # 1. ログインURLを探索
            login_url = self.find_login_url()
            if not login_url:
                logger.error("ログインURLが見つかりませんでした")
                # フォールバック: 検索機能のみで動作
                logger.info("ログインなしでスクレイピングを継続します")
                self.logged_in = False
                return True  # 検索は可能なのでTrueを返す
            
            # 2. ログインページにアクセスしてCSRFトークンを取得
            logger.info(f"ログインページにアクセス: {login_url}")
            response = self.session.get(login_url)
            response.encoding = 'utf-8'
            
            if response.status_code != 200:
                logger.error(f"ログインページへのアクセスに失敗: {response.status_code}")
                return False
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 3. ログインフォームを特定
            login_form = None
            forms = soup.find_all('form')
            
            for form in forms:
                has_email = form.find('input', {'type': 'email'}) or form.find('input', {'name': re.compile('email', re.I)})
                has_password = form.find('input', {'type': 'password'})
                
                if has_email and has_password:
                    login_form = form
                    break
            
            if not login_form:
                logger.error("ログインフォームが見つかりませんでした")
                return False
            
            # 4. CSRFトークンを抽出
            csrf_data = self.extract_csrf_token(soup)
            
            # 5. フォームアクションURLを決定
            action = login_form.get('action', '')
            if action:
                if action.startswith('/'):
                    action_url = f'https://prtimes.jp{action}'
                elif not action.startswith('http'):
                    action_url = f'{login_url.rsplit("/", 1)[0]}/{action}'
                else:
                    action_url = action
            else:
                action_url = login_url
            
            # 6. ログインデータを準備
            login_data = {
                'email': self.email,
                'password': self.password
            }
            
            # CSRFトークンを追加
            login_data.update(csrf_data)
            
            # 隠しフィールドも追加
            hidden_inputs = login_form.find_all('input', {'type': 'hidden'})
            for hidden in hidden_inputs:
                name = hidden.get('name', '')
                value = hidden.get('value', '')
                if name and name not in login_data:
                    login_data[name] = value
            
            logger.info(f"ログイン送信先: {action_url}")
            logger.info(f"送信データ: {[k for k in login_data.keys() if k != 'password']}")
            
            # 7. ログインPOST
            headers = {
                'Referer': login_url,
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = self.session.post(action_url, data=login_data, headers=headers, allow_redirects=True)
            
            # 8. ログイン成功の確認
            success_indicators = ['logout', 'ログアウト', 'mypage', 'マイページ', 'dashboard', 'ダッシュボード']
            failure_indicators = ['error', 'エラー', 'invalid', '無効', 'incorrect', '間違い']
            
            response_text = response.text.lower()
            
            # 成功判定
            if any(indicator in response_text for indicator in success_indicators):
                logger.info("ログインに成功しました")
                self.logged_in = True
                
                # 9. マイページ等でログイン確認
                mypage_urls = [
                    'https://prtimes.jp/mypage',
                    'https://prtimes.jp/main/mypage',
                    'https://prtimes.jp/main/action.php?run=html&page=mypage'
                ]
                
                for mypage_url in mypage_urls:
                    try:
                        test_response = self.session.get(mypage_url)
                        if test_response.status_code == 200 and 'logout' in test_response.text.lower():
                            logger.info(f"ログイン確認完了: {mypage_url}")
                            return True
                    except:
                        continue
                
                return True
            
            # 失敗判定
            elif any(indicator in response_text for indicator in failure_indicators):
                logger.error("ログインに失敗しました（認証エラー）")
                return False
            
            # 判定できない場合
            else:
                logger.warning("ログイン結果の判定ができませんでした")
                logger.info(f"レスポンスURL: {response.url}")
                logger.info(f"ステータスコード: {response.status_code}")
                
                # レスポンスを保存
                with open('login_response.html', 'w', encoding='utf-8') as f:
                    f.write(response.text)
                logger.info("レスポンスを login_response.html に保存しました")
                
                return False
                
        except Exception as e:
            logger.error(f"ログイン中にエラーが発生しました: {e}")
            return False
    
    def search_articles(self, keyword: str, max_articles: int = 50) -> List[str]:
        """
        キーワードで検索し、記事URLを収集（requests.Session維持）
        
        Args:
            keyword: 検索キーワード
            max_articles: 収集する最大記事数
            
        Returns:
            List[str]: 記事URLのリスト
        """
        article_urls = []
        encoded_keyword = urllib.parse.quote(keyword)
        
        # セッションを維持したまま検索
        base_url = f'https://prtimes.jp/main/action.php?run=html&page=searchkey&search_word={encoded_keyword}'
        
        page = 0
        while len(article_urls) < max_articles and page < 5:
            if page == 0:
                url = base_url
            else:
                url = f'{base_url}&search_page={page}'
            
            try:
                logger.info(f"検索中: {url}")
                response = self.session.get(url)  # セッション維持
                
                if response.status_code != 200:
                    logger.warning(f"ステータスコード {response.status_code}: {url}")
                    break
                
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'html.parser')
                
                links = soup.select('a[href*="/main/html/rd/p/"]')
                
                if not links:
                    logger.info("これ以上記事が見つかりません")
                    break
                
                for link in links:
                    href = link.get('href', '')
                    if href:
                        if href.startswith('/'):
                            href = f'https://prtimes.jp{href}'
                        elif not href.startswith('http'):
                            href = f'https://prtimes.jp/{href}'
                        
                        if href not in article_urls:
                            article_urls.append(href)
                            
                            if len(article_urls) >= max_articles:
                                break
                
                logger.info(f"ページ {page + 1} から {len(links)} 件の記事を発見（累計: {len(article_urls)}件）")
                
                page += 1
                time.sleep(1)
                
            except Exception as e:
                logger.error(f"検索エラー: {e}")
                break
        
        logger.info(f"合計 {len(article_urls)} 件の記事URLを収集しました")
        return article_urls[:max_articles]
    
    def extract_info(self, article_url: str) -> Dict[str, str]:
        """
        記事ページから情報を抽出（セッション維持）
        
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
            response = self.session.get(article_url)  # セッション維持
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 会社名の抽出
            company_elem = soup.find('div', {'class': 'release-company'})
            if not company_elem:
                company_elem = soup.find('a', {'class': 'link-to-company'})
            if not company_elem:
                company_elem = soup.find(['div', 'span', 'a'], class_=re.compile('company'))
            
            if company_elem:
                info['会社名'] = company_elem.text.strip()
            
            # お問い合わせ情報の抽出
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
                    trans_table = str.maketrans('０１２３４５６７８９－（）　', '0123456789-() ')
                    phone = phone.translate(trans_table).strip()
                    phone = re.sub(r'\s+', '', phone)
                    if len(phone) >= 10:
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
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_path, scope)
            client = gspread.authorize(creds)
            
            spreadsheet = client.open_by_key(spreadsheet_id)
            
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
            
            worksheet.clear()
            
            if data:
                header = list(data[0].keys())
                values = []
                for row in data:
                    values.append([row.get(h, '') for h in header])
                
                worksheet.update('A1', [header])
                
                if values:
                    worksheet.update(f'A2:E{len(values)+1}', values)
            
            logger.info(f"Google Sheetsへの書き込みが完了しました: {len(data)}件")
            
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
        EMAIL = 'your_email@example.com'
        PASSWORD = 'your_password'
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_Data'
        SEARCH_KEYWORD = 'サプリ'
    
    # スクレイパーの初期化
    scraper = PRTimesCorrectedScraper(EMAIL, PASSWORD, CREDENTIALS_PATH)
    
    # ログイン試行
    if not scraper.login():
        logger.error("ログインに失敗しました。認証情報を確認してください。")
        logger.info("ログインなしで検索を続行します...")
    
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
        
        if i <= 5:
            logger.info(f"  会社名: {info['会社名']}")
            if info['メールアドレス']:
                logger.info(f"  Email: {info['メールアドレス']}")
            if info['電話番号']:
                logger.info(f"  TEL: {info['電話番号']}")
        
        if i % 5 == 0:
            time.sleep(2)
        else:
            time.sleep(0.5)
    
    # Google Sheetsに書き込み
    scraper.write_to_sheets(results, SPREADSHEET_ID, SHEET_NAME)
    
    # ローカルにも保存
    with open('prtimes_corrected_data.csv', 'w', newline='', encoding='utf-8-sig') as f:
        if results:
            fieldnames = list(results[0].keys())
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
    
    logger.info(f"\n処理が完了しました。")
    logger.info(f"収集した記事数: {len(results)}件")
    
    email_count = sum(1 for r in results if r['メールアドレス'])
    phone_count = sum(1 for r in results if r['電話番号'])
    logger.info(f"メールアドレス取得数: {email_count}件")
    logger.info(f"電話番号取得数: {phone_count}件")

if __name__ == '__main__':
    main()