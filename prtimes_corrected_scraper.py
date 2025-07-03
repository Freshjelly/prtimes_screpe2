#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import gspread
from google.oauth2.service_account import Credentials
import time
import logging
from typing import List, Dict, Optional
import csv
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import requests.utils

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PRTimesCorrectedScraper:
    def __init__(self, email: str, password: str, credentials_path: str, headless: bool = True):
        """
        PR Timesスクレイパーの修正版
        
        Args:
            email: PR Timesログイン用メールアドレス
            password: PR Timesログイン用パスワード
            credentials_path: Google認証用JSONファイルのパス
            headless: ヘッドレスモードで実行するか（デフォルト: True）
        """
        self.email = email
        self.password = password
        self.credentials_path = credentials_path
        self.headless = headless
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
    
    
    
    def login(self) -> bool:
        """
        PR TimesにSeleniumを使ってログイン
        
        Returns:
            bool: ログイン成功時True、失敗時False
        """
        driver = None
        try:
            # Chrome オプション設定
            chrome_options = Options()
            if self.headless:
                chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--disable-software-rasterizer')
            chrome_options.add_argument('--remote-debugging-port=9222')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            
            # Chromiumのバイナリパスを明示的に指定
            chrome_options.binary_location = '/snap/bin/chromium'
            
            # ChromeDriver 自動管理 - Chromium用の設定
            from webdriver_manager.chrome import ChromeDriverManager
            from webdriver_manager.core.driver_cache import DriverCacheManager
            from webdriver_manager.core.os_manager import ChromeType
            
            # Chromium用のドライバーをインストール
            service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.implicitly_wait(10)
            
            # メディアユーザーログインURLにアクセス
            login_url = 'https://prtimes.jp/main/html/medialogin'
            logger.info(f"Seleniumでログインページにアクセス: {login_url}")
            driver.get(login_url)
            
            # ページ読み込み待機
            wait = WebDriverWait(driver, 20)
            
            # メールアドレス入力フィールドを探す
            email_field = wait.until(
                EC.presence_of_element_located((By.NAME, "mail"))
            )
            email_field.clear()
            email_field.send_keys(self.email)
            logger.info("メールアドレスを入力しました")
            
            # パスワード入力フィールドを探す
            password_field = driver.find_element(By.NAME, "pass")
            password_field.clear()
            password_field.send_keys(self.password)
            logger.info("パスワードを入力しました")
            
            # ログインボタンを探してクリック
            # 複数のセレクタを試す
            login_button = None
            button_selectors = [
                (By.XPATH, "//button[contains(text(), 'ログイン')]"),
                (By.XPATH, "//input[@type='submit' and (@value='ログイン' or @value='Login')]"),
                (By.XPATH, "//button[@type='submit']"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']")
            ]
            
            for selector_type, selector_value in button_selectors:
                try:
                    login_button = driver.find_element(selector_type, selector_value)
                    if login_button:
                        break
                except:
                    continue
            
            if not login_button:
                logger.error("ログインボタンが見つかりませんでした")
                return False
            
            # スクリーンショット（デバッグ用）
            if not self.headless:
                driver.save_screenshot("before_login.png")
            
            login_button.click()
            logger.info("ログインボタンをクリックしました")
            
            # ログイン処理の待機（ページ遷移またはエラーメッセージ）
            time.sleep(3)
            
            # ログイン成功の確認
            current_url = driver.current_url
            page_source = driver.page_source.lower()
            
            # 成功判定
            success_indicators = ['logout', 'ログアウト', 'mypage', 'マイページ', 'dashboard', 'ダッシュボード']
            failure_indicators = ['error', 'エラー', 'invalid', '無効', 'incorrect', '間違い', 'ログインできませんでした']
            
            login_success = False
            if any(indicator in page_source for indicator in success_indicators):
                logger.info("ログインに成功しました")
                login_success = True
            elif any(indicator in page_source for indicator in failure_indicators):
                logger.error("ログインに失敗しました（認証エラー）")
                if not self.headless:
                    driver.save_screenshot("login_error.png")
                return False
            elif current_url != login_url:
                # URLが変わっていればログイン成功の可能性が高い
                logger.info(f"ログイン後のURL: {current_url}")
                login_success = True
            else:
                logger.warning("ログイン結果の判定ができませんでした")
                if not self.headless:
                    driver.save_screenshot("login_unknown.png")
            
            if login_success:
                # Cookieを取得してrequests.Sessionにコピー
                logger.info("CookieをSeleniumからrequestsにコピー中...")
                selenium_cookies = driver.get_cookies()
                
                for cookie in selenium_cookies:
                    # requests用のCookie形式に変換
                    cookie_obj = {
                        'name': cookie['name'],
                        'value': cookie['value'],
                        'domain': cookie.get('domain', ''),
                        'path': cookie.get('path', '/'),
                        'secure': cookie.get('secure', False),
                        'expires': cookie.get('expiry', None)
                    }
                    
                    # requestsのセッションにCookieを追加
                    self.session.cookies.set(
                        cookie_obj['name'],
                        cookie_obj['value'],
                        domain=cookie_obj['domain'],
                        path=cookie_obj['path'],
                        secure=cookie_obj['secure']
                    )
                
                logger.info(f"{len(selenium_cookies)}個のCookieをコピーしました")
                self.logged_in = True
                
                # ログイン確認のためマイページにアクセス
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
            
            return False
                
        except Exception as e:
            logger.error(f"Seleniumログイン中にエラーが発生しました: {e}")
            if not self.headless and driver:
                driver.save_screenshot("login_exception.png")
            return False
        finally:
            if driver:
                driver.quit()
                logger.info("Seleniumドライバーを終了しました")
    
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
            # デバッグログ: 処理対象URL
            logger.debug(f"処理対象URL: {article_url}")
            
            response = self.session.get(article_url)  # セッション維持
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 会社名の抽出（既存ロジックを維持）
            company_elem = soup.find('div', {'class': 'release-company'})
            if not company_elem:
                company_elem = soup.find('a', {'class': 'link-to-company'})
            if not company_elem:
                company_elem = soup.find(['div', 'span', 'a'], class_=re.compile('company'))
            
            if company_elem:
                info['会社名'] = company_elem.text.strip()
            
            # 記事本文エリアを特定（フッター・ヘッダーを除外）
            main_content = None
            content_selectors = [
                'main', 'article', '.content', '.main-content', 
                '.article-content', '.release-content', '.press-release'
            ]
            
            for selector in content_selectors:
                main_content = soup.select_one(selector)
                if main_content:
                    logger.debug(f"記事本文エリア特定: {selector}")
                    break
            
            # 本文エリアが見つからない場合は、フッター・ヘッダーを除外したbody
            if not main_content:
                main_content = soup.find('body')
                if main_content:
                    # フッター・ヘッダーを除外
                    for exclude_tag in main_content.find_all(['header', 'footer', 'nav']):
                        exclude_tag.decompose()
                    # PR TIMESの共通要素を除外
                    for exclude_class in main_content.find_all(class_=re.compile(r'header|footer|nav|menu|sidebar')):
                        exclude_class.decompose()
                    logger.debug("記事本文エリア: body（フッター・ヘッダー除外後）")
            
            # メディア関係者限定セクションを優先的に探す（本文エリア内で）
            section = None
            media_keywords = [
                'メディア関係者限定', '本件に関するお問い合わせ', 'プレスリリースに関するお問い合わせ',
                '報道関係者お問い合わせ先', '広報担当', 'PR担当', '取材依頼', '問い合わせ先'
            ]
            
            search_area = main_content if main_content else soup
            
            for keyword in media_keywords:
                # キーワードを含む要素を探す（本文エリア内で）
                elements = search_area.find_all(string=re.compile(keyword))
                for element in elements:
                    if element.parent:
                        # 親要素を遡って適切なセクションを見つける
                        parent = element.parent
                        while parent and parent.name not in ['div', 'section', 'p', 'td']:
                            parent = parent.parent
                        if parent:
                            # PR TIMESの共通フッターでないことを確認
                            parent_text = parent.get_text(strip=True)
                            if not ('Copyright' in parent_text and 'PR TIMES' in parent_text):
                                section = parent
                                logger.debug(f"対象セクション発見: {keyword}")
                                break
                if section:
                    break
            
            # デバッグログ: 抽出対象セクションのHTML
            if section:
                section_html = str(section)[:500]  # 最初の500文字
                logger.debug(f"対象セクションHTML: {section_html}...")
            else:
                logger.debug("対象セクションが見つからず、全文を使用")
            
            # セクションが見つかった場合はその範囲内のテキストを使用
            if section:
                text = section.get_text(separator=' ', strip=True)
            else:
                # フォールバック1: 記事本文エリアから抽出
                if main_content:
                    text = main_content.get_text(separator=' ', strip=True)
                    logger.debug("フォールバック1: 記事本文エリア全体を使用")
                else:
                    # フォールバック2: 全文テキストを使用（最後の手段）
                    text = soup.get_text(separator=' ', strip=True)
                    logger.debug("フォールバック2: 全文テキストを使用")
            
            # デバッグログ: 抜き出したテキストの詳細情報
            logger.debug(f"抽出対象テキスト（冒頭100文字）: {text[:100]}")
            logger.debug(f"抽出対象テキスト長: {len(text)}文字")
            
            # 主要キーワードの有無をチェック
            keywords_check = {
                'お問い合わせ': 'お問い合わせ' in text,
                '担当': '担当' in text,
                'TEL': 'TEL' in text or '電話' in text,
                '@': '@' in text
            }
            logger.debug(f"キーワード存在チェック: {keywords_check}")
            
            # ピンポイント正規表現による抽出
            
            # 会社名の抽出（HTML要素から取得できなかった場合）
            if not info['会社名']:
                company_match = re.search(r'(株式会社[^\s、。；;]{1,30})', text)
                if company_match:
                    info['会社名'] = company_match.group(1)
                    logger.debug(f"会社名をテキストから抽出: {info['会社名']}")
            
            # 担当者名の抽出（複数パターン）
            person_patterns = [
                r'(?:担当者?[:：]\s*)([一-龥]{2,10})',
                r'(?:広報担当[:：]\s*)([一-龥]{2,10})',
                r'(?:PR担当[:：]\s*)([一-龥]{2,10})',
                r'([一-龥]{2,4}\s+[一-龥]{2,4})(?:\s*(?:まで|宛|様|氏))',
                r'(?:連絡先[:：]\s*)([一-龥]{2,10})',
                r'(?:問い合わせ先[:：]\s*)([一-龥]{2,10})'
            ]
            
            for pattern in person_patterns:
                person_match = re.search(pattern, text)
                if person_match:
                    candidate = person_match.group(1).strip()
                    # 無効な候補を除外
                    if (candidate and len(candidate) >= 2 and len(candidate) <= 10 and
                        '会社' not in candidate and '株式' not in candidate and
                        '法人' not in candidate and '企業' not in candidate):
                        info['担当者名'] = candidate
                        logger.debug(f"担当者名を抽出: {info['担当者名']} (パターン: {pattern})")
                        break
            
            # メールアドレスの抽出（複数パターン）
            email_patterns = [
                r'([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})',
                r'(?:E-?mail[:：]\s*)?([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})',
                r'(?:メール[:：]\s*)?([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})'
            ]
            
            for pattern in email_patterns:
                email_matches = re.findall(pattern, text)
                for email in email_matches:
                    if ('prtimes' not in email.lower() and 'example' not in email.lower() and
                        'test' not in email.lower()):
                        info['メールアドレス'] = email
                        logger.debug(f"メールアドレスを抽出: {info['メールアドレス']}")
                        break
                if info['メールアドレス']:
                    break
            
            # 電話番号の抽出（複数パターン）
            phone_patterns = [
                r'(?:TEL|Tel|tel|電話|℡)[:：\s]*([0-9０-９\-－\(\)\s]{10,20})',
                r'(0\d{1,4}[-－]\d{1,4}[-－]\d{3,4})',
                r'(０[\d０-９]{1,4}[－-][\d０-９]{1,4}[－-][\d０-９]{3,4})',
                r'(\d{2,4}[-－]\d{2,4}[-－]\d{4})',
                r'(\d{10,11})'  # ハイフンなしの電話番号
            ]
            
            for pattern in phone_patterns:
                phone_matches = re.findall(pattern, text)
                for phone in phone_matches:
                    # 全角を半角に変換
                    phone_clean = phone.translate(str.maketrans('０１２３４５６７８９－（）　', '0123456789-() '))
                    phone_clean = re.sub(r'[\s\(\)]', '', phone_clean)
                    
                    # 電話番号として妥当かチェック
                    if (len(phone_clean) >= 10 and len(phone_clean) <= 15 and
                        phone_clean.startswith('0') and phone_clean.replace('-', '').isdigit()):
                        # ハイフンがない場合は追加
                        if '-' not in phone_clean and len(phone_clean) == 10:
                            phone_clean = phone_clean[:3] + '-' + phone_clean[3:6] + '-' + phone_clean[6:]
                        elif '-' not in phone_clean and len(phone_clean) == 11:
                            phone_clean = phone_clean[:3] + '-' + phone_clean[3:7] + '-' + phone_clean[7:]
                        
                        info['電話番号'] = phone_clean
                        logger.debug(f"電話番号を抽出: {info['電話番号']} (パターン: {pattern})")
                        break
                if info['電話番号']:
                    break
            
            # 抽出結果が不十分な場合の追加処理
            if not info['メールアドレス'] and not info['電話番号'] and not info['担当者名']:
                logger.debug("セクション抽出で結果が得られなかったため、全文検索を実行")
                full_text = soup.get_text(separator=' ', strip=True)
                
                # 全文からメールアドレスを再検索
                if not info['メールアドレス']:
                    email_match = re.search(r'([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})', full_text)
                    if email_match:
                        email = email_match.group(1)
                        if 'prtimes' not in email.lower():
                            info['メールアドレス'] = email
                            logger.debug(f"全文検索でメールアドレスを抽出: {info['メールアドレス']}")
                
                # 全文から電話番号を再検索
                if not info['電話番号']:
                    phone_match = re.search(r'(?:TEL|Tel|tel|電話)[:：\s]*([0-9０-９\-－\(\)\s]{10,20})', full_text)
                    if phone_match:
                        phone = phone_match.group(1)
                        phone_clean = phone.translate(str.maketrans('０１２３４５６７８９－（）　', '0123456789-() '))
                        phone_clean = re.sub(r'[\s\(\)]', '', phone_clean)
                        if len(phone_clean) >= 10:
                            info['電話番号'] = phone_clean
                            logger.debug(f"全文検索で電話番号を抽出: {info['電話番号']}")
            
            # 抽出結果のサマリー
            extraction_summary = {
                '会社名': bool(info['会社名']),
                '担当者名': bool(info['担当者名']),
                'メールアドレス': bool(info['メールアドレス']),
                '電話番号': bool(info['電話番号'])
            }
            logger.debug(f"抽出結果サマリー: {extraction_summary}")
            
            if not any(extraction_summary.values()):
                logger.warning(f"連絡先情報が一切抽出できませんでした: {article_url}")
            
            logger.info(f"記事から情報を抽出: {article_url}")
            
        except Exception as e:
            logger.error(f"情報抽出中にエラーが発生しました ({article_url}): {e}")
        
        return info
    

def write_to_google_sheets(dataframe: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
    """
    DataFrameをGoogle Sheetsに書き込む
    
    Args:
        dataframe: 書き込むpandas DataFrame
        spreadsheet_id: Google SheetsのスプレッドシートID
        sheet_name: シート名
    """
    try:
        # Google Sheets認証
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(
            'credentials.json', scopes=scope)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        # シートを取得または作成
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
            logger.info(f"新しいシート '{sheet_name}' を作成しました")
        
        # シートの内容をクリア
        worksheet.clear()
        
        if not dataframe.empty:
            # ヘッダー（列名）を取得
            headers = dataframe.columns.tolist()
            
            # データを取得（リスト形式に変換）
            values = dataframe.values.tolist()
            
            # ヘッダーを書き込み
            worksheet.update('A1', [headers])
            
            # データを書き込み
            if values:
                # データの範囲を計算
                end_row = len(values) + 1
                end_col = chr(ord('A') + len(headers) - 1)
                range_name = f'A2:{end_col}{end_row}'
                
                worksheet.update(range_name, values)
            
            logger.info(f"Google Sheetsに {len(dataframe)} 件のデータを書き込みました")
        else:
            logger.warning("DataFrameが空のため、ヘッダーのみ書き込みます")
            
    except FileNotFoundError:
        logger.error("credentials.json ファイルが見つかりません")
    except Exception as e:
        logger.error(f"Google Sheetsへの書き込み中にエラーが発生しました: {e}")

def main(headless=True, search_keyword=None):
    # 設定をconfig.pyから読み込む
    try:
        import config
        EMAIL = config.PRTIMES_EMAIL
        PASSWORD = config.PRTIMES_PASSWORD
        CREDENTIALS_PATH = config.GOOGLE_CREDENTIALS_PATH
        SPREADSHEET_ID = config.SPREADSHEET_ID
        SHEET_NAME = config.SHEET_NAME
        SEARCH_KEYWORD = search_keyword or config.DEFAULT_SEARCH_KEYWORD
    except ImportError:
        EMAIL = 'your_email@example.com'
        PASSWORD = 'your_password'
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_Data'
        SEARCH_KEYWORD = search_keyword or 'サプリ'
    
    # スクレイパーの初期化（ヘッドレスモードを指定）
    scraper = PRTimesCorrectedScraper(EMAIL, PASSWORD, CREDENTIALS_PATH, headless=headless)
    
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
    
    # DataFrameに変換
    if results:
        df = pd.DataFrame(results)
        
        # ローカルにも保存
        df.to_csv('prtimes_corrected_data.csv', index=False, encoding='utf-8-sig')
        
        # Google Sheetsに書き込み
        write_to_google_sheets(df, SPREADSHEET_ID, SHEET_NAME)
    else:
        logger.warning("結果が空のため、DataFrameの作成をスキップします")
    
    logger.info(f"\n処理が完了しました。")
    logger.info(f"収集した記事数: {len(results)}件")
    
    email_count = sum(1 for r in results if r['メールアドレス'])
    phone_count = sum(1 for r in results if r['電話番号'])
    logger.info(f"メールアドレス取得数: {email_count}件")
    logger.info(f"電話番号取得数: {phone_count}件")


if __name__ == '__main__':
    import argparse
    
    # コマンドライン引数の処理
    parser = argparse.ArgumentParser(description='PR Times スクレイパー')
    parser.add_argument('--keyword', '-k', type=str, help='検索キーワード（例: --keyword "美容"）')
    parser.add_argument('--no-headless', action='store_true', help='ブラウザを表示して実行')
    
    args = parser.parse_args()
    
    # 実行
    headless_mode = not args.no_headless
    main(headless=headless_mode, search_keyword=args.keyword)