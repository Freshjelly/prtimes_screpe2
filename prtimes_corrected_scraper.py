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
import unicodedata
from webdriver_manager.chrome import ChromeDriverManager
import requests.utils
import os
import subprocess
import platform
from datetime import datetime

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
            try:
                chromedriver_path = ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()
                logger.info(f"ChromeDriverパス: {chromedriver_path}")
                
                # 正しいパスが指定されているか確認
                if not chromedriver_path.endswith('chromedriver'):
                    # 正しいパスを構築
                    chromedriver_dir = os.path.dirname(chromedriver_path)
                    chromedriver_path = os.path.join(chromedriver_dir, 'chromedriver')
                    logger.info(f"修正されたChromeDriverパス: {chromedriver_path}")
                
                service = Service(chromedriver_path)
            except Exception as e:
                logger.error(f"ChromeDriverの初期化エラー: {e}")
                raise
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
    
    def search_articles(self, keyword: str, max_articles: int = 80) -> List[str]:
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
    
    def normalize_text(self, text: str) -> str:
        """テキストの正規化処理"""
        # Unicode正規化
        text = unicodedata.normalize('NFKC', text)
        # 余分な空白を削除
        text = re.sub(r'\s+', ' ', text)
        # 改行を空白に変換
        text = text.replace('\n', ' ').replace('\r', ' ')
        return text.strip()
    
    def decode_email(self, text: str) -> str:
        """難読化されたメールアドレスをデコード"""
        # メールアドレスの難読化パターン
        email_replacements = [
            (r'\[at\]', '@'),
            (r'\(at\)', '@'),
            (r'\[dot\]', '.'),
            (r'\(dot\)', '.'),
            (r'＠', '@'),
            (r'@', '@'),  # 全角アットマーク
            (r'．', '.'),  # 全角ピリオド
            (r'\s*at\s*', '@'),
            (r'\s*dot\s*', '.'),
            (r'&#64;', '@'),
            (r'&#46;', '.'),
        ]
        
        decoded = text
        for pattern, replacement in email_replacements:
            decoded = re.sub(pattern, replacement, decoded, flags=re.IGNORECASE)
        return decoded
    
    def normalize_phone(self, phone: str) -> str:
        """電話番号の正規化"""
        # 全角を半角に変換
        phone = phone.translate(str.maketrans('０１２３４５６７８９－（）　', '0123456789-() '))
        # 余分な文字を削除
        phone = re.sub(r'[^\d\-\(\)\+\s]', '', phone)
        # 空白を削除
        phone = re.sub(r'\s+', '', phone)
        # 括弧を処理
        phone = re.sub(r'\((\d+)\)', r'\1-', phone)
        # 複数のハイフンを一つに
        phone = re.sub(r'-+', '-', phone)
        # 先頭と末尾のハイフンを削除
        phone = phone.strip('-')
        
        # ハイフンがない10桁の電話番号にハイフンを追加
        if re.match(r'^\d{10}$', phone):
            if phone.startswith('03') or phone.startswith('06'):
                phone = f"{phone[:2]}-{phone[2:6]}-{phone[6:]}"
            else:
                phone = f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"
        elif re.match(r'^\d{11}$', phone):
            phone = f"{phone[:3]}-{phone[3:7]}-{phone[7:]}"
        
        return phone
    
    def extract_info(self, article_url: str, keyword: str = '') -> Dict[str, str]:
        """
        記事ページから情報を抽出（セッション維持）
        
        Args:
            article_url: 記事のURL
            keyword: 検索キーワード
            
        Returns:
            Dict[str, str]: 抽出した情報
        """
        info = {
            '記事URL': article_url,
            '検索キーワード': keyword,
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
            if not company_elem:
                # PR Times特有のセレクタを追加
                company_elem = soup.select_one('.content-header-sub-text, .release-header__company')
            
            if company_elem:
                info['会社名'] = company_elem.text.strip()
            else:
                # メタデータから会社名を取得
                meta_company = soup.find('meta', {'property': 'og:site_name'})
                if meta_company and meta_company.get('content'):
                    company_name = meta_company.get('content').replace('のプレスリリース', '').strip()
                    if company_name and company_name != 'PR TIMES':
                        info['会社名'] = company_name
            
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
                '報道関係者お問い合わせ先', '広報担当', 'PR担当', '取材依頼', '問い合わせ先',
                'お問合せ先', 'お問合わせ先', '問合せ先', '問合わせ先', '連絡先',
                'ご連絡先', 'Contact', 'CONTACT', '広報窓口', 'プレスお問い合わせ',
                '報道関係者', 'メディア問い合わせ', '取材申し込み', 'プレスコンタクト'
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
            
            # 問い合わせ先テーブルを探す（PR Times特有の構造）
            if not section:
                # テーブル形式の問い合わせ先情報を探す
                tables = search_area.find_all('table')
                for table in tables:
                    table_text = table.get_text()
                    if any(keyword in table_text for keyword in media_keywords):
                        section = table
                        logger.debug("問い合わせ先テーブルを発見")
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
            
            # 問い合わせ先の可能性が高い部分を抽出
            contact_section_patterns = [
                r'(?:【[^】]*(?:問い?合わ?せ|連絡先|広報)[^】]*】)([^【]+)',
                r'(?:■[^■\n]*(?:問い?合わ?せ|連絡先|広報)[^■\n]*)([^■]+)',
                r'(?:▼[^▼\n]*(?:問い?合わ?せ|連絡先|広報)[^▼\n]*)([^▼]+)',
                r'(?:●[^●\n]*(?:問い?合わ?せ|連絡先|広報)[^●\n]*)([^●]+)',
                r'(?:＜[^＞]*(?:問い?合わ?せ|連絡先|広報)[^＞]*＞)([^＜]+)'
            ]
            
            for pattern in contact_section_patterns:
                contact_match = re.search(pattern, text, re.DOTALL)
                if contact_match:
                    contact_text = contact_match.group(1)[:500]  # 最大500文字
                    logger.debug(f"問い合わせセクション検出: {contact_text[:100]}...")
                    # このセクションから優先的に情報を抽出
                    text = contact_text + ' ' + text  # 前に追加して優先度を上げる
            
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
                r'(?:担当者?[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:広報担当[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:PR担当[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'([一-龥]{2,4}[\s　]+[一-龥]{2,4})(?:\s*(?:まで|宛|様|氏))',
                r'(?:連絡先[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:問い合わせ先[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:お問い?合わ?せ先?[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:広報[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:ご担当[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})',
                r'(?:責任者[:：]\s*)([一-龥ぁ-んァ-ヶー]{2,10})'
            ]
            
            for pattern in person_patterns:
                person_match = re.search(pattern, text)
                if person_match:
                    candidate = person_match.group(1).strip()
                    # 無効な候補を除外
                    if (candidate and len(candidate) >= 2 and len(candidate) <= 10 and
                        '会社' not in candidate and '株式' not in candidate and
                        '法人' not in candidate and '企業' not in candidate and
                        '部' not in candidate and '課' not in candidate and
                        '室' not in candidate and 'チーム' not in candidate):
                        info['担当者名'] = candidate
                        logger.debug(f"担当者名を抽出: {info['担当者名']} (パターン: {pattern})")
                        break
            
            # メールアドレスの抽出（改良版）
            # まず難読化されたメールアドレスをデコード
            decoded_text = self.decode_email(text)
            
            # HTMLのテーブルやリストから構造的に抽出を試みる
            potential_emails = []
            
            # テーブル形式の情報を探す
            for table in soup.find_all('table'):
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    for i in range(len(cells) - 1):
                        cell_text = self.normalize_text(cells[i].get_text())
                        if any(keyword in cell_text for keyword in ['メール', 'Mail', 'Email', 'E-mail', 'e-mail']):
                            next_cell_text = self.normalize_text(cells[i + 1].get_text())
                            potential_emails.append(self.decode_email(next_cell_text))
            
            # リスト形式の情報を探す
            for li in soup.find_all('li'):
                li_text = self.normalize_text(li.get_text())
                if '@' in li_text or any(pattern in li_text for pattern in ['[at]', '(at)', '[dot]', '(dot)']):
                    potential_emails.append(self.decode_email(li_text))
            
            # メールアドレスの正規表現パターン
            email_patterns = [
                r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',
                r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Zぁ-ゔァ-ヴー]{2,}',  # 日本語ドメイン対応
            ]
            
            # 構造から抽出した候補を優先的にチェック
            for candidate in potential_emails:
                for pattern in email_patterns:
                    matches = re.findall(pattern, candidate)
                    for email in matches:
                        if 'prtimes' not in email.lower():
                            info['メールアドレス'] = email
                            logger.debug(f"メールアドレスを抽出（構造）: {info['メールアドレス']}")
                            break
                    if info['メールアドレス']:
                        break
                if info['メールアドレス']:
                    break
            
            # 見つからなければキーワード周辺を検索
            if not info['メールアドレス']:
                email_contexts = re.findall(
                    r'(?:メール|Mail|Email|E-mail|e-mail|連絡先)[^。\n]{0,50}',
                    decoded_text,
                    re.IGNORECASE
                )
                for context in email_contexts:
                    for pattern in email_patterns:
                        matches = re.findall(pattern, context)
                        for email in matches:
                            if 'prtimes' not in email.lower():
                                info['メールアドレス'] = email
                                logger.debug(f"メールアドレスを抽出（コンテキスト）: {info['メールアドレス']}")
                                break
                        if info['メールアドレス']:
                            break
                    if info['メールアドレス']:
                        break
            
            # それでも見つからなければ全文検索
            if not info['メールアドレス']:
                for pattern in email_patterns:
                    matches = re.findall(pattern, decoded_text)
                    for email in matches:
                        if 'prtimes' not in email.lower():
                            info['メールアドレス'] = email
                            logger.debug(f"メールアドレスを抽出（全文）: {info['メールアドレス']}")
                            break
                    if info['メールアドレス']:
                        break
            
            # 電話番号の抽出（改良版）
            # HTMLのテーブルやリストから構造的に抽出を試みる
            potential_phones = []
            
            # 電話番号のキーワードパターン
            phone_keywords = [
                r'TEL', r'Tel', r'tel', r'ＴＥＬ', r'Ｔｅｌ',
                r'電話', r'℡', r'☎', r'Phone', r'phone',
                r'電話番号', r'お電話', r'連絡先.*電話',
                r'T[:：]', r'Ｔ[:：]'
            ]
            
            # テーブル形式の情報を探す
            for table in soup.find_all('table'):
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    for i in range(len(cells) - 1):
                        cell_text = self.normalize_text(cells[i].get_text())
                        if any(re.search(keyword, cell_text, re.IGNORECASE) for keyword in phone_keywords):
                            next_cell_text = self.normalize_text(cells[i + 1].get_text())
                            potential_phones.append(next_cell_text)
            
            # リスト形式の情報を探す
            for li in soup.find_all('li'):
                li_text = self.normalize_text(li.get_text())
                if any(re.search(keyword, li_text, re.IGNORECASE) for keyword in phone_keywords):
                    potential_phones.append(li_text)
            
            # strongタグの後の電話番号を探す
            for strong in soup.find_all(['strong', 'b']):
                strong_text = self.normalize_text(strong.get_text())
                if any(re.search(keyword, strong_text, re.IGNORECASE) for keyword in phone_keywords):
                    next_text = strong.next_sibling
                    if next_text:
                        potential_phones.append(str(next_text).strip())
            
            # 電話番号の正規表現パターン
            phone_patterns = [
                # キーワード付きパターン
                r'(?:TEL|Tel|tel|電話|℡|ＴＥＬ)[:：\s]*([0-9０-９\-－\(\)\s（）]{10,20})',
                r'(?:T[:：]\s*)([0-9０-９\-－\(\)\s（）]{10,20})',
                r'(?:電話番号[:：]\s*)([0-9０-９\-－\(\)\s（）]{10,20})',
                # 標準的な日本の電話番号
                r'(0\d{1,4}[-－]\d{1,4}[-－]\d{3,4})',
                r'(０[\d０-９]{1,4}[－-][\d０-９]{1,4}[－-][\d０-９]{3,4})',
                # 括弧付き
                r'(\(0\d{1,4}\)\s*\d{1,4}[-－]?\d{3,4})',
                r'(（0[\d０-９]{1,4}）\s*[\d０-９]{1,4}[－-]?[\d０-９]{3,4})',
                # ハイフンなし
                r'(0\d{9,10})',
                r'(０[\d０-９]{9,10})',
                # 国際電話番号
                r'(\+81[-\s]?\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4})',
            ]
            
            # 構造から抽出した候補を優先的にチェック
            for candidate in potential_phones:
                for pattern in phone_patterns:
                    matches = re.findall(pattern, candidate)
                    for phone in matches:
                        normalized = self.normalize_phone(phone)
                        if len(normalized.replace('-', '')) >= 10 and normalized[0] in '0+':
                            info['電話番号'] = normalized
                            logger.debug(f"電話番号を抽出（構造）: {info['電話番号']}")
                            break
                    if info['電話番号']:
                        break
                if info['電話番号']:
                    break
            
            # 見つからなければキーワード周辺を検索
            if not info['電話番号']:
                for pattern in phone_patterns[:3]:  # キーワード付きパターンのみ
                    matches = re.findall(pattern, text)
                    for phone in matches:
                        normalized = self.normalize_phone(phone)
                        if len(normalized.replace('-', '')) >= 10 and normalized[0] in '0+':
                            info['電話番号'] = normalized
                            logger.debug(f"電話番号を抽出（キーワード）: {info['電話番号']}")
                            break
                    if info['電話番号']:
                        break
            
            # それでも見つからなければ全文検索（ただし慎重に）
            if not info['電話番号']:
                # 問い合わせセクション内のみ検索
                contact_section = re.search(
                    r'(?:問い?合わ?せ|連絡先|Contact|広報)[^。]*?([0-9０-９\-－\(\)\s（）]{10,20})',
                    text,
                    re.IGNORECASE | re.DOTALL
                )
                if contact_section:
                    phone_text = contact_section.group(1)
                    normalized = self.normalize_phone(phone_text)
                    if len(normalized.replace('-', '')) >= 10 and normalized[0] in '0+':
                        info['電話番号'] = normalized
                        logger.debug(f"電話番号を抽出（問い合わせセクション）: {info['電話番号']}")
            
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
    

def write_to_csv_with_pages(dataframe: pd.DataFrame, filename: str = None):
    """
    DataFrameをCSVファイルに書き込み、記事ごとにページを分けて出力
    
    Args:
        dataframe: 書き込むpandas DataFrame
        filename: 出力ファイル名（省略時はタイムスタンプ付きファイル名）
    
    Returns:
        str: 出力したファイルパス
    """
    try:
        # ファイル名の生成
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'prtimes_data_{timestamp}.csv'
        
        # CSVファイルに書き込み（記事ごとにページを分ける）
        with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            # ヘッダー行を書き込み
            writer.writerow(dataframe.columns.tolist())
            
            # 記事ごとにページを分けて書き込み
            for index, row in dataframe.iterrows():
                # 記事データを書き込み
                writer.writerow(row.tolist())
                
                # 記事の区切りとして改ページ文字を追加（最後の記事以外）
                if index < len(dataframe) - 1:
                    writer.writerow([''] * len(dataframe.columns))  # 空行
                    writer.writerow([f'--- Page Break (Article {index + 2}) ---'] + [''] * (len(dataframe.columns) - 1))
                    writer.writerow([''] * len(dataframe.columns))  # 空行
        
        logger.info(f"CSVファイルを作成しました: {filename}")
        return filename
        
    except Exception as e:
        logger.error(f"CSVファイルの作成中にエラーが発生しました: {e}")
        return None

def write_to_excel_with_keywords(keyword_data_dict: dict, filename: str = None):
    """
    キーワードごとのデータを別々のシートにExcelファイルとして保存
    各シート内では記事ごとにページ分割
    
    Args:
        keyword_data_dict: キーワードをキーとしてデータのリストを値とする辞書
        filename: 出力ファイル名（省略時はタイムスタンプ付きファイル名）
    
    Returns:
        str: 出力したファイルパス
    """
    try:
        # ファイル名の生成
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'prtimes_all_keywords_{timestamp}.xlsx'
        
        # Excelファイルに書き込み（xlsxwriterエンジンを使用）
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # ヘッダーのフォーマット設定
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center'
            })
            
            # ページ区切りのフォーマット設定
            page_break_format = workbook.add_format({
                'bold': True,
                'bg_color': '#FFE4B5',
                'border': 1,
                'align': 'center'
            })
            
            for keyword, data_list in keyword_data_dict.items():
                if not data_list:
                    continue
                
                # DataFrameを作成
                df = pd.DataFrame(data_list)
                
                # シート名をキーワードに設定（Excelシート名の制限に合わせて調整）
                sheet_name = keyword[:31]  # Excelシート名は31文字まで
                sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('?', '_')
                sheet_name = sheet_name.replace('*', '_').replace('[', '_').replace(']', '_')
                sheet_name = sheet_name.replace(':', '_').replace('<', '_').replace('>', '_')
                sheet_name = sheet_name.replace('|', '_')
                
                # データをシートに書き込み
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                
                # ワークシートを取得してフォーマット適用
                worksheet = writer.sheets[sheet_name]
                
                # ヘッダーにフォーマット適用
                for col_num, header in enumerate(df.columns):
                    worksheet.write(0, col_num, header, header_format)
                
                # 記事ごとにページ分割マーカーを追加
                current_row = len(df) + 2  # データの下に空行を一つ空けて開始
                
                for i in range(1, len(df)):
                    # 5行ごとに改ページマーカーを挿入
                    if i % 5 == 0:
                        worksheet.write(current_row, 0, f'--- Page Break (Next 5 Articles) ---', page_break_format)
                        current_row += 2  # 空行も追加
                
                # 列幅の自動調整
                for i, col in enumerate(df.columns):
                    # 列の最大文字数を取得
                    max_len = max(
                        df[col].astype(str).map(len).max() if len(df) > 0 else 10,
                        len(col)
                    ) + 2
                    # 列幅を設定（最大50文字まで）
                    worksheet.set_column(i, i, min(max_len, 50))
                
                logger.info(f"シート '{sheet_name}' を作成: {len(df)}件")
        
        logger.info(f"Excelファイルを作成しました: {filename}")
        
        # Excelファイルを自動で開く
        abs_path = os.path.abspath(filename)
        system = platform.system()
        
        try:
            if system == 'Windows':
                os.startfile(abs_path)
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', abs_path])
            elif system == 'Linux':
                # WSLの場合、Windows側のExcelで開く
                if 'microsoft' in platform.uname().release.lower():
                    # WSLパスをWindowsパスに変換
                    windows_path = subprocess.check_output(['wslpath', '-w', abs_path]).decode().strip()
                    subprocess.run(['cmd.exe', '/c', 'start', windows_path])
                else:
                    # 通常のLinux
                    subprocess.run(['xdg-open', abs_path])
            
            logger.info(f"Excelファイルを開きました: {abs_path}")
        except Exception as e:
            logger.warning(f"Excelファイルの自動起動に失敗しました: {e}")
            logger.info(f"手動でファイルを開いてください: {abs_path}")
        
        return abs_path
        
    except Exception as e:
        logger.error(f"Excelファイルの作成中にエラーが発生しました: {e}")
        return None

def write_to_excel_and_open(dataframe: pd.DataFrame, filename: str = None):
    """
    DataFrameをExcelファイルに書き込み、自動で開く
    
    Args:
        dataframe: 書き込むpandas DataFrame
        filename: 出力ファイル名（省略時はタイムスタンプ付きファイル名）
    
    Returns:
        str: 出力したファイルパス
    """
    try:
        # ファイル名の生成
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'prtimes_data_{timestamp}.xlsx'
        
        # Excelファイルに書き込み（xlsxwriterエンジンを使用）
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            dataframe.to_excel(writer, sheet_name='PR_Times_Data', index=False)
            
            # ワークシートの取得
            workbook = writer.book
            worksheet = writer.sheets['PR_Times_Data']
            
            # 列幅の自動調整
            for i, col in enumerate(dataframe.columns):
                # 列の最大文字数を取得
                max_len = max(
                    dataframe[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                # 列幅を設定（最大50文字まで）
                worksheet.set_column(i, i, min(max_len, 50))
            
            # ヘッダーのフォーマット設定
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            # ヘッダーに書式を適用
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        logger.info(f"Excelファイルを作成しました: {filename}")
        
        # Excelファイルを自動で開く
        abs_path = os.path.abspath(filename)
        system = platform.system()
        
        try:
            if system == 'Windows':
                os.startfile(abs_path)
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', abs_path])
            elif system == 'Linux':
                # WSLの場合、Windows側のExcelで開く
                if 'microsoft' in platform.uname().release.lower():
                    # WSLパスをWindowsパスに変換
                    windows_path = subprocess.check_output(['wslpath', '-w', abs_path]).decode().strip()
                    subprocess.run(['cmd.exe', '/c', 'start', windows_path])
                else:
                    # 通常のLinux
                    subprocess.run(['xdg-open', abs_path])
            
            logger.info(f"Excelファイルを開きました: {abs_path}")
        except Exception as e:
            logger.warning(f"Excelファイルの自動起動に失敗しました: {e}")
            logger.info(f"手動でファイルを開いてください: {abs_path}")
        
        return abs_path
        
    except Exception as e:
        logger.error(f"Excelファイルの作成中にエラーが発生しました: {e}")
        return None

def main(headless=True, search_keyword=None, use_multiple_keywords=False):
    # 設定をconfig.pyから読み込む
    try:
        import config
        EMAIL = config.PRTIMES_EMAIL
        PASSWORD = config.PRTIMES_PASSWORD
        CREDENTIALS_PATH = config.GOOGLE_CREDENTIALS_PATH
        SPREADSHEET_ID = config.SPREADSHEET_ID
        SHEET_NAME = config.SHEET_NAME
        if use_multiple_keywords:
            SEARCH_KEYWORDS = config.SEARCH_KEYWORDS
        else:
            SEARCH_KEYWORDS = [search_keyword or config.DEFAULT_SEARCH_KEYWORD]
    except ImportError:
        EMAIL = 'your_email@example.com'
        PASSWORD = 'your_password'
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_Data'
        SEARCH_KEYWORDS = [search_keyword or 'サプリ']
    
    # スクレイパーの初期化（ヘッドレスモードを指定）
    scraper = PRTimesCorrectedScraper(EMAIL, PASSWORD, CREDENTIALS_PATH, headless=headless)
    
    # ログイン試行
    if not scraper.login():
        logger.error("ログインに失敗しました。認証情報を確認してください。")
        logger.info("ログインなしで検索を続行します...")
    
    # 全結果を格納するリスト
    all_results = []
    # キーワードごとの結果を格納する辞書
    keyword_results_dict = {}
    
    # 各キーワードで検索
    for keyword_index, keyword in enumerate(SEARCH_KEYWORDS, 1):
        logger.info(f"\n{'='*50}")
        logger.info(f"キーワード {keyword_index}/{len(SEARCH_KEYWORDS)}: '{keyword}' で検索を開始します")
        logger.info(f"{'='*50}")
        
        # 記事URLの収集
        article_urls = scraper.search_articles(keyword, max_articles=100)
        
        if not article_urls:
            logger.warning(f"キーワード '{keyword}' では記事が見つかりませんでした")
            continue
        
        # 各記事から情報を抽出
        keyword_results = []
        for i, url in enumerate(article_urls, 1):
            logger.info(f"処理中: {i}/{len(article_urls)} - {keyword}")
            info = scraper.extract_info(url, keyword)
            keyword_results.append(info)
            all_results.append(info)
            
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
        
        # キーワード別の結果を辞書に保存
        keyword_results_dict[keyword] = keyword_results
        logger.info(f"キーワード '{keyword}' の結果: {len(keyword_results)}件")
        
        # キーワード間の待機時間
        if keyword_index < len(SEARCH_KEYWORDS):
            logger.info(f"次のキーワードまで5秒待機...")
            time.sleep(5)
    
    # Excel出力（キーワードごとにシート分け）
    if keyword_results_dict:
        # キーワードごとにシートを分けたExcelファイルを作成
        excel_path = write_to_excel_with_keywords(keyword_results_dict)
        
        # 全データを統合したDataFrameも作成（バックアップ用）
        if all_results:
            df = pd.DataFrame(all_results)
            # CSVファイルに記事ごとにページを分けて保存
            csv_path = write_to_csv_with_pages(df)
        
        # Google Sheetsへの書き込みはオプション（必要に応じて）
        # write_to_google_sheets(df, SPREADSHEET_ID, SHEET_NAME)
    else:
        logger.warning("結果が空のため、ファイルの作成をスキップします")
    
    logger.info(f"\n{'='*50}")
    logger.info(f"処理が完了しました。")
    logger.info(f"{'='*50}")
    logger.info(f"検索キーワード数: {len(SEARCH_KEYWORDS)}個")
    logger.info(f"収集した記事数: {len(all_results)}件")
    
    email_count = sum(1 for r in all_results if r['メールアドレス'])
    phone_count = sum(1 for r in all_results if r['電話番号'])
    logger.info(f"メールアドレス取得数: {email_count}件")
    logger.info(f"電話番号取得数: {phone_count}件")
    
    # キーワード別の集計
    logger.info(f"\nキーワード別集計:")
    for keyword in SEARCH_KEYWORDS:
        keyword_count = sum(1 for r in all_results if r.get('検索キーワード') == keyword)
        if keyword_count > 0:
            logger.info(f"  {keyword}: {keyword_count}件")



if __name__ == '__main__':
    import argparse
    
    # コマンドライン引数の処理
    parser = argparse.ArgumentParser(description='PR Times スクレイパー')
    parser.add_argument('--keyword', '-k', type=str, help='検索キーワード（例: --keyword "美容"）')
    parser.add_argument('--no-headless', action='store_true', help='ブラウザを表示して実行')
    parser.add_argument('--multiple', '-m', action='store_true', help='複数キーワードモード（config.pyのSEARCH_KEYWORDSを使用）')
    
    args = parser.parse_args()
    
    # 実行
    headless_mode = not args.no_headless
    main(headless=headless_mode, search_keyword=args.keyword, use_multiple_keywords=args.multiple)