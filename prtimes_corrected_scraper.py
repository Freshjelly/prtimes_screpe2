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
import pandas as pd

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
        
        # 複数のパターンでログインURLを探す（2025年6月更新版）
        potential_urls = [
            'https://prtimes.jp/main/html/medialogin',  # メディアユーザー向けログイン（優先）
            'https://prtimes.jp/auth/login',  # 企業ユーザー向けログイン（2025年現在）
            'https://prtimes.jp/main/html/sociallogin',  # 個人ユーザー向けログイン
            # 以下は旧URLパターン（互換性のため残す）
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
                logger.debug(f"ログインURL候補をチェック中: {url}")
                response = self.session.get(url, timeout=10)
                logger.debug(f"  ステータスコード: {response.status_code}")
                
                if response.status_code == 200:
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # ログインフォームがあるかチェック
                    forms = soup.find_all('form')
                    logger.debug(f"  検出されたフォーム数: {len(forms)}")
                    
                    for i, form in enumerate(forms):
                        # より柔軟な入力フィールド検出
                        has_email = (form.find('input', {'type': 'email'}) or 
                                   form.find('input', {'name': re.compile('email', re.I)}) or
                                   form.find('input', {'name': re.compile('mail', re.I)}) or
                                   form.find('input', {'placeholder': re.compile('メール|email', re.I)}))
                        has_password = form.find('input', {'type': 'password'})
                        
                        # ユーザー名フィールドもチェック
                        has_username = (form.find('input', {'name': re.compile('user', re.I)}) or
                                      form.find('input', {'name': re.compile('login', re.I)}) or
                                      form.find('input', {'placeholder': re.compile('ユーザー|ID', re.I)}))
                        
                        logger.debug(f"  フォーム{i+1}: email={bool(has_email)}, password={bool(has_password)}, username={bool(has_username)}")
                        
                        # ログインフォームの条件を緩和
                        if has_password and (has_email or has_username):
                            logger.info(f"ログインフォームが見つかりました: {url}")
                            logger.debug(f"  フォームのaction属性: {form.get('action', 'なし')}")
                            return url
                else:
                    logger.debug(f"  アクセス失敗: HTTP {response.status_code}")
                            
            except Exception as e:
                logger.debug(f"URL {url} のチェック中にエラー: {e}")
                continue
        
        # トップページからログインリンクを探す
        try:
            logger.debug("トップページからログインリンクを探索中...")
            response = self.session.get('https://prtimes.jp/')
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # より詳細なログインリンクパターンを探す
            login_patterns = [
                r'ログイン|login',  # 一般的なログイン
                r'企業.*ログイン|corporate.*login',  # 企業ログイン
                r'メディア.*ログイン|media.*login',  # メディアログイン
                r'会員.*ログイン|member.*login',  # 会員ログイン
                r'配信.*ログイン|distribution.*login'  # 配信ログイン
            ]
            
            found_links = []
            for pattern in login_patterns:
                login_links = soup.find_all('a', string=re.compile(pattern, re.I))
                for link in login_links:
                    href = link.get('href', '')
                    link_text = link.get_text(strip=True)
                    if href:
                        if href.startswith('/'):
                            href = f'https://prtimes.jp{href}'
                        elif not href.startswith('http'):
                            href = f'https://prtimes.jp/{href}'
                        
                        found_links.append((href, link_text))
                        logger.debug(f"  発見されたリンク: '{link_text}' -> {href}")
            
            # 発見されたリンクをチェック
            for href, link_text in found_links:
                try:
                    logger.debug(f"リンクをチェック中: {href}")
                    test_response = self.session.get(href, timeout=10)
                    if test_response.status_code == 200:
                        test_soup = BeautifulSoup(test_response.text, 'html.parser')
                        if test_soup.find('input', {'type': 'password'}):
                            logger.info(f"ログインフォームが見つかりました: {href} ('{link_text}')")
                            return href
                except Exception as link_error:
                    logger.debug(f"  リンクチェック中にエラー: {link_error}")
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
                # メディアログインの場合は 'mail' フィールドを探す
                has_email = (form.find('input', {'type': 'email'}) or 
                           form.find('input', {'name': re.compile('email|mail', re.I)}) or
                           form.find('input', {'type': 'text', 'name': re.compile('mail', re.I)}))
                has_password = form.find('input', {'type': 'password'}) or form.find('input', {'name': 'pass'})
                
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
            # メディアユーザーログインの場合、フィールド名が異なる
            if 'medialogin' in login_url:
                login_data = {
                    'mail': self.email,  # メディアログインは 'mail' フィールド
                    'pass': self.password  # メディアログインは 'pass' フィールド
                }
            else:
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
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            'credentials.json', scope)
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
    
    # DataFrameに変換
    if results:
        df = pd.DataFrame(results)
        
        # ローカルにも保存
        df.to_csv('prtimes_corrected_data.csv', index=False, encoding='utf-8-sig')
        
        # Google Sheetsに書き込み
        SPREADSHEET_ID_NEW = "1ABCdefGHiJKlmNOPqrsTUvwxyz12345"  # 実際のスプレッドシートIDに変更してください
        SHEET_NAME_NEW = "Sheet1"
        write_to_google_sheets(df, SPREADSHEET_ID_NEW, SHEET_NAME_NEW)
    else:
        logger.warning("結果が空のため、DataFrameの作成をスキップします")
    
    logger.info(f"\n処理が完了しました。")
    logger.info(f"収集した記事数: {len(results)}件")
    
    email_count = sum(1 for r in results if r['メールアドレス'])
    phone_count = sum(1 for r in results if r['電話番号'])
    logger.info(f"メールアドレス取得数: {email_count}件")
    logger.info(f"電話番号取得数: {phone_count}件")

def test_extract_info():
    """動作検証用テスト関数"""
    # テスト用記事URLリスト
    test_urls = [
        'https://prtimes.jp/main/html/rd/p/000001554.000006302.html',
        'https://prtimes.jp/main/html/rd/p/000000520.000024045.html',
        'https://prtimes.jp/main/html/rd/p/000000156.000061950.html'
    ]
    
    # ダミー認証情報でスクレイパーを初期化
    scraper = PRTimesCorrectedScraper('test@example.com', 'test_password', 'dummy_credentials.json')
    
    print("=== extract_info() 動作検証 ===")
    results = []
    
    for i, url in enumerate(test_urls, 1):
        print(f"\n[テスト {i}] {url}")
        try:
            info = scraper.extract_info(url)
            print(f"  会社名: {info['会社名']}")
            print(f"  担当者名: {info['担当者名']}")
            print(f"  メールアドレス: {info['メールアドレス']}")
            print(f"  電話番号: {info['電話番号']}")
            results.append(info)
        except Exception as e:
            print(f"  エラー: {e}")
    
    # テストモードでもDataFrameがあれば書き込み
    if results:
        df = pd.DataFrame(results)
        print(f"\n=== DataFrame作成完了: {len(df)}件 ===")
        
        # Google Sheetsに書き込み（テスト用）
        SPREADSHEET_ID_TEST = "1ABCdefGHiJKlmNOPqrsTUvwxyz12345"  # テスト用スプレッドシートID
        SHEET_NAME_TEST = "TestSheet"
        
        print("Google Sheetsへの書き込みを実行中...")
        write_to_google_sheets(df, SPREADSHEET_ID_TEST, SHEET_NAME_TEST)

if __name__ == '__main__':
    import sys
    
    # コマンドライン引数でテストモードを指定
    if len(sys.argv) > 1 and sys.argv[1] == '--test':
        # デバッグログレベルに設定
        logging.getLogger().setLevel(logging.DEBUG)
        test_extract_info()
    else:
        main()