#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PR TIMES スクレイパー V2.0
分析結果に基づく完全再設計版
"""

import requests
from bs4 import BeautifulSoup
import re
import json
import time
import logging
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass
from urllib.parse import quote, urljoin
import csv
from pathlib import Path
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('prtimes_scraper.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

@dataclass
class ContactInfo:
    """連絡先情報を格納するデータクラス"""
    article_url: str = ""
    company_name: str = ""
    contact_person: str = ""
    email: str = ""
    phone: str = ""
    confidence_score: float = 0.0
    extraction_source: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {
            '記事URL': self.article_url,
            '会社名': self.company_name,
            '担当者名': self.contact_person,
            'メールアドレス': self.email,
            '電話番号': self.phone,
            '信頼度': f"{self.confidence_score:.2f}",
            '抽出元': self.extraction_source
        }

class DataExtractor:
    """データ抽出を専門に行うクラス"""
    
    def __init__(self):
        self.email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
        self.phone_patterns = [
            re.compile(r'(?:TEL|Tel|tel|電話|℡)[\s:：]*([0-9０-９\-－\s\(\)]{10,20})'),
            re.compile(r'(?:^|\s)(0[0-9]{1,4}[\-－\s]?[0-9]{1,4}[\-－\s]?[0-9]{3,4})(?:\s|$)'),
            re.compile(r'(?:^|\s)(０[０-９]{1,4}[\-－\s]?[０-９]{1,4}[\-－\s]?[０-９]{3,4})(?:\s|$)')
        ]
        self.name_patterns = [
            re.compile(r'(?:担当|広報|PR|問い合わせ)[\s:：]*([一-龥ぁ-んァ-ヶー]{2,4}[\s　]*[一-龥ぁ-んァ-ヶー]{2,4})'),
            re.compile(r'(?:責任者|代表|manager)[\s:：]*([一-龥ぁ-んァ-ヶー]{2,8})'),
            re.compile(r'([一-龥]{2,4})[\s　]+([一-龥]{2,4})(?:まで|宛|様)')
        ]
    
    def extract_company_name(self, soup: BeautifulSoup, url: str) -> Tuple[str, float]:
        """会社名を抽出（信頼度付き）"""
        # パターン1: company classを持つaタグ（最高信頼度）
        company_links = soup.find_all('a', class_=re.compile('company', re.I))
        if company_links:
            company_name = company_links[0].text.strip()
            if company_name and '株式会社' in company_name or '会社' in company_name:
                return company_name, 0.95
        
        # パターン2: タイトルから抽出
        title = soup.find('title')
        if title:
            title_text = title.text
            # "| 会社名のプレスリリース" パターン
            title_match = re.search(r'\|\s*(.+?)のプレスリリース', title_text)
            if title_match:
                company_name = title_match.group(1).strip()
                return company_name, 0.85
        
        # パターン3: 構造化データから
        json_scripts = soup.find_all('script', {'type': 'application/ld+json'})
        for script in json_scripts:
            try:
                data = json.loads(script.string)
                if isinstance(data, list):
                    data = data[0]
                
                # author.name を探す
                if 'author' in data and isinstance(data['author'], dict):
                    author_name = data['author'].get('name', '')
                    if author_name and '会社' in author_name:
                        return author_name, 0.80
            except:
                continue
        
        return "", 0.0
    
    def extract_contact_info(self, soup: BeautifulSoup, text: str) -> Tuple[str, str, str, Dict[str, float]]:
        """連絡先情報を抽出"""
        email = ""
        phone = ""
        person = ""
        sources = {}
        
        # メールアドレス抽出
        email_matches = self.email_pattern.findall(text)
        for email_candidate in email_matches:
            if 'prtimes' not in email_candidate.lower():
                email = email_candidate
                sources['email'] = 0.90
                break
        
        # 電話番号抽出
        for i, pattern in enumerate(self.phone_patterns):
            phone_matches = pattern.findall(text)
            for match in phone_matches:
                if isinstance(match, tuple):
                    match = match[0] if match[0] else match[1]
                
                # 全角を半角に変換
                phone_clean = match.translate(str.maketrans('０１２３４５６７８９－（）　', '0123456789-() '))
                phone_clean = re.sub(r'\s+', '', phone_clean.strip())
                
                if len(phone_clean) >= 10:
                    phone = phone_clean
                    sources['phone'] = 0.95 - (i * 0.1)  # パターンの信頼度を調整
                    break
            
            if phone:
                break
        
        # 担当者名抽出
        for i, pattern in enumerate(self.name_patterns):
            name_matches = pattern.findall(text)
            for match in name_matches:
                if isinstance(match, tuple):
                    candidate = ' '.join(match).strip()
                else:
                    candidate = match.strip()
                
                # 有効性チェック
                if (2 <= len(candidate) <= 10 and 
                    '会社' not in candidate and 
                    '株式' not in candidate and
                    'メール' not in candidate):
                    person = candidate
                    sources['person'] = 0.70 - (i * 0.1)
                    break
            
            if person:
                break
        
        return email, phone, person, sources
    
    def calculate_confidence_score(self, info: ContactInfo, sources: Dict[str, float]) -> float:
        """抽出情報の総合信頼度を計算"""
        scores = []
        
        if info.company_name:
            scores.append(0.95)  # 会社名は必須で高信頼度
        
        if info.email and 'email' in sources:
            scores.append(sources['email'])
        
        if info.phone and 'phone' in sources:
            scores.append(sources['phone'])
        
        if info.contact_person and 'person' in sources:
            scores.append(sources['person'])
        
        return sum(scores) / len(scores) if scores else 0.0

class PRTimesSearcher:
    """PR TIMESの検索を専門に行うクラス"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        })
    
    def search_articles(self, keyword: str, max_articles: int = 50) -> List[str]:
        """キーワードで記事を検索してURLリストを返す"""
        article_urls = []
        encoded_keyword = quote(keyword)
        
        logger.info(f"キーワード '{keyword}' で検索開始（最大{max_articles}件）")
        
        page = 0
        while len(article_urls) < max_articles and page < 10:  # 最大10ページ
            if page == 0:
                url = f'https://prtimes.jp/main/action.php?run=html&page=searchkey&search_word={encoded_keyword}'
            else:
                url = f'https://prtimes.jp/main/action.php?run=html&page=searchkey&search_word={encoded_keyword}&search_page={page}'
            
            try:
                logger.debug(f"検索ページ {page + 1}: {url}")
                response = self.session.get(url, timeout=15)
                
                if response.status_code != 200:
                    logger.warning(f"検索失敗 (ステータス: {response.status_code}): {url}")
                    break
                
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # 記事リンクを抽出
                links = soup.select('a[href*="/main/html/rd/p/"]')
                
                if not links:
                    logger.info(f"ページ {page + 1} で記事が見つかりませんでした。検索終了。")
                    break
                
                page_urls = []
                for link in links:
                    href = link.get('href', '')
                    if href:
                        if href.startswith('/'):
                            href = f'https://prtimes.jp{href}'
                        elif not href.startswith('http'):
                            href = f'https://prtimes.jp/{href}'
                        
                        if href not in article_urls:
                            article_urls.append(href)
                            page_urls.append(href)
                            
                            if len(article_urls) >= max_articles:
                                break
                
                logger.info(f"ページ {page + 1}: {len(page_urls)}件の新しい記事を発見（累計: {len(article_urls)}件）")
                
                page += 1
                time.sleep(1)  # レート制限
                
            except Exception as e:
                logger.error(f"検索エラー (ページ {page + 1}): {e}")
                break
        
        logger.info(f"検索完了: 合計 {len(article_urls)} 件の記事URLを収集")
        return article_urls[:max_articles]

class GoogleSheetsConnector:
    """Google Sheets接続を専門に行うクラス"""
    
    def __init__(self, credentials_path: str):
        self.credentials_path = Path(credentials_path)
        self.client = None
        self._authenticate()
    
    def _authenticate(self):
        """Google Sheets認証"""
        try:
            if not self.credentials_path.exists():
                raise FileNotFoundError(f"認証ファイルが見つかりません: {self.credentials_path}")
            
            scope = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive'
            ]
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                str(self.credentials_path), scope
            )
            self.client = gspread.authorize(creds)
            logger.info("Google Sheets認証に成功しました")
            
        except Exception as e:
            logger.error(f"Google Sheets認証に失敗しました: {e}")
            self.client = None
    
    def write_data(self, data: List[ContactInfo], spreadsheet_id: str, sheet_name: str) -> bool:
        """データをGoogle Sheetsに書き込む"""
        if not self.client:
            logger.error("Google Sheetsクライアントが初期化されていません")
            return False
        
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            
            # シートを取得または作成
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
                logger.info(f"新しいシート '{sheet_name}' を作成しました")
            
            # データを準備
            if not data:
                logger.warning("書き込むデータがありません")
                return False
            
            # ヘッダー
            headers = list(data[0].to_dict().keys())
            
            # データ行
            rows = [headers]
            for item in data:
                rows.append(list(item.to_dict().values()))
            
            # 既存データをクリアして新しいデータを書き込み
            worksheet.clear()
            worksheet.update('A1', rows)
            
            logger.info(f"Google Sheetsに {len(data)} 件のデータを書き込みました")
            return True
            
        except Exception as e:
            logger.error(f"Google Sheetsへの書き込み中にエラーが発生しました: {e}")
            return False

class PRTimesScraper:
    """メインスクレイパークラス"""
    
    def __init__(self, credentials_path: str):
        self.searcher = PRTimesSearcher()
        self.extractor = DataExtractor()
        self.sheets_connector = GoogleSheetsConnector(credentials_path)
        
    def process_article(self, url: str) -> Optional[ContactInfo]:
        """単一記事を処理して連絡先情報を抽出"""
        try:
            logger.debug(f"記事を処理中: {url}")
            response = self.searcher.session.get(url, timeout=15)
            response.encoding = 'utf-8'
            
            if response.status_code != 200:
                logger.warning(f"記事アクセス失敗 (ステータス: {response.status_code}): {url}")
                return None
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 基本情報を抽出
            company_name, company_confidence = self.extractor.extract_company_name(soup, url)
            email, phone, person, sources = self.extractor.extract_contact_info(soup, response.text)
            
            # ContactInfoオブジェクトを作成
            info = ContactInfo(
                article_url=url,
                company_name=company_name,
                contact_person=person,
                email=email,
                phone=phone
            )
            
            # 信頼度を計算
            info.confidence_score = self.extractor.calculate_confidence_score(info, sources)
            info.extraction_source = f"company:{company_confidence:.2f}, " + ", ".join([f"{k}:{v:.2f}" for k, v in sources.items()])
            
            return info
            
        except Exception as e:
            logger.error(f"記事処理中にエラーが発生しました ({url}): {e}")
            return None
    
    def scrape(self, keyword: str, max_articles: int = 50, min_confidence: float = 0.3) -> List[ContactInfo]:
        """メインスクレイピング処理"""
        logger.info(f"スクレイピング開始: キーワード='{keyword}', 最大記事数={max_articles}")
        
        # 1. 記事URLを検索
        article_urls = self.searcher.search_articles(keyword, max_articles)
        
        if not article_urls:
            logger.error("記事が見つかりませんでした")
            return []
        
        # 2. 各記事を処理
        results = []
        for i, url in enumerate(article_urls, 1):
            logger.info(f"処理中: {i}/{len(article_urls)} - {url}")
            
            info = self.process_article(url)
            if info and info.confidence_score >= min_confidence:
                results.append(info)
                
                # 進捗表示（最初の5件）
                if i <= 5:
                    logger.info(f"  会社名: {info.company_name}")
                    if info.email:
                        logger.info(f"  Email: {info.email}")
                    if info.phone:
                        logger.info(f"  TEL: {info.phone}")
                    logger.info(f"  信頼度: {info.confidence_score:.2f}")
            
            # レート制限
            if i % 5 == 0:
                time.sleep(2)
            else:
                time.sleep(0.5)
        
        logger.info(f"スクレイピング完了: {len(results)}件のデータを抽出")
        return results
    
    def save_to_csv(self, data: List[ContactInfo], filename: str = "prtimes_v2_data.csv"):
        """CSVファイルに保存"""
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                if data:
                    fieldnames = list(data[0].to_dict().keys())
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    
                    for item in data:
                        writer.writerow(item.to_dict())
                    
                    logger.info(f"CSVファイルに保存しました: {filename} ({len(data)}件)")
                else:
                    logger.warning("保存するデータがありません")
                    
        except Exception as e:
            logger.error(f"CSV保存中にエラーが発生しました: {e}")
    
    def save_to_sheets(self, data: List[ContactInfo], spreadsheet_id: str, sheet_name: str) -> bool:
        """Google Sheetsに保存"""
        return self.sheets_connector.write_data(data, spreadsheet_id, sheet_name)
    
    def generate_report(self, data: List[ContactInfo]) -> Dict[str, any]:
        """結果レポートを生成"""
        if not data:
            return {"total": 0}
        
        email_count = sum(1 for item in data if item.email)
        phone_count = sum(1 for item in data if item.phone)
        person_count = sum(1 for item in data if item.contact_person)
        
        avg_confidence = sum(item.confidence_score for item in data) / len(data)
        
        high_confidence = [item for item in data if item.confidence_score >= 0.8]
        
        return {
            "total": len(data),
            "email_count": email_count,
            "phone_count": phone_count,
            "person_count": person_count,
            "email_rate": f"{email_count/len(data)*100:.1f}%",
            "phone_rate": f"{phone_count/len(data)*100:.1f}%",
            "person_rate": f"{person_count/len(data)*100:.1f}%",
            "avg_confidence": f"{avg_confidence:.2f}",
            "high_confidence_count": len(high_confidence),
            "high_confidence_rate": f"{len(high_confidence)/len(data)*100:.1f}%"
        }

def main():
    """メイン実行関数"""
    # 設定
    try:
        import config
        CREDENTIALS_PATH = config.GOOGLE_CREDENTIALS_PATH
        SPREADSHEET_ID = config.SPREADSHEET_ID
        SHEET_NAME = config.SHEET_NAME
        SEARCH_KEYWORD = config.DEFAULT_SEARCH_KEYWORD
    except ImportError:
        logger.warning("config.pyが見つかりません。デフォルト設定を使用します。")
        CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        SPREADSHEET_ID = 'your_spreadsheet_id'
        SHEET_NAME = 'PR_Times_V2_Data'
        SEARCH_KEYWORD = 'サプリ'
    
    # スクレイパー初期化
    scraper = PRTimesScraper(CREDENTIALS_PATH)
    
    # スクレイピング実行
    results = scraper.scrape(SEARCH_KEYWORD, max_articles=50, min_confidence=0.3)
    
    if results:
        # CSV保存
        scraper.save_to_csv(results, "prtimes_v2_data.csv")
        
        # Google Sheets保存
        sheets_success = scraper.save_to_sheets(results, SPREADSHEET_ID, SHEET_NAME)
        
        # レポート生成
        report = scraper.generate_report(results)
        
        logger.info("\n" + "="*50)
        logger.info("スクレイピング結果レポート")
        logger.info("="*50)
        logger.info(f"総記事数: {report['total']}件")
        logger.info(f"メールアドレス取得: {report['email_count']}件 ({report['email_rate']})")
        logger.info(f"電話番号取得: {report['phone_count']}件 ({report['phone_rate']})")
        logger.info(f"担当者名取得: {report['person_count']}件 ({report['person_rate']})")
        logger.info(f"平均信頼度: {report['avg_confidence']}")
        logger.info(f"高信頼度データ: {report['high_confidence_count']}件 ({report['high_confidence_rate']})")
        logger.info(f"Google Sheets保存: {'成功' if sheets_success else '失敗'}")
        logger.info("="*50)
        
    else:
        logger.error("データが取得できませんでした")

if __name__ == '__main__':
    main()