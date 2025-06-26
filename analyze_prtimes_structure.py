#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import json
from typing import Dict, List

def analyze_article_structure():
    """PR TIMES記事の構造を詳細に分析"""
    
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    })
    
    # 実際の記事URLを分析
    sample_urls = [
        'https://prtimes.jp/main/html/rd/p/000001550.000006302.html',
        'https://prtimes.jp/main/html/rd/p/000000156.000061950.html',
        'https://prtimes.jp/main/html/rd/p/000000067.000037237.html'
    ]
    
    analysis_results = []
    
    for i, url in enumerate(sample_urls):
        print(f"\n{'='*50}")
        print(f"記事 {i+1}: {url}")
        print(f"{'='*50}")
        
        try:
            response = session.get(url)
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            
            result = {
                'url': url,
                'title': '',
                'company_candidates': [],
                'contact_sections': [],
                'email_candidates': [],
                'phone_candidates': [],
                'structured_data': {}
            }
            
            # タイトル
            title = soup.find('title')
            if title:
                result['title'] = title.text.strip()
                print(f"タイトル: {result['title']}")
            
            # 会社名の候補を複数パターンで探す
            company_patterns = [
                ('div', {'class': re.compile('company', re.I)}),
                ('span', {'class': re.compile('company', re.I)}),
                ('a', {'class': re.compile('company', re.I)}),
                ('div', {'class': re.compile('release', re.I)}),
                ('meta', {'property': 'og:site_name'}),
                ('meta', {'name': 'author'})
            ]
            
            print("\n会社名候補:")
            for tag, attrs in company_patterns:
                elements = soup.find_all(tag, attrs)
                for elem in elements:
                    if tag == 'meta':
                        text = elem.get('content', '')
                    else:
                        text = elem.text.strip()
                    
                    if text and len(text) < 100:
                        result['company_candidates'].append({
                            'pattern': f"{tag}_{str(attrs)}",
                            'text': text
                        })
                        print(f"  {tag}: {text}")
            
            # お問い合わせセクションを探す
            contact_keywords = [
                'お問い合わせ', '問い合わせ', '連絡先', 'Contact', 'CONTACT',
                'プレスリリース詳細', '本件に関する', '報道関係者', 'メディア',
                '広報', 'PR', 'Press'
            ]
            
            print("\nお問い合わせセクション:")
            for keyword in contact_keywords:
                # 見出しを探す
                headings = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'], 
                                        string=re.compile(keyword, re.I))
                for heading in headings:
                    # 見出しの次のコンテンツを取得
                    next_content = heading.find_next_siblings()
                    if next_content:
                        content_text = ' '.join([elem.text.strip() for elem in next_content[:3]])
                        result['contact_sections'].append({
                            'keyword': keyword,
                            'heading': heading.text.strip(),
                            'content': content_text[:200]
                        })
                        print(f"  {keyword}: {heading.text.strip()}")
                        print(f"    内容: {content_text[:100]}...")
            
            # テーブル形式の情報を探す
            tables = soup.find_all('table')
            print(f"\nテーブル数: {len(tables)}")
            for i, table in enumerate(tables):
                if any(keyword in table.text for keyword in contact_keywords):
                    print(f"  テーブル{i+1}: お問い合わせ関連")
                    rows = table.find_all('tr')
                    for row in rows:
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= 2:
                            key = cells[0].text.strip()
                            value = cells[1].text.strip()
                            print(f"    {key}: {value}")
            
            # メールアドレスを詳細に探す
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            email_matches = re.findall(email_pattern, response.text)
            print(f"\nメールアドレス候補: {len(email_matches)}件")
            for email in set(email_matches):
                if 'prtimes' not in email.lower():
                    result['email_candidates'].append(email)
                    print(f"  {email}")
            
            # 電話番号を詳細に探す
            phone_patterns = [
                r'(?:TEL|Tel|tel|電話|℡)[\s:：]*([0-9０-９\-－\s\(\)]{10,20})',
                r'(?:^|\s)(0[0-9]{1,4}[\-－\s]?[0-9]{1,4}[\-－\s]?[0-9]{3,4})(?:\s|$)',
                r'(?:^|\s)(０[０-９]{1,4}[\-－\s]?[０-９]{1,4}[\-－\s]?[０-９]{3,4})(?:\s|$)'
            ]
            
            print(f"\n電話番号候補:")
            for pattern in phone_patterns:
                phone_matches = re.findall(pattern, response.text)
                for match in phone_matches:
                    if isinstance(match, tuple):
                        match = match[0] if match[0] else match[1]
                    
                    # 全角を半角に変換
                    phone = match.translate(str.maketrans('０１２３４５６７８９－（）　', '0123456789-() '))
                    phone = re.sub(r'\s+', '', phone.strip())
                    
                    if len(phone) >= 10 and phone not in [r['phone'] for r in result['phone_candidates']]:
                        result['phone_candidates'].append({
                            'pattern': pattern,
                            'phone': phone,
                            'original': match
                        })
                        print(f"  {phone} (元: {match})")
            
            # 構造化データ（JSON-LD）を探す
            json_scripts = soup.find_all('script', {'type': 'application/ld+json'})
            print(f"\n構造化データ: {len(json_scripts)}件")
            for script in json_scripts:
                try:
                    data = json.loads(script.string)
                    result['structured_data'] = data
                    print(f"  構造化データが見つかりました: {list(data.keys())}")
                except:
                    pass
            
            analysis_results.append(result)
            
        except Exception as e:
            print(f"エラー: {e}")
    
    # 分析結果を保存
    with open('prtimes_structure_analysis.json', 'w', encoding='utf-8') as f:
        json.dump(analysis_results, f, ensure_ascii=False, indent=2)
    
    print(f"\n分析完了。結果をprtimes_structure_analysis.jsonに保存しました。")
    return analysis_results

if __name__ == '__main__':
    analyze_article_structure()