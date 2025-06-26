#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Sheets認証テスト
"""

import json
from pathlib import Path
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def test_google_auth():
    """Google Sheets認証をテスト"""
    
    # 設定
    try:
        import config
        credentials_path = config.GOOGLE_CREDENTIALS_PATH
        spreadsheet_id = config.SPREADSHEET_ID
        sheet_name = config.SHEET_NAME
    except ImportError:
        credentials_path = '/mnt/c/Users/hyuga/Downloads/credentials.json'
        spreadsheet_id = 'your_spreadsheet_id'
        sheet_name = 'test'
    
    print("Google Sheets認証テストを開始します...")
    print(f"認証ファイル: {credentials_path}")
    print(f"スプレッドシートID: {spreadsheet_id}")
    print(f"シート名: {sheet_name}")
    
    # ステップ1: 認証ファイルの存在確認
    print("\n1. 認証ファイルの確認...")
    creds_path = Path(credentials_path)
    if not creds_path.exists():
        print(f"❌ 認証ファイルが見つかりません: {credentials_path}")
        return False
    print(f"✅ 認証ファイルが見つかりました")
    
    # ステップ2: 認証ファイルの内容確認
    print("\n2. 認証ファイルの内容確認...")
    try:
        with open(creds_path, 'r', encoding='utf-8') as f:
            creds_data = json.load(f)
        
        required_fields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email', 'client_id']
        missing_fields = [field for field in required_fields if field not in creds_data]
        
        if missing_fields:
            print(f"❌ 必要なフィールドが不足しています: {missing_fields}")
            return False
        
        print(f"✅ 認証ファイルの形式が正しいです")
        print(f"   プロジェクトID: {creds_data['project_id']}")
        print(f"   クライアントメール: {creds_data['client_email']}")
        
    except Exception as e:
        print(f"❌ 認証ファイルの読み込みに失敗しました: {e}")
        return False
    
    # ステップ3: Google Sheets API認証
    print("\n3. Google Sheets API認証...")
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            str(creds_path), scope
        )
        client = gspread.authorize(creds)
        print("✅ Google Sheets API認証に成功しました")
        
    except Exception as e:
        print(f"❌ Google Sheets API認証に失敗しました: {e}")
        print(f"   エラー詳細: {type(e).__name__}")
        return False
    
    # ステップ4: スプレッドシートアクセステスト
    print("\n4. スプレッドシートアクセステスト...")
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        print(f"✅ スプレッドシートにアクセスできました")
        print(f"   スプレッドシート名: {spreadsheet.title}")
        
        # シート一覧
        worksheets = spreadsheet.worksheets()
        print(f"   既存シート: {[ws.title for ws in worksheets]}")
        
    except Exception as e:
        print(f"❌ スプレッドシートアクセスに失敗しました: {e}")
        print(f"   エラー詳細: {type(e).__name__}")
        
        if 'Forbidden' in str(e) or '403' in str(e):
            print("   → サービスアカウントにスプレッドシートの権限がない可能性があります")
            print(f"   → スプレッドシートの共有設定で {creds_data['client_email']} に編集権限を付与してください")
        elif 'not found' in str(e).lower() or '404' in str(e):
            print("   → スプレッドシートIDが間違っている可能性があります")
        
        return False
    
    # ステップ5: 書き込みテスト
    print("\n5. 書き込みテスト...")
    try:
        # テスト用シートを取得または作成
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=10)
            print(f"   新しいシート '{sheet_name}' を作成しました")
        
        # テストデータを書き込み
        test_data = [
            ['テスト項目', 'テスト値', 'タイムスタンプ'],
            ['認証テスト', '成功', '2025-06-26 11:00:00'],
            ['書き込みテスト', '完了', '2025-06-26 11:00:01']
        ]
        
        worksheet.clear()
        worksheet.update('A1', test_data)
        print("✅ 書き込みテストに成功しました")
        
        # 読み込みテスト
        values = worksheet.get_all_values()
        print(f"   読み込み確認: {len(values)}行のデータを読み込めました")
        
    except Exception as e:
        print(f"❌ 書き込みテストに失敗しました: {e}")
        print(f"   エラー詳細: {type(e).__name__}")
        return False
    
    print("\n🎉 すべてのテストに成功しました！")
    return True

if __name__ == '__main__':
    success = test_google_auth()
    if success:
        print("\nGoogle Sheets連携の準備が完了しています。")
    else:
        print("\nGoogle Sheets連携に問題があります。上記のエラーを確認してください。")