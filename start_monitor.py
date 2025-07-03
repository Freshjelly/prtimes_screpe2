#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import subprocess
import sys
import os
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Google Sheetsモニターを起動するメインスクリプト
    """
    print("🚀 PR Times Google Sheets スクレイパーシステム")
    print("=" * 50)
    
    # 1. Google Sheetsの初期設定
    print("📋 Step 1: Google Sheetsの初期設定...")
    try:
        setup_script = os.path.join(os.path.dirname(__file__), 'setup_sheets.py')
        result = subprocess.run([sys.executable, setup_script], 
                               capture_output=True, text=True)
        if result.returncode == 0:
            print("✅ Google Sheetsの設定が完了しました")
        else:
            print(f"⚠️  設定警告: {result.stderr}")
    except Exception as e:
        print(f"❌ 設定エラー: {e}")
        return
    
    # 2. 設定情報の表示
    try:
        import config
        spreadsheet_id = config.SPREADSHEET_ID
        print(f"\n📊 スプレッドシートURL:")
        print(f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
        print(f"\n🎯 使用方法:")
        print("1. 上記URLでGoogle Sheetsを開く")
        print("2. 'Control'シートを選択")
        print("3. A4セルにキーワードを入力（例：美容）")
        print("4. B4セルに「実行」と入力")
        print("5. 自動でスクレイピングが開始されます")
    except ImportError:
        print("❌ config.py が見つかりません")
        return
    
    # 3. モニター開始の確認
    print(f"\n🔍 Step 2: Google Sheetsモニターを開始します...")
    print("📝 Google SheetsでControlシートを確認してください")
    
    response = input("\nモニターを開始しますか？ (y/n): ").lower()
    if response != 'y':
        print("👋 モニターの開始をキャンセルしました")
        return
    
    # 4. モニター開始
    print("\n🎯 Google Sheetsモニターを開始中...")
    print("🛑 停止するには Ctrl+C を押してください")
    print("-" * 50)
    
    try:
        monitor_script = os.path.join(os.path.dirname(__file__), 'sheets_monitor.py')
        subprocess.run([sys.executable, monitor_script])
    except KeyboardInterrupt:
        print("\n👋 モニターを停止しました")
    except Exception as e:
        print(f"\n❌ モニターエラー: {e}")

if __name__ == '__main__':
    main()