# -*- coding: utf-8 -*-
"""
PR TIMESスクレイパー設定ファイル
実際の値に置き換えてください
"""

# PR TIMESログイン情報
PRTIMES_USERNAME = "your_email@example.com"  # PR TIMESのログインメールアドレス
PRTIMES_PASSWORD = "your_password"  # PR TIMESのパスワード

# Google Sheets認証情報
GOOGLE_CREDENTIALS_FILE = "credentials.json"  # Google Cloud Consoleから取得したJSONファイルのパス

# Google Sheetsの設定
SPREADSHEET_ID = "your_spreadsheet_id"  # スプレッドシートのID（URLの/d/と/editの間の文字列）
SHEET_NAME = "PR_TIMES_Data"  # 書き込むシート名

# スクレイピング設定
SEARCH_KEYWORD = "サプリ"  # 検索キーワード
MAX_ARTICLES = 50  # 取得する最大記事数

# オプション設定
REQUEST_TIMEOUT = 30  # リクエストのタイムアウト（秒）
RETRY_COUNT = 3  # リトライ回数
DELAY_BETWEEN_REQUESTS = 1.5  # リクエスト間の待機時間（秒）