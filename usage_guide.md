# PR Times Scraper 使用方法

## セットアップ手順

### 1. 必要なライブラリのインストール
```bash
pip install -r requirements.txt
```

### 2. Google Sheets APIの設定
1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新しいプロジェクトを作成
3. Google Sheets APIとGoogle Drive APIを有効化
4. サービスアカウントを作成し、JSONキーをダウンロード
5. ダウンロードしたJSONファイルを指定の場所に配置

### 3. スクリプトの設定
`prtimes_scraper.py`の`main()`関数内の以下の変数を設定：

```python
EMAIL = 'your_email@example.com'  # PR Timesのメールアドレス
PASSWORD = 'your_password'  # PR Timesのパスワード
CREDENTIALS_PATH = '/mnt/c/Users/hyuga/Downloads/credentials.json'  # Google認証情報
SPREADSHEET_ID = 'your_spreadsheet_id'  # Google SheetsのスプレッドシートID
SHEET_NAME = 'PR_Times_Data'  # シート名
SEARCH_KEYWORD = 'サプリ'  # 検索キーワード
```

### 4. Google Sheetsの準備
1. 新しいGoogleスプレッドシートを作成
2. スプレッドシートのURLから`spreadsheet_id`を取得
   - URL例: `https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit`
3. サービスアカウントのメールアドレスをスプレッドシートの共有設定に追加（編集権限を付与）

## 実行方法
```bash
python prtimes_scraper.py
```

## 出力
- Google Sheetsに以下の情報が書き込まれます：
  - 記事URL
  - 会社名
  - 担当者名
  - メールアドレス
  - 電話番号
- バックアップとして`prtimes_data.csv`も作成されます

## 注意事項
- PR Timesの利用規約を遵守してください
- 過度なアクセスを避けるため、適切な間隔でリクエストを送信しています
- ログイン情報は安全に管理してください