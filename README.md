# PR TIMES 自動スクレイピング & Google Sheets連携ツール

PR TIMESに自動ログインし、指定キーワードで検索して記事情報を収集し、Google Sheetsに自動書き込みするPythonスクリプトです。

## 機能

- PR TIMESへの自動ログイン（CSRFトークン対応）
- キーワード検索による記事URL収集（最大50件）
- 各記事からメディア関係者情報の抽出
  - 会社名
  - 担当者名
  - メールアドレス
  - 電話番号
- Google Sheetsへの自動書き込み
- CSVファイルへのバックアップ保存

## 必要な準備

### 1. Pythonライブラリのインストール

```bash
pip install -r requirements.txt
```

### 2. Google Cloud Console設定

1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新しいプロジェクトを作成
3. Google Sheets APIとGoogle Drive APIを有効化
4. サービスアカウントを作成
5. 認証情報（JSONファイル）をダウンロード
6. ダウンロードしたJSONファイルを`credentials.json`として保存

### 3. Google Sheetsの準備

1. Google Sheetsで新しいスプレッドシートを作成
2. スプレッドシートのURLから`SPREADSHEET_ID`を取得
   - 例: `https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit`
3. サービスアカウントのメールアドレスにスプレッドシートの編集権限を付与

### 4. 設定ファイルの編集

`config.py`を編集して、以下の情報を設定：

```python
# PR TIMESログイン情報
PRTIMES_USERNAME = "your_email@example.com"
PRTIMES_PASSWORD = "your_password"

# Google Sheets認証情報
GOOGLE_CREDENTIALS_FILE = "credentials.json"

# Google Sheetsの設定
SPREADSHEET_ID = "your_spreadsheet_id"
SHEET_NAME = "PR_TIMES_Data"

# スクレイピング設定
SEARCH_KEYWORD = "サプリ"
MAX_ARTICLES = 50
```

## 使い方

```bash
python prtimes_scraper.py
```

## 出力

### Google Sheets
指定したスプレッドシートに以下の形式でデータが書き込まれます：

| 記事URL | 会社名 | 担当者名 | メールアドレス | 電話番号 |
|---------|--------|----------|----------------|----------|
| https://... | 株式会社〇〇 | 山田太郎 | info@example.com | 03-1234-5678 |

### CSVファイル
バックアップとして、以下の形式でCSVファイルが保存されます：
- ファイル名: `prtimes_results_[キーワード]_[日時].csv`

## 注意事項

- PR TIMESの利用規約を遵守してください
- 過度なアクセスを避けるため、適切な待機時間を設定しています
- ログイン情報は安全に管理してください
- Google Sheetsの認証情報（credentials.json）は公開しないでください

## トラブルシューティング

### ログインできない場合
- ユーザー名とパスワードが正しいか確認
- PR TIMESのログインページの仕様が変更されていないか確認

### Google Sheetsに書き込めない場合
- サービスアカウントに編集権限があるか確認
- スプレッドシートIDが正しいか確認
- credentials.jsonが正しい場所にあるか確認

### データが抽出できない場合
- PR TIMESのHTML構造が変更されている可能性があります
- ログを確認して、エラーの詳細を確認してください

## ライセンス

このスクリプトは教育・研究目的で作成されています。
商用利用の際は、PR TIMESの利用規約を確認してください。