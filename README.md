# PR Times Scraper

PR Timesのプレスリリースから企業の連絡先情報を収集するスクレイパーです。

## 機能

- PR Timesにログインして記事を検索
- キーワードに基づいて記事URLを自動収集
- 各記事から以下の情報を抽出：
  - 会社名
  - 担当者名
  - メールアドレス
  - 電話番号
- 結果をExcelファイル（キーワード別シート）とCSVファイルに出力
- Google Sheetsとの連携機能（オプション）

## 必要な環境

- Python 3.7以上
- Google Chrome/Chromium
- ChromeDriver（自動インストール）

## インストール

1. リポジトリをクローン
```bash
git clone https://github.com/your-username/prtimes_scraper.git
cd prtimes_scraper
```

2. 必要なパッケージをインストール
```bash
pip install -r requirements.txt
```

3. 設定ファイルを作成
```bash
cp config.sample.py config.py
```

4. `config.py`を編集して認証情報を設定

## 使用方法

### 単一キーワードで検索
```bash
python prtimes_corrected_scraper.py --keyword "美容"
```

### 複数キーワードで検索（config.pyのSEARCH_KEYWORDSを使用）
```bash
python prtimes_corrected_scraper.py --multiple
```

### ブラウザを表示して実行（デバッグ用）
```bash
python prtimes_corrected_scraper.py --keyword "美容" --no-headless
```

## 設定

`config.py`で以下の設定が必要です：

- `PRTIMES_EMAIL`: PR Timesのログインメールアドレス
- `PRTIMES_PASSWORD`: PR Timesのログインパスワード
- `GOOGLE_CREDENTIALS_PATH`: Google認証用JSONファイルのパス（オプション）
- `SEARCH_KEYWORDS`: 複数キーワード検索時のキーワードリスト
- `DEFAULT_SEARCH_KEYWORD`: デフォルトの検索キーワード

### 設定例（config.py）

```python
# PR Timesのログイン情報
PRTIMES_EMAIL = 'your_email@example.com'
PRTIMES_PASSWORD = 'your_password'

# Google認証情報（オプション）
GOOGLE_CREDENTIALS_PATH = '/path/to/credentials.json'
SPREADSHEET_ID = 'your_spreadsheet_id'
SHEET_NAME = 'PR_Times_Data'

# 検索キーワード
DEFAULT_SEARCH_KEYWORD = 'サプリ'
SEARCH_KEYWORDS = [
    '匂い', '臭い', '体臭', '脱毛', '薄毛', '植毛', 'ハゲ', 
    'ホワイトニング', '矯正', '痩身', 'アンチエイジング', '清潔', 
    'メンズ美容', '若見え', '男性用化粧品', 'メンズコスメ', 'AGA'
]
```

## 出力ファイル

- **Excel**: `prtimes_all_keywords_YYYYMMDD_HHMMSS.xlsx` (キーワード別シート)
- **CSV**: `prtimes_data_YYYYMMDD_HHMMSS.csv` (記事ごとにページ分割)

## オプション

- `--keyword`, `-k`: 単一キーワード指定
- `--multiple`, `-m`: 複数キーワードモード
- `--no-headless`: ブラウザ表示モード
- `--help`, `-h`: ヘルプ表示

## 注意事項

- PR Timesの利用規約を遵守してください
- 過度なアクセスを避けるため、適切な待機時間を設定しています
- ログイン情報は安全に管理してください
- Google認証情報ファイル（credentials.json）は絶対にリポジトリにコミットしないでください

## トラブルシューティング

### ChromeDriverエラー
- Chromiumがインストールされていることを確認
- WSL環境の場合は `sudo snap install chromium` でインストール

### ログインエラー
- PR Timesのログイン情報が正しいか確認
- `--no-headless`オプションでブラウザの動作を確認

## ライセンス

MIT License