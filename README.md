# 📰 PR Times Google Sheets スクレイパー

PR Timesのプレスリリースから企業の連絡先情報を自動抽出し、Google Sheetsで管理できるWebスクレイパーです。

## 🚀 機能

- **Google Sheets連携**: シート上でキーワードを入力するだけでスクレイピング実行
- **自動データ抽出**: 会社名、担当者名、メール、電話番号を自動抽出
- **リアルタイム保存**: 結果を自動でGoogle Sheetsに保存
- **キーワード検索**: 任意のキーワードで検索可能

## 📋 セットアップ

### 1. 依存関係のインストール
```bash
pip install -r requirements.txt
```

### 2. 設定ファイルの準備
```bash
cp config.py.example config.py
# config.pyを編集して認証情報を入力
```

### 3. Google認証ファイル
`credentials.json`をプロジェクトルートに配置

## 🎯 使用方法

### Google Sheets連携モード（推奨）

#### 1. システム起動
```bash
python start_monitor.py
```

#### 2. Google Sheetsで操作
1. 表示されたURLでGoogle Sheetsを開く
2. 'Control'シートを選択  
3. **A4セル**にキーワードを入力（例：「美容」）
4. **B4セル**に「実行」と入力
5. 自動でスクレイピングが開始されます

### コマンドライン実行モード

```bash
# 特定キーワードで検索
python prtimes_corrected_scraper.py --keyword "美容"

# ブラウザ表示で実行
python prtimes_corrected_scraper.py --keyword "健康食品" --no-headless

# デフォルト（サプリ）で検索
python prtimes_corrected_scraper.py
```

## 📊 出力データ

以下の情報が抽出されます：
- 記事URL
- 会社名  
- 担当者名
- メールアドレス
- 電話番号

結果は以下に保存されます：
- **CSV**: `prtimes_corrected_data.csv`
- **Google Sheets**: 設定したスプレッドシートに自動保存

## ⚙️ 設定

`config.py`で以下を設定：
- PR Times ログイン情報
- Google Sheets スプレッドシートID
- 検索設定

## 🔄 システム構成

- **prtimes_corrected_scraper.py**: メインスクレイパー
- **sheets_monitor.py**: Google Sheets監視システム  
- **setup_sheets.py**: Google Sheets初期設定
- **start_monitor.py**: システム起動スクリプト

## 📞 サポート

問題が発生した場合は、ログを確認してください。Google Sheets上の'Control'シートでステータスも確認できます。