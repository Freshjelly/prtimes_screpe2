# PR Times スクレイパー 使用方法

## 基本的な使い方

### 1. 単一キーワードで検索
```bash
python prtimes_corrected_scraper.py --keyword "美容"
```

### 2. 複数キーワードで検索（config.pyのSEARCH_KEYWORDSを使用）
```bash
python prtimes_corrected_scraper.py --multiple
```

### 3. ブラウザを表示して実行（デバッグ用）
```bash
python prtimes_corrected_scraper.py --keyword "美容" --no-headless
```

## config.pyでキーワードを設定

```python
SEARCH_KEYWORDS = [
    '匂い', '臭い', '体臭', '脱毛', '薄毛', '植毛', 'ハゲ', 
    'ホワイトニング', '矯正', '痩身', 'アンチエイジング', '清潔', 
    'メンズ美容', '若見え', '男性用化粧品', 'メンズコスメ', 'AGA', 
    'マインドフルネス', '瞑想', '語学', '脳トレ', '集中', '記憶力', 
    '知的', 'ニューロフィードバック'
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