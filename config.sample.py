# PR Timesのログイン情報
PRTIMES_EMAIL = 'your_email@example.com'
PRTIMES_PASSWORD = 'your_password'

# Google認証情報（オプション）
# Google Sheetsを使用する場合は、Google Cloud Consoleで取得した認証情報ファイルのパスを指定
GOOGLE_CREDENTIALS_PATH = '/path/to/credentials.json'
SPREADSHEET_ID = 'your_spreadsheet_id'
SHEET_NAME = 'PR_Times_Data'

# 検索設定
DEFAULT_SEARCH_KEYWORD = 'サプリ'

# 複数キーワード検索時に使用するキーワードリスト
SEARCH_KEYWORDS = [
    '匂い', '臭い', '体臭', '脱毛', '薄毛', '植毛', 'ハゲ', 
    'ホワイトニング', '矯正', '痩身', 'アンチエイジング', '清潔', 
    'メンズ美容', '若見え', '男性用化粧品', 'メンズコスメ', 'AGA', 
    'マインドフルネス', '瞑想', '語学', '脳トレ', '集中', '記憶力', 
    '知的', 'ニューロフィードバック'
]

# スクレイピング設定
MAX_ARTICLES_PER_KEYWORD = 100  # 各キーワードで収集する最大記事数
WAIT_TIME_BETWEEN_KEYWORDS = 5  # キーワード間の待機時間（秒）
WAIT_TIME_BETWEEN_ARTICLES = 0.5  # 記事間の待機時間（秒）