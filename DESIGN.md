# WOC講義資料ダウンロードシステム 設計ドキュメント

## アーキテクチャ概要

```
┌─────────────────────────────────────────────────────────────┐
│                         main.py                              │
│  ┌───────────────────────────────────────────────────────┐  │
│  │              WOCDownloadManager                        │  │
│  │  - Excelファイル読み込み                              │  │
│  │  - シートごとの処理オーケストレーション                │  │
│  │  - 進捗レポート生成                                   │  │
│  └───────────────────────────────────────────────────────┘  │
│           │                    │                    │         │
│           ▼                    ▼                    ▼         │
│  ┌─────────────┐    ┌──────────────┐    ┌──────────────┐    │
│  │  DownloadDB │    │  URLDedup    │    │DownloadQueue │    │
│  │  (Resume)   │    │  (重複排除)  │    │  (並列処理)  │    │
│  └─────────────┘    └──────────────┘    └──────────────┘    │
│           │                    │                    │         │
│           └────────────────────┴────────────────────┘         │
│                              │                                │
│                              ▼                                │
│                  ┌───────────────────────┐                    │
│                  │  DownloadExecutor     │                    │
│                  │  - yt-dlp実行         │                    │
│                  │  - gdown実行          │                    │
│                  │  - エラーハンドリング │                    │
│                  └───────────────────────┘                    │
└─────────────────────────────────────────────────────────────┘
```

## クラス設計

### 1. WOCDownloadManager

**目的**: システム全体の統括、Excelファイルの処理、各コンポーネントの調整

**属性**:
```python
class WOCDownloadManager:
    def __init__(
        self,
        excel_path: str,
        download_dir: str = "downloads",
        state_dir: str = "downloads/.download_state",
        dry_run: bool = False,
        overwrite: bool = False,
        no_dedup: bool = False,
        dedup_mode: str = "symlink",
        parallel: int = 1,
        target_sheets: Optional[List[str]] = None
    ):
        """
        Args:
            excel_path: Excelファイルのパス
            download_dir: ダウンロード先ディレクトリ
            state_dir: 状態管理DBの保存ディレクトリ
            dry_run: ドライランモード（実際にダウンロードしない）
            overwrite: 既存ファイルを上書きするか
            no_dedup: 重複排除機能を無効化
            dedup_mode: 重複時の動作 ("symlink"|"copy"|"reference")
            parallel: 並列ダウンロード数
            target_sheets: 処理対象シート名リスト（None=全シート）
        """
        self.excel_path: str
        self.download_dir: Path
        self.state_dir: Path
        self.dry_run: bool
        self.overwrite: bool
        self.no_dedup: bool
        self.dedup_mode: str
        self.parallel: int
        self.target_sheets: Optional[List[str]]

        # 各コンポーネント
        self.download_db: DownloadDB
        self.url_dedup: URLDedup
        self.executor: DownloadExecutor
        self.logger: Logger
```

**メソッド**:

#### `run() -> DownloadReport`
```python
def run(self) -> DownloadReport:
    """
    メイン処理を実行

    Returns:
        DownloadReport: ダウンロード結果のレポート
    """
```

#### `process_sheet(sheet_name: str, df: pd.DataFrame) -> SheetReport`
```python
def process_sheet(self, sheet_name: str, df: pd.DataFrame) -> SheetReport:
    """
    1つのシートを処理

    Args:
        sheet_name: シート名
        df: シートのDataFrame

    Returns:
        SheetReport: シート処理結果
    """
```

#### `process_row(sheet_name: str, row_index: int, row: pd.Series) -> List[DownloadTask]`
```python
def process_row(
    self,
    sheet_name: str,
    row_index: int,
    row: pd.Series
) -> List[DownloadTask]:
    """
    1行を処理してダウンロードタスクを生成

    Args:
        sheet_name: シート名
        row_index: 行インデックス
        row: 行データ

    Returns:
        List[DownloadTask]: 生成されたダウンロードタスクのリスト
    """
```

#### `generate_filename(sheet_name: str, row: pd.Series, column_name: str) -> str`
```python
def generate_filename(
    self,
    sheet_name: str,
    row: pd.Series,
    column_name: str
) -> str:
    """
    ファイル名を生成

    Args:
        sheet_name: シート名
        row: 行データ
        column_name: 列名（例: "録画（動画DLリンク）", "資料1"）

    Returns:
        str: サニタイズされたファイル名（拡張子なし）
    """
```

#### `show_status() -> None`
```python
def show_status(self) -> None:
    """
    現在のダウンロード状態を表示
    """
```

#### `show_dedup_stats() -> None`
```python
def show_dedup_stats(self) -> None:
    """
    重複排除の統計情報を表示
    """
```

#### `reset() -> None`
```python
def reset(self) -> None:
    """
    ダウンロード状態をリセット
    """
```

#### `retry_failed() -> DownloadReport`
```python
def retry_failed(self) -> DownloadReport:
    """
    失敗したダウンロードを再試行

    Returns:
        DownloadReport: 再試行結果のレポート
    """
```

---

### 2. DownloadDB

**目的**: ダウンロード状態の管理（Resume機能）

**属性**:
```python
class DownloadDB:
    def __init__(self, db_path: str):
        """
        Args:
            db_path: download_db.jsonのパス
        """
        self.db_path: Path
        self.data: Dict[str, Any]  # JSONデータ
        self.downloads: Dict[str, DownloadRecord]  # URLをキーとする辞書
```

**メソッド**:

#### `load() -> None`
```python
def load(self) -> None:
    """
    DBファイルを読み込む
    """
```

#### `save() -> None`
```python
def save(self) -> None:
    """
    DBファイルに保存
    """
```

#### `get_record(url: str) -> Optional[DownloadRecord]`
```python
def get_record(self, url: str) -> Optional[DownloadRecord]:
    """
    URLに対応するレコードを取得

    Args:
        url: ダウンロードURL

    Returns:
        Optional[DownloadRecord]: レコード（存在しない場合はNone）
    """
```

#### `is_completed(url: str, file_path: str) -> bool`
```python
def is_completed(self, url: str, file_path: str) -> bool:
    """
    ダウンロードが完了しているかチェック

    Args:
        url: ダウンロードURL
        file_path: ファイルパス

    Returns:
        bool: 完了している場合True

    Note:
        以下の条件をすべて満たす場合に完了とみなす:
        - DBにレコードが存在
        - statusが"completed"
        - ファイルが実際に存在
        - ファイルサイズが0より大きい
    """
```

#### `mark_completed(url: str, file_path: str, file_size: int) -> None`
```python
def mark_completed(self, url: str, file_path: str, file_size: int) -> None:
    """
    ダウンロード完了をマーク

    Args:
        url: ダウンロードURL
        file_path: ファイルパス
        file_size: ファイルサイズ（バイト）
    """
```

#### `mark_failed(url: str, file_path: str, error_message: str) -> None`
```python
def mark_failed(self, url: str, file_path: str, error_message: str) -> None:
    """
    ダウンロード失敗をマーク

    Args:
        url: ダウンロードURL
        file_path: ファイルパス
        error_message: エラーメッセージ
    """
```

#### `get_failed_records() -> List[DownloadRecord]`
```python
def get_failed_records(self) -> List[DownloadRecord]:
    """
    失敗したレコードの一覧を取得

    Returns:
        List[DownloadRecord]: 失敗レコードのリスト
    """
```

#### `get_statistics() -> DownloadStatistics`
```python
def get_statistics(self) -> DownloadStatistics:
    """
    統計情報を取得

    Returns:
        DownloadStatistics: 統計情報
    """
```

#### `reset() -> None`
```python
def reset(self) -> None:
    """
    すべてのレコードをクリア
    """
```

---

### 3. URLDedup

**目的**: URL重複排除管理（Dedup機能）

**属性**:
```python
class URLDedup:
    def __init__(self, db_path: str, dedup_mode: str = "symlink"):
        """
        Args:
            db_path: url_dedup.jsonのパス
            dedup_mode: 重複時の動作 ("symlink"|"copy"|"reference")
        """
        self.db_path: Path
        self.dedup_mode: str
        self.data: Dict[str, Any]  # JSONデータ
        self.url_hash_to_file: Dict[str, DedupRecord]
```

**メソッド**:

#### `load() -> None`
```python
def load(self) -> None:
    """
    DBファイルを読み込む
    """
```

#### `save() -> None`
```python
def save(self) -> None:
    """
    DBファイルに保存
    """
```

#### `get_url_hash(url: str) -> str`
```python
@staticmethod
def get_url_hash(url: str) -> str:
    """
    URLのハッシュを生成

    Args:
        url: ダウンロードURL

    Returns:
        str: SHA256ハッシュ（16進数文字列）
    """
```

#### `normalize_url(url: str) -> str`
```python
@staticmethod
def normalize_url(url: str) -> str:
    """
    URLを正規化

    Args:
        url: ダウンロードURL

    Returns:
        str: 正規化されたURL

    Note:
        - クエリパラメータを除去
        - 末尾のスラッシュを除去
        - スキーム、ホスト、パスのみを使用
    """
```

#### `is_duplicate(url: str) -> Tuple[bool, Optional[str]]`
```python
def is_duplicate(self, url: str) -> Tuple[bool, Optional[str]]:
    """
    URLが重複しているかチェック

    Args:
        url: ダウンロードURL

    Returns:
        Tuple[bool, Optional[str]]: (重複フラグ, 元ファイルパス)

    Note:
        元ファイルが存在しない場合は重複とみなさない
    """
```

#### `register(url: str, file_path: str, file_size: int) -> None`
```python
def register(self, url: str, file_path: str, file_size: int) -> None:
    """
    新規ダウンロードを登録

    Args:
        url: ダウンロードURL
        file_path: ファイルパス
        file_size: ファイルサイズ（バイト）
    """
```

#### `add_reference(url: str, new_file_path: str, link_type: str) -> str`
```python
def add_reference(
    self,
    url: str,
    new_file_path: str,
    link_type: str
) -> str:
    """
    参照を追加

    Args:
        url: ダウンロードURL
        new_file_path: 新しいファイルパス
        link_type: リンクタイプ ("symlink"|"copy"|"reference")

    Returns:
        str: 元ファイルのパス

    Raises:
        ValueError: URLが登録されていない場合
    """
```

#### `create_link(original_path: str, new_path: str) -> None`
```python
def create_link(self, original_path: str, new_path: str) -> None:
    """
    ファイルリンクを作成

    Args:
        original_path: 元ファイルのパス
        new_path: 新しいファイルのパス

    Note:
        dedup_modeに応じて以下を実行:
        - symlink: シンボリックリンク作成
        - copy: ファイルコピー
        - reference: 何もしない（DBに記録のみ）
    """
```

#### `get_statistics() -> DedupStatistics`
```python
def get_statistics(self) -> DedupStatistics:
    """
    重複排除の統計情報を取得

    Returns:
        DedupStatistics: 統計情報
    """
```

#### `get_top_duplicates(n: int = 5) -> List[Tuple[str, int]]`
```python
def get_top_duplicates(self, n: int = 5) -> List[Tuple[str, int]]:
    """
    重複が多いURLのトップN件を取得

    Args:
        n: 取得件数

    Returns:
        List[Tuple[str, int]]: [(URL, 参照数), ...] のリスト
    """
```

---

### 4. DownloadExecutor

**目的**: 実際のダウンロード処理の実行

**属性**:
```python
class DownloadExecutor:
    def __init__(
        self,
        dry_run: bool = False,
        logger: Optional[Logger] = None
    ):
        """
        Args:
            dry_run: ドライランモード
            logger: ロガー
        """
        self.dry_run: bool
        self.logger: Logger
```

**メソッド**:

#### `download(task: DownloadTask) -> DownloadResult`
```python
def download(self, task: DownloadTask) -> DownloadResult:
    """
    ダウンロードタスクを実行

    Args:
        task: ダウンロードタスク

    Returns:
        DownloadResult: ダウンロード結果
    """
```

#### `download_video(url: str, output_path: str) -> DownloadResult`
```python
def download_video(self, url: str, output_path: str) -> DownloadResult:
    """
    yt-dlpを使って動画をダウンロード

    Args:
        url: 動画URL
        output_path: 出力パス（拡張子なし）

    Returns:
        DownloadResult: ダウンロード結果
    """
```

#### `download_document(url: str, output_path: str) -> DownloadResult`
```python
def download_document(self, url: str, output_path: str) -> DownloadResult:
    """
    gdownを使って資料をダウンロード

    Args:
        url: 資料URL（Google Docs/Drive）
        output_path: 出力パス

    Returns:
        DownloadResult: ダウンロード結果

    Note:
        URLの種類に応じて適切なフォーマットで保存:
        - Google Slides -> PDF
        - Google Sheets -> XLSX
        - Google Docs -> PDF
        - Google Drive ファイル -> そのまま
    """
```

#### `detect_url_type(url: str) -> URLType`
```python
@staticmethod
def detect_url_type(url: str) -> URLType:
    """
    URLの種類を判定

    Args:
        url: ダウンロードURL

    Returns:
        URLType: URL種別
    """
```

#### `get_file_size(file_path: str) -> int`
```python
@staticmethod
def get_file_size(file_path: str) -> int:
    """
    ファイルサイズを取得

    Args:
        file_path: ファイルパス

    Returns:
        int: ファイルサイズ（バイト）
    """
```

---

### 5. FileNameGenerator

**目的**: ファイル名の生成とサニタイズ

**メソッド**:

#### `generate_filename(sheet_name: str, row: pd.Series, column_name: str) -> str`
```python
@staticmethod
def generate_filename(
    sheet_name: str,
    row: pd.Series,
    column_name: str
) -> str:
    """
    ファイル名を生成

    Args:
        sheet_name: シート名
        row: 行データ
        column_name: 列名

    Returns:
        str: ファイル名（拡張子なし、サニタイズ済み）
    """
```

#### `sanitize_filename(filename: str, max_length: int = 200) -> str`
```python
@staticmethod
def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """
    ファイル名をサニタイズ

    Args:
        filename: 元のファイル名
        max_length: 最大文字数

    Returns:
        str: サニタイズされたファイル名

    Note:
        以下の処理を行う:
        - 禁止文字を全角に置換
        - 改行・タブ・連続空白を除去
        - 最大長を制限
    """
```

#### `extract_date_from_row(row: pd.Series) -> str`
```python
@staticmethod
def extract_date_from_row(row: pd.Series) -> str:
    """
    行データから日付を抽出してYYYYMMDD形式にフォーマット

    Args:
        row: 行データ

    Returns:
        str: YYYYMMDD形式の日付文字列

    Note:
        実施年・実施日列から抽出
        "2025年"と"5月15日"から"20250515"を生成
    """
```

#### `extract_chapter_number(title: str) -> str`
```python
@staticmethod
def extract_chapter_number(title: str) -> str:
    """
    タイトルから章番号を抽出

    Args:
        title: コンテンツタイトル

    Returns:
        str: 章番号（例: "1-1", "2-3"）

    Note:
        正規表現 r'(\d+-\d+|\d+)' で抽出
    """
```

---

### 6. Logger

**目的**: ログ出力の管理

**属性**:
```python
class Logger:
    def __init__(self, log_dir: str = "logs"):
        """
        Args:
            log_dir: ログディレクトリ
        """
        self.log_dir: Path
        self.loggers: Dict[str, logging.Logger]
```

**メソッド**:

#### `info(message: str, category: str = "general") -> None`
```python
def info(self, message: str, category: str = "general") -> None:
    """
    INFOレベルのログを出力

    Args:
        message: ログメッセージ
        category: カテゴリ ("success"|"skip"|"resume"|"dedup"|"general")
    """
```

#### `error(message: str, category: str = "error") -> None`
```python
def error(self, message: str, category: str = "error") -> None:
    """
    ERRORレベルのログを出力

    Args:
        message: ログメッセージ
        category: カテゴリ
    """
```

#### `warning(message: str) -> None`
```python
def warning(self, message: str) -> None:
    """
    WARNINGレベルのログを出力

    Args:
        message: ログメッセージ
    """
```

---

## データクラス

### DownloadTask
```python
@dataclass
class DownloadTask:
    """ダウンロードタスク"""
    url: str
    file_path: str
    sheet_name: str
    row_index: int
    column_name: str
    url_type: URLType
```

### DownloadResult
```python
@dataclass
class DownloadResult:
    """ダウンロード結果"""
    success: bool
    file_path: Optional[str]
    file_size: Optional[int]
    error_message: Optional[str]
    skipped: bool = False
    deduped: bool = False
    original_file_path: Optional[str] = None  # Dedup時の元ファイルパス
```

### DownloadRecord
```python
@dataclass
class DownloadRecord:
    """ダウンロードレコード（Resume用）"""
    id: str  # UUID
    url: str
    file_path: str
    status: str  # "completed"|"failed"|"in_progress"
    file_size: Optional[int]
    downloaded_at: Optional[str]  # ISO 8601形式
    error_message: Optional[str]
    sheet_name: str
    row_index: int
    column_name: str
```

### DedupRecord
```python
@dataclass
class DedupRecord:
    """重複排除レコード"""
    url: str
    original_file_path: str
    file_size: int
    downloaded_at: str  # ISO 8601形式
    references: List[DedupReference]
```

### DedupReference
```python
@dataclass
class DedupReference:
    """重複参照"""
    file_path: str
    link_type: str  # "symlink"|"copy"|"reference"
    created_at: str  # ISO 8601形式
```

### DownloadStatistics
```python
@dataclass
class DownloadStatistics:
    """ダウンロード統計"""
    total_downloads: int
    completed: int
    failed: int
    skipped: int
    in_progress: int
```

### DedupStatistics
```python
@dataclass
class DedupStatistics:
    """重複排除統計"""
    total_unique_urls: int
    total_references: int
    space_saved_bytes: int
```

### SheetReport
```python
@dataclass
class SheetReport:
    """シート処理結果レポート"""
    sheet_name: str
    total_tasks: int
    completed: int
    failed: int
    skipped: int
    deduped: int
    errors: List[str]
```

### DownloadReport
```python
@dataclass
class DownloadReport:
    """全体レポート"""
    sheets: List[SheetReport]
    total_tasks: int
    completed: int
    failed: int
    skipped: int
    deduped: int
    execution_time: float  # 秒

    def print_summary(self) -> None:
        """サマリーを出力"""
        pass
```

### URLType
```python
class URLType(Enum):
    """URL種別"""
    VIMEO = "vimeo"
    YOUTUBE = "youtube"
    LOOM = "loom"
    UTAGE = "utage"
    M3U8 = "m3u8"
    GOOGLE_SLIDES = "google_slides"
    GOOGLE_SHEETS = "google_sheets"
    GOOGLE_DOCS = "google_docs"
    GOOGLE_DRIVE_FILE = "google_drive_file"
    GOOGLE_DRIVE_FOLDER = "google_drive_folder"
    UNKNOWN = "unknown"
```

---

## ユーティリティ関数

### `format_file_size(size_bytes: int) -> str`
```python
def format_file_size(size_bytes: int) -> str:
    """
    ファイルサイズを人間が読みやすい形式にフォーマット

    Args:
        size_bytes: バイト数

    Returns:
        str: フォーマットされた文字列（例: "125.5 MB"）
    """
```

### `parse_date(date_str: str) -> Optional[datetime]`
```python
def parse_date(date_str: str) -> Optional[datetime]:
    """
    日本語の日付文字列をパース

    Args:
        date_str: 日付文字列（例: "5月15日", "2025-05-15"）

    Returns:
        Optional[datetime]: パース結果（失敗時はNone）
    """
```

---

## メイン処理フロー

### main.py
```python
def main():
    """
    メインエントリーポイント
    """
    # コマンドライン引数のパース
    args = parse_arguments()

    # WOCDownloadManagerを初期化
    manager = WOCDownloadManager(
        excel_path=args.excel_path,
        download_dir=args.download_dir,
        dry_run=args.dry_run,
        overwrite=args.overwrite,
        no_dedup=args.no_dedup,
        dedup_mode=args.dedup_mode,
        parallel=args.parallel,
        target_sheets=args.sheet
    )

    # モードに応じて処理を実行
    if args.status:
        manager.show_status()
    elif args.dedup_stats:
        manager.show_dedup_stats()
    elif args.reset:
        manager.reset()
    elif args.retry_failed:
        report = manager.retry_failed()
        report.print_summary()
    else:
        report = manager.run()
        report.print_summary()


def parse_arguments() -> argparse.Namespace:
    """
    コマンドライン引数をパース

    Returns:
        argparse.Namespace: パース結果
    """
    parser = argparse.ArgumentParser(
        description="WOC講義資料ダウンロードシステム"
    )

    parser.add_argument(
        "excel_path",
        nargs="?",
        default="draft_【WOC】AIチャットボット開発データベース.xlsx",
        help="Excelファイルのパス"
    )

    parser.add_argument(
        "--download-dir",
        default="downloads",
        help="ダウンロード先ディレクトリ"
    )

    parser.add_argument(
        "--sheet",
        nargs="+",
        help="処理対象シート名"
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="ドライランモード"
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="既存ファイルを上書き"
    )

    parser.add_argument(
        "--parallel",
        type=int,
        default=1,
        help="並列ダウンロード数"
    )

    # Resume機能
    parser.add_argument(
        "--status",
        action="store_true",
        help="ダウンロード状態を表示"
    )

    parser.add_argument(
        "--reset",
        action="store_true",
        help="ダウンロード状態をリセット"
    )

    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="失敗したダウンロードを再試行"
    )

    # Dedup機能
    parser.add_argument(
        "--no-dedup",
        action="store_true",
        help="重複排除機能を無効化"
    )

    parser.add_argument(
        "--dedup-mode",
        choices=["symlink", "copy", "reference"],
        default="symlink",
        help="重複時の動作"
    )

    parser.add_argument(
        "--dedup-stats",
        action="store_true",
        help="重複排除の統計情報を表示"
    )

    return parser.parse_args()
```

---

## ファイル構成

```
download_woc_materials.py       # メインスクリプト（すべてのクラスを含む）
または
src/
├── __init__.py
├── main.py                     # エントリーポイント
├── manager.py                  # WOCDownloadManager
├── download_db.py              # DownloadDB
├── url_dedup.py                # URLDedup
├── executor.py                 # DownloadExecutor
├── filename.py                 # FileNameGenerator
├── logger.py                   # Logger
├── models.py                   # データクラス
└── utils.py                    # ユーティリティ関数
```

---

## 実装の優先順位

1. **Phase 1**: 基本機能
   - WOCDownloadManager（基本構造）
   - DownloadExecutor（yt-dlp, gdown実行）
   - FileNameGenerator
   - Logger

2. **Phase 2**: Resume機能
   - DownloadDB
   - 状態管理ロジック

3. **Phase 3**: Dedup機能
   - URLDedup
   - シンボリックリンク作成

4. **Phase 4**: 高度な機能
   - 並列処理
   - 統計レポート
   - エラーリトライ
