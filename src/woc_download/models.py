"""データモデルとEnum定義"""

from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, List, Optional

from rich.console import Console
from rich.table import Table


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


@dataclass
class DownloadTask:
    """ダウンロードタスク"""
    url: str
    file_path: str
    sheet_name: str
    row_index: int
    column_name: str
    url_type: URLType


@dataclass
class DownloadResult:
    """ダウンロード結果"""
    success: bool
    file_path: Optional[str]
    file_size: Optional[int]
    error_message: Optional[str]
    skipped: bool = False
    deduped: bool = False
    original_file_path: Optional[str] = None


@dataclass
class DownloadRecord:
    """ダウンロードレコード（Resume用）"""
    id: str
    url: str
    file_path: str
    status: str  # "completed"|"failed"|"in_progress"
    file_size: Optional[int]
    downloaded_at: Optional[str]
    error_message: Optional[str]
    sheet_name: str
    row_index: int
    column_name: str


@dataclass
class DedupReference:
    """重複参照"""
    file_path: str
    link_type: str
    created_at: str


@dataclass
class DedupRecord:
    """重複排除レコード"""
    url: str
    original_file_path: str
    file_size: int
    downloaded_at: str
    references: List[Dict[str, str]]


@dataclass
class DownloadStatistics:
    """ダウンロード統計"""
    total_downloads: int
    completed: int
    failed: int
    skipped: int
    in_progress: int


@dataclass
class DedupStatistics:
    """重複排除統計"""
    total_unique_urls: int
    total_references: int
    space_saved_bytes: int


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


@dataclass
class DownloadReport:
    """全体レポート"""
    sheets: List[SheetReport]
    total_tasks: int
    completed: int
    failed: int
    skipped: int
    deduped: int
    execution_time: float

    def print_summary(self, console: Console):
        """サマリーを出力"""
        console.print("\n[bold cyan]===== ダウンロード完了レポート =====[/bold cyan]\n")

        # 全体統計
        table = Table(title="全体統計")
        table.add_column("項目", style="cyan")
        table.add_column("件数", style="magenta", justify="right")

        table.add_row("総タスク数", str(self.total_tasks))
        table.add_row("✓ 完了", f"[green]{self.completed}[/green]")
        table.add_row("⊗ 失敗", f"[red]{self.failed}[/red]")
        table.add_row("⊘ スキップ", f"[yellow]{self.skipped}[/yellow]")
        table.add_row("🔗 重複排除", f"[blue]{self.deduped}[/blue]")
        table.add_row("実行時間", f"{self.execution_time:.2f}秒")

        console.print(table)

        # シート別統計
        if self.sheets:
            console.print("\n[bold]シート別統計:[/bold]")
            sheet_table = Table()
            sheet_table.add_column("シート名", style="cyan")
            sheet_table.add_column("完了", style="green", justify="right")
            sheet_table.add_column("失敗", style="red", justify="right")
            sheet_table.add_column("スキップ", style="yellow", justify="right")
            sheet_table.add_column("重複", style="blue", justify="right")

            for sheet in self.sheets:
                sheet_table.add_row(
                    sheet.sheet_name,
                    str(sheet.completed),
                    str(sheet.failed),
                    str(sheet.skipped),
                    str(sheet.deduped)
                )

            console.print(sheet_table)
