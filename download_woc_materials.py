#!/usr/bin/env python3
"""
WOC講義資料ダウンロードシステム

Excelファイルに記載された講義録画URLと資料URLから、
動画と資料を自動的にダウンロードし、適切なファイル名で保存します。

Features:
- Resume機能: 中断しても続きから再開
- Dedup機能: 重複URLを自動検出して容量節約
- 進捗表示: ダウンロード状況をリアルタイム表示
"""

import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import sys
import time
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.table import Table
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import yt_dlp
import gdown


# ============================================================================
# Enum定義
# ============================================================================

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


# ============================================================================
# データクラス
# ============================================================================

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


# ============================================================================
# ユーティリティ関数
# ============================================================================

def format_file_size(size_bytes: int) -> str:
    """ファイルサイズを人間が読みやすい形式にフォーマット"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} PB"


def parse_japanese_date(year_str: str, date_str: str) -> Tuple[str, str, str]:
    """
    日本語の日付文字列をパース

    Args:
        year_str: 年の文字列（例: "2025年"）
        date_str: 日付の文字列（例: "5月15日" または "2025-05-15"）

    Returns:
        Tuple[str, str, str]: (年, 月, 日) の4桁・2桁文字列
    """
    # 年をパース
    year = re.search(r'(\d{4})', str(year_str))
    year = year.group(1) if year else "0000"

    # 日付をパース
    # パターン1: "5月15日"
    match = re.search(r'(\d{1,2})月(\d{1,2})日', str(date_str))
    if match:
        month = match.group(1).zfill(2)
        day = match.group(2).zfill(2)
        return year, month, day

    # パターン2: "2025-05-15"
    match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', str(date_str))
    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)
        day = match.group(3).zfill(2)
        return year, month, day

    return year, "00", "00"


# ============================================================================
# Logger クラス
# ============================================================================

class Logger:
    """ログ出力の管理"""

    def __init__(self, log_dir: str = "logs"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.console = Console()

        # ログファイルの設定
        self.loggers: Dict[str, logging.Logger] = {}
        self._setup_loggers()

    def _setup_loggers(self):
        """ロガーのセットアップ"""
        log_files = {
            'success': 'download_success.log',
            'error': 'download_error.log',
            'skip': 'download_skip.log',
            'resume': 'download_resume.log',
            'dedup': 'download_dedup.log',
        }

        for name, filename in log_files.items():
            logger = logging.getLogger(f'woc.{name}')
            logger.setLevel(logging.INFO)

            # ファイルハンドラ
            fh = logging.FileHandler(self.log_dir / filename, encoding='utf-8')
            fh.setLevel(logging.INFO)

            # フォーマッタ
            formatter = logging.Formatter(
                '[%(asctime)s] [%(levelname)s] %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            fh.setFormatter(formatter)

            logger.addHandler(fh)
            self.loggers[name] = logger

    def info(self, message: str, category: str = "success"):
        """INFOレベルのログを出力"""
        if category in self.loggers:
            self.loggers[category].info(message)

    def error(self, message: str):
        """ERRORレベルのログを出力"""
        self.loggers['error'].error(message)

    def console_print(self, message: str, style: str = ""):
        """コンソールに出力"""
        if style:
            self.console.print(f"[{style}]{message}[/{style}]")
        else:
            self.console.print(message)


# ============================================================================
# FileNameGenerator クラス
# ============================================================================

class FileNameGenerator:
    """ファイル名の生成とサニタイズ"""

    @staticmethod
    def sanitize_filename(filename: str, max_length: int = 200) -> str:
        """
        ファイル名をサニタイズ

        Args:
            filename: 元のファイル名
            max_length: 最大文字数

        Returns:
            str: サニタイズされたファイル名
        """
        # 禁止文字を全角に置換
        replacements = {
            '/': '／',
            '\\': '＼',
            ':': '：',
            '*': '＊',
            '?': '？',
            '"': '"',
            '<': '＜',
            '>': '＞',
            '|': '｜',
        }

        for char, replacement in replacements.items():
            filename = filename.replace(char, replacement)

        # 改行・タブ・連続空白を除去
        filename = re.sub(r'[\n\r\t]+', '', filename)
        filename = re.sub(r'\s+', ' ', filename)
        filename = filename.strip()

        # 最大長を制限
        if len(filename) > max_length:
            filename = filename[:max_length]

        return filename

    @staticmethod
    def extract_chapter_number(title: str) -> str:
        """
        タイトルから章番号を抽出

        Args:
            title: コンテンツタイトル

        Returns:
            str: 章番号（例: "1-1", "2-3"）
        """
        match = re.search(r'(\d+-\d+|\d+)', str(title))
        return match.group(1) if match else ""

    @staticmethod
    def generate_filename(sheet_name: str, row: pd.Series, column_name: str) -> str:
        """
        ファイル名を生成（拡張子なし）

        Args:
            sheet_name: シート名
            row: 行データ
            column_name: 列名

        Returns:
            str: ファイル名（拡張子なし、サニタイズ済み）
        """
        if sheet_name == "コンテンツ":
            # コンテンツシート
            title = str(row.get('コンテンツタイトル', ''))
            chapter = FileNameGenerator.extract_chapter_number(title)

            # 章番号を除去したタイトル
            clean_title = re.sub(r'^\d+-\d+\.?|^\d+\.?', '', title).strip()

            # ファイル種別
            if 'DL' in column_name or '動画' in column_name:
                suffix = "video"
            elif '資料1' in column_name:
                suffix = "資料1"
            elif '資料2' in column_name:
                suffix = "資料2"
            else:
                suffix = "file"

            if chapter:
                filename = f"{chapter}_{clean_title}_{suffix}"
            else:
                filename = f"{clean_title}_{suffix}"

        else:
            # 講義録画・資料、グルコン、合宿シート
            year_str = str(row.get('実施年', ''))
            date_str = str(row.get('実施日', ''))
            event_type = str(row.get('開催種別', ''))
            title = str(row.get('講義タイトル', ''))

            # 日付をパース
            year, month, day = parse_japanese_date(year_str, date_str)
            date_prefix = f"{year}{month}{day}"

            # ファイル種別
            if '視聴' in column_name:
                suffix = "video_view"
            elif 'DL' in column_name or '動画' in column_name:
                suffix = "video"
            elif '資料1' in column_name:
                suffix = "資料1"
            elif '資料2' in column_name:
                suffix = "資料2"
            elif '資料3' in column_name:
                suffix = "資料3"
            elif '資料4' in column_name:
                suffix = "資料4"
            else:
                suffix = "file"

            # ファイル名生成
            parts = [date_prefix, event_type, title, suffix]
            parts = [p for p in parts if p and p != 'nan']
            filename = "_".join(parts)

        return FileNameGenerator.sanitize_filename(filename)


# ============================================================================
# DownloadExecutor クラス
# ============================================================================

class DownloadExecutor:
    """実際のダウンロード処理の実行"""

    def __init__(self, dry_run: bool = False, logger: Optional[Logger] = None):
        self.dry_run = dry_run
        self.logger = logger or Logger()

    @staticmethod
    def detect_url_type(url: str) -> URLType:
        """URLの種類を判定"""
        url_lower = url.lower()

        if 'vimeo.com' in url_lower:
            return URLType.VIMEO
        elif 'youtube.com' in url_lower or 'youtu.be' in url_lower:
            return URLType.YOUTUBE
        elif 'loom.com' in url_lower:
            return URLType.LOOM
        elif 'utage-system.com' in url_lower:
            return URLType.UTAGE
        elif '.m3u8' in url_lower:
            return URLType.M3U8
        elif 'docs.google.com/presentation' in url_lower:
            return URLType.GOOGLE_SLIDES
        elif 'docs.google.com/spreadsheets' in url_lower:
            return URLType.GOOGLE_SHEETS
        elif 'docs.google.com/document' in url_lower:
            return URLType.GOOGLE_DOCS
        elif 'drive.google.com/file' in url_lower:
            return URLType.GOOGLE_DRIVE_FILE
        elif 'drive.google.com/drive/folders' in url_lower:
            return URLType.GOOGLE_DRIVE_FOLDER
        else:
            return URLType.UNKNOWN

    @staticmethod
    def get_file_size(file_path: str) -> int:
        """ファイルサイズを取得"""
        try:
            return os.path.getsize(file_path)
        except:
            return 0

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((ConnectionError, TimeoutError))
    )
    def download_video(self, url: str, output_path: str) -> DownloadResult:
        """
        yt-dlpを使って動画をダウンロード

        Args:
            url: 動画URL
            output_path: 出力パス（拡張子なし）

        Returns:
            DownloadResult: ダウンロード結果
        """
        if self.dry_run:
            self.logger.console_print(f"[DRY RUN] Would download video: {url}", "yellow")
            return DownloadResult(
                success=True,
                file_path=output_path + ".mp4",
                file_size=0,
                error_message=None
            )

        try:
            ydl_opts = {
                'format': 'bestvideo+bestaudio/best',
                'outtmpl': output_path + '.%(ext)s',
                'quiet': True,
                'no_warnings': True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])

            # ダウンロードされたファイルを探す
            parent_dir = Path(output_path).parent
            base_name = Path(output_path).name

            downloaded_files = list(parent_dir.glob(f"{base_name}.*"))
            if downloaded_files:
                file_path = str(downloaded_files[0])
                file_size = self.get_file_size(file_path)

                self.logger.info(
                    f"Downloaded video: {file_path} ({format_file_size(file_size)})",
                    "success"
                )

                return DownloadResult(
                    success=True,
                    file_path=file_path,
                    file_size=file_size,
                    error_message=None
                )
            else:
                raise FileNotFoundError("Downloaded file not found")

        except Exception as e:
            error_msg = f"Failed to download video {url}: {str(e)}"
            self.logger.error(error_msg)
            return DownloadResult(
                success=False,
                file_path=None,
                file_size=None,
                error_message=error_msg
            )

    def download_document(self, url: str, output_path: str, url_type: URLType) -> DownloadResult:
        """
        gdownを使って資料をダウンロード

        Args:
            url: 資料URL
            output_path: 出力パス（拡張子なし）
            url_type: URL種別

        Returns:
            DownloadResult: ダウンロード結果
        """
        if self.dry_run:
            self.logger.console_print(f"[DRY RUN] Would download document: {url}", "yellow")
            return DownloadResult(
                success=True,
                file_path=output_path + ".pdf",
                file_size=0,
                error_message=None
            )

        try:
            # URL種別に応じて拡張子を決定
            if url_type == URLType.GOOGLE_SLIDES or url_type == URLType.GOOGLE_DOCS:
                ext = ".pdf"
                format_type = "pdf"
            elif url_type == URLType.GOOGLE_SHEETS:
                ext = ".xlsx"
                format_type = "xlsx"
            else:
                ext = ""
                format_type = None

            output_file = output_path + ext

            # フォルダの場合
            if url_type == URLType.GOOGLE_DRIVE_FOLDER:
                folder_path = output_path + "_folder"
                os.makedirs(folder_path, exist_ok=True)
                gdown.download_folder(url, output=folder_path, quiet=True)

                # フォルダサイズを計算
                total_size = sum(
                    f.stat().st_size for f in Path(folder_path).rglob('*') if f.is_file()
                )

                self.logger.info(
                    f"Downloaded folder: {folder_path} ({format_file_size(total_size)})",
                    "success"
                )

                return DownloadResult(
                    success=True,
                    file_path=folder_path,
                    file_size=total_size,
                    error_message=None
                )

            # ファイルの場合
            if format_type:
                gdown.download(url, output=output_file, quiet=True, fuzzy=True)
            else:
                gdown.download(url, output=output_file, quiet=True, fuzzy=True)

            if os.path.exists(output_file):
                file_size = self.get_file_size(output_file)
                self.logger.info(
                    f"Downloaded document: {output_file} ({format_file_size(file_size)})",
                    "success"
                )

                return DownloadResult(
                    success=True,
                    file_path=output_file,
                    file_size=file_size,
                    error_message=None
                )
            else:
                raise FileNotFoundError("Downloaded file not found")

        except Exception as e:
            error_msg = f"Failed to download document {url}: {str(e)}"
            self.logger.error(error_msg)
            return DownloadResult(
                success=False,
                file_path=None,
                file_size=None,
                error_message=error_msg
            )

    def download(self, task: DownloadTask) -> DownloadResult:
        """
        ダウンロードタスクを実行

        Args:
            task: ダウンロードタスク

        Returns:
            DownloadResult: ダウンロード結果
        """
        # ディレクトリ作成
        Path(task.file_path).parent.mkdir(parents=True, exist_ok=True)

        # URL種別に応じてダウンロード
        if task.url_type in [URLType.VIMEO, URLType.YOUTUBE, URLType.LOOM,
                             URLType.UTAGE, URLType.M3U8]:
            return self.download_video(task.url, task.file_path)
        else:
            return self.download_document(task.url, task.file_path, task.url_type)


# ============================================================================
# DownloadDB クラス（Resume機能）
# ============================================================================

class DownloadDB:
    """ダウンロード状態の管理"""

    def __init__(self, db_path: str):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.data: Dict[str, Any] = {
            'downloads': [],
            'metadata': {
                'last_updated': None,
                'total_downloads': 0,
                'completed': 0,
                'failed': 0,
                'skipped': 0
            }
        }
        self.downloads: Dict[str, DownloadRecord] = {}
        self.load()

    def load(self):
        """DBファイルを読み込む"""
        if self.db_path.exists():
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)

                # downloadsを辞書に変換
                for dl in self.data.get('downloads', []):
                    record = DownloadRecord(**dl)
                    self.downloads[record.url] = record
            except Exception as e:
                print(f"Warning: Failed to load DB: {e}")

    def save(self):
        """DBファイルに保存"""
        self.data['downloads'] = [asdict(r) for r in self.downloads.values()]
        self.data['metadata']['last_updated'] = datetime.now().isoformat()

        with open(self.db_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get_record(self, url: str) -> Optional[DownloadRecord]:
        """URLに対応するレコードを取得"""
        return self.downloads.get(url)

    def is_completed(self, url: str, file_path: str) -> bool:
        """ダウンロードが完了しているかチェック"""
        record = self.get_record(url)
        if not record:
            return False

        if record.status != 'completed':
            return False

        # ファイルの存在チェック
        if not os.path.exists(file_path):
            # 拡張子が異なる可能性があるので、ベース名で検索
            parent = Path(file_path).parent
            base_name = Path(file_path).stem
            matching_files = list(parent.glob(f"{base_name}.*"))
            if not matching_files:
                return False
            file_path = str(matching_files[0])

        # ファイルサイズチェック
        file_size = os.path.getsize(file_path)
        return file_size > 0

    def mark_completed(self, url: str, file_path: str, file_size: int,
                      sheet_name: str = "", row_index: int = 0, column_name: str = ""):
        """ダウンロード完了をマーク"""
        record = self.downloads.get(url)
        if not record:
            record = DownloadRecord(
                id=str(uuid.uuid4()),
                url=url,
                file_path=file_path,
                status='completed',
                file_size=file_size,
                downloaded_at=datetime.now().isoformat(),
                error_message=None,
                sheet_name=sheet_name,
                row_index=row_index,
                column_name=column_name
            )
        else:
            record.status = 'completed'
            record.file_size = file_size
            record.downloaded_at = datetime.now().isoformat()
            record.error_message = None

        self.downloads[url] = record
        self.save()

    def mark_failed(self, url: str, file_path: str, error_message: str,
                   sheet_name: str = "", row_index: int = 0, column_name: str = ""):
        """ダウンロード失敗をマーク"""
        record = self.downloads.get(url)
        if not record:
            record = DownloadRecord(
                id=str(uuid.uuid4()),
                url=url,
                file_path=file_path,
                status='failed',
                file_size=None,
                downloaded_at=None,
                error_message=error_message,
                sheet_name=sheet_name,
                row_index=row_index,
                column_name=column_name
            )
        else:
            record.status = 'failed'
            record.error_message = error_message

        self.downloads[url] = record
        self.save()

    def get_failed_records(self) -> List[DownloadRecord]:
        """失敗したレコードの一覧を取得"""
        return [r for r in self.downloads.values() if r.status == 'failed']

    def get_statistics(self) -> DownloadStatistics:
        """統計情報を取得"""
        total = len(self.downloads)
        completed = sum(1 for r in self.downloads.values() if r.status == 'completed')
        failed = sum(1 for r in self.downloads.values() if r.status == 'failed')
        in_progress = sum(1 for r in self.downloads.values() if r.status == 'in_progress')

        return DownloadStatistics(
            total_downloads=total,
            completed=completed,
            failed=failed,
            skipped=0,
            in_progress=in_progress
        )

    def reset(self):
        """すべてのレコードをクリア"""
        self.downloads.clear()
        self.data['downloads'] = []
        self.save()


# ============================================================================
# URLDedup クラス（重複排除機能）
# ============================================================================

class URLDedup:
    """URL重複排除管理"""

    def __init__(self, db_path: str, dedup_mode: str = "symlink"):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.dedup_mode = dedup_mode
        self.data: Dict[str, Any] = {
            'url_hash_to_file': {},
            'metadata': {
                'total_unique_urls': 0,
                'total_references': 0,
                'space_saved_bytes': 0
            }
        }
        self.url_hash_to_file: Dict[str, DedupRecord] = {}
        self.load()

    def load(self):
        """DBファイルを読み込む"""
        if self.db_path.exists():
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)

                # url_hash_to_fileを復元
                for hash_key, record_dict in self.data.get('url_hash_to_file', {}).items():
                    self.url_hash_to_file[hash_key] = DedupRecord(**record_dict)
            except Exception as e:
                print(f"Warning: Failed to load Dedup DB: {e}")

    def save(self):
        """DBファイルに保存"""
        self.data['url_hash_to_file'] = {
            k: asdict(v) for k, v in self.url_hash_to_file.items()
        }

        with open(self.db_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    @staticmethod
    def normalize_url(url: str) -> str:
        """URLを正規化"""
        parsed = urlparse(url)
        normalized = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        return normalized.rstrip('/')

    @staticmethod
    def get_url_hash(url: str) -> str:
        """URLのハッシュを生成"""
        normalized = URLDedup.normalize_url(url)
        return hashlib.sha256(normalized.encode('utf-8')).hexdigest()

    def is_duplicate(self, url: str) -> Tuple[bool, Optional[str]]:
        """
        URLが重複しているかチェック

        Returns:
            Tuple[bool, Optional[str]]: (重複フラグ, 元ファイルパス)
        """
        url_hash = self.get_url_hash(url)

        if url_hash not in self.url_hash_to_file:
            return False, None

        record = self.url_hash_to_file[url_hash]

        # 元ファイルが存在するか確認
        if os.path.exists(record.original_file_path):
            return True, record.original_file_path

        # 元ファイルが存在しない場合は重複とみなさない
        del self.url_hash_to_file[url_hash]
        self.save()
        return False, None

    def register(self, url: str, file_path: str, file_size: int):
        """新規ダウンロードを登録"""
        url_hash = self.get_url_hash(url)

        record = DedupRecord(
            url=url,
            original_file_path=file_path,
            file_size=file_size,
            downloaded_at=datetime.now().isoformat(),
            references=[]
        )

        self.url_hash_to_file[url_hash] = record
        self.save()

    def add_reference(self, url: str, new_file_path: str) -> str:
        """
        参照を追加

        Returns:
            str: 元ファイルのパス
        """
        url_hash = self.get_url_hash(url)

        if url_hash not in self.url_hash_to_file:
            raise ValueError(f"URL not registered: {url}")

        record = self.url_hash_to_file[url_hash]

        # 参照を追加
        reference = {
            'file_path': new_file_path,
            'link_type': self.dedup_mode,
            'created_at': datetime.now().isoformat()
        }
        record.references.append(reference)

        self.save()
        return record.original_file_path

    def create_link(self, original_path: str, new_path: str):
        """ファイルリンクを作成"""
        # ディレクトリ作成
        Path(new_path).parent.mkdir(parents=True, exist_ok=True)

        if self.dedup_mode == "symlink":
            # シンボリックリンク作成
            os.symlink(os.path.abspath(original_path), new_path)
        elif self.dedup_mode == "copy":
            # ファイルコピー
            shutil.copy2(original_path, new_path)
        # "reference"の場合は何もしない

    def get_statistics(self) -> DedupStatistics:
        """重複排除の統計情報を取得"""
        total_unique = len(self.url_hash_to_file)
        total_refs = sum(len(r.references) for r in self.url_hash_to_file.values())
        space_saved = sum(
            r.file_size * len(r.references)
            for r in self.url_hash_to_file.values()
            if self.dedup_mode == "symlink"
        )

        return DedupStatistics(
            total_unique_urls=total_unique,
            total_references=total_refs,
            space_saved_bytes=space_saved
        )

    def get_top_duplicates(self, n: int = 5) -> List[Tuple[str, int]]:
        """重複が多いURLのトップN件を取得"""
        items = [
            (r.url, len(r.references))
            for r in self.url_hash_to_file.values()
        ]
        items.sort(key=lambda x: x[1], reverse=True)
        return items[:n]


# ============================================================================
# WOCDownloadManager クラス
# ============================================================================

class WOCDownloadManager:
    """システム全体の統括"""

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
        self.excel_path = excel_path
        self.download_dir = Path(download_dir)
        self.state_dir = Path(state_dir)
        self.dry_run = dry_run
        self.overwrite = overwrite
        self.no_dedup = no_dedup
        self.dedup_mode = dedup_mode
        self.parallel = parallel
        self.target_sheets = target_sheets

        # ディレクトリ作成
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.state_dir.mkdir(parents=True, exist_ok=True)

        # コンポーネント初期化
        self.logger = Logger()
        self.console = Console()
        self.download_db = DownloadDB(str(self.state_dir / "download_db.json"))
        self.url_dedup = URLDedup(
            str(self.state_dir / "url_dedup.json"),
            dedup_mode=dedup_mode
        )
        self.executor = DownloadExecutor(dry_run=dry_run, logger=self.logger)

    def run(self) -> DownloadReport:
        """メイン処理を実行"""
        self.console.print("[bold cyan]WOC講義資料ダウンロードシステム[/bold cyan]")
        self.console.print(f"Excel: {self.excel_path}\n")

        start_time = time.time()

        # Excelファイルを読み込む
        try:
            xl_file = pd.ExcelFile(self.excel_path)
            sheet_names = xl_file.sheet_names
        except Exception as e:
            self.console.print(f"[bold red]Error: Failed to load Excel file: {e}[/bold red]")
            sys.exit(1)

        # 対象シートをフィルタ
        if self.target_sheets:
            sheet_names = [s for s in sheet_names if s in self.target_sheets]

        self.console.print(f"Processing {len(sheet_names)} sheets: {', '.join(sheet_names)}\n")

        # 各シートを処理
        sheet_reports: List[SheetReport] = []

        for sheet_name in sheet_names:
            self.console.print(f"[bold]Processing sheet: {sheet_name}[/bold]")

            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            report = self.process_sheet(sheet_name, df)
            sheet_reports.append(report)

            self.console.print(
                f"  Completed: {report.completed}, Failed: {report.failed}, "
                f"Skipped: {report.skipped}, Deduped: {report.deduped}\n"
            )

        # 全体レポート
        total_tasks = sum(r.total_tasks for r in sheet_reports)
        completed = sum(r.completed for r in sheet_reports)
        failed = sum(r.failed for r in sheet_reports)
        skipped = sum(r.skipped for r in sheet_reports)
        deduped = sum(r.deduped for r in sheet_reports)

        execution_time = time.time() - start_time

        return DownloadReport(
            sheets=sheet_reports,
            total_tasks=total_tasks,
            completed=completed,
            failed=failed,
            skipped=skipped,
            deduped=deduped,
            execution_time=execution_time
        )

    def process_sheet(self, sheet_name: str, df: pd.DataFrame) -> SheetReport:
        """1つのシートを処理"""
        # メタデータ継承用
        previous_year = None
        previous_date = None

        tasks: List[DownloadTask] = []

        # 各行を処理
        for index, row in df.iterrows():
            # 実施年・実施日の継承
            if '実施年' in row:
                if pd.notna(row['実施年']):
                    previous_year = row['実施年']
                elif previous_year:
                    row['実施年'] = previous_year

            if '実施日' in row:
                if pd.notna(row['実施日']):
                    previous_date = row['実施日']
                elif previous_date:
                    row['実施日'] = previous_date

            # タスク生成
            row_tasks = self.process_row(sheet_name, index, row)
            tasks.extend(row_tasks)

        # タスクを実行
        completed = 0
        failed = 0
        skipped = 0
        deduped = 0
        errors: List[str] = []

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            console=self.console
        ) as progress:
            task_id = progress.add_task(f"[cyan]{sheet_name}", total=len(tasks))

            for task in tasks:
                result = self.execute_task(task)

                if result.skipped:
                    skipped += 1
                elif result.deduped:
                    deduped += 1
                elif result.success:
                    completed += 1
                else:
                    failed += 1
                    if result.error_message:
                        errors.append(result.error_message)

                progress.update(task_id, advance=1)

        return SheetReport(
            sheet_name=sheet_name,
            total_tasks=len(tasks),
            completed=completed,
            failed=failed,
            skipped=skipped,
            deduped=deduped,
            errors=errors
        )

    def process_row(self, sheet_name: str, row_index: int, row: pd.Series) -> List[DownloadTask]:
        """1行を処理してダウンロードタスクを生成"""
        tasks: List[DownloadTask] = []

        # URLを含む列を探す
        url_columns = []

        if sheet_name == "コンテンツ":
            url_columns = ['動画リンク', '動画DLリンク', '資料1', '資料2']
        else:
            url_columns = [
                '録画（動画視聴リンク）', '録画（動画DLリンク）',
                '資料1', '資料2', '資料3', '資料4'
            ]

        for col in url_columns:
            if col not in row:
                continue

            url = row[col]

            # URLの有効性チェック
            if pd.isna(url) or str(url).strip() == '' or str(url) == '-':
                continue

            url = str(url).strip()

            # URL種別を判定
            url_type = DownloadExecutor.detect_url_type(url)

            # ファイル名生成
            filename = FileNameGenerator.generate_filename(sheet_name, row, col)

            # 出力パス
            file_path = str(self.download_dir / sheet_name / filename)

            # タスク作成
            task = DownloadTask(
                url=url,
                file_path=file_path,
                sheet_name=sheet_name,
                row_index=int(row_index),
                column_name=col,
                url_type=url_type
            )

            tasks.append(task)

        return tasks

    def execute_task(self, task: DownloadTask) -> DownloadResult:
        """タスクを実行"""
        # Resume機能: 既に完了しているかチェック
        if not self.overwrite and self.download_db.is_completed(task.url, task.file_path):
            self.logger.info(
                f"[RESUME] Skipped (already completed): {task.file_path}",
                "resume"
            )
            return DownloadResult(
                success=True,
                file_path=task.file_path,
                file_size=None,
                error_message=None,
                skipped=True
            )

        # Dedup機能: 重複チェック
        if not self.no_dedup:
            is_dup, original_path = self.url_dedup.is_duplicate(task.url)

            if is_dup and original_path:
                # リンク作成
                try:
                    self.url_dedup.create_link(original_path, task.file_path)
                    self.url_dedup.add_reference(task.url, task.file_path)

                    self.logger.info(
                        f"[DEDUP] Created link: {task.file_path} -> {original_path}",
                        "dedup"
                    )

                    return DownloadResult(
                        success=True,
                        file_path=task.file_path,
                        file_size=0,
                        error_message=None,
                        deduped=True,
                        original_file_path=original_path
                    )
                except Exception as e:
                    self.logger.error(f"Failed to create link: {e}")
                    # リンク作成失敗時は通常ダウンロードへフォールバック

        # ダウンロード実行
        result = self.executor.download(task)

        # 結果を記録
        if result.success and result.file_path:
            self.download_db.mark_completed(
                task.url, result.file_path, result.file_size or 0,
                task.sheet_name, task.row_index, task.column_name
            )

            # Dedupに登録
            if not self.no_dedup and result.file_size:
                self.url_dedup.register(task.url, result.file_path, result.file_size)
        else:
            self.download_db.mark_failed(
                task.url, task.file_path, result.error_message or "Unknown error",
                task.sheet_name, task.row_index, task.column_name
            )

        return result

    def show_status(self):
        """現在のダウンロード状態を表示"""
        stats = self.download_db.get_statistics()

        table = Table(title="ダウンロード状態")
        table.add_column("項目", style="cyan")
        table.add_column("件数", style="magenta", justify="right")

        table.add_row("総ダウンロード数", str(stats.total_downloads))
        table.add_row("✓ 完了", f"[green]{stats.completed}[/green]")
        table.add_row("⊗ 失敗", f"[red]{stats.failed}[/red]")
        table.add_row("⏳ 進行中", f"[yellow]{stats.in_progress}[/yellow]")

        self.console.print(table)

    def show_dedup_stats(self):
        """重複排除の統計情報を表示"""
        stats = self.url_dedup.get_statistics()

        table = Table(title="重複排除統計")
        table.add_column("項目", style="cyan")
        table.add_column("値", style="magenta", justify="right")

        table.add_row("ユニークURL数", str(stats.total_unique_urls))
        table.add_row("参照数", str(stats.total_references))
        table.add_row("節約容量", format_file_size(stats.space_saved_bytes))

        self.console.print(table)

        # Top重複URL
        top_dups = self.url_dedup.get_top_duplicates(5)
        if top_dups:
            self.console.print("\n[bold]重複が多いURL Top 5:[/bold]")
            for i, (url, count) in enumerate(top_dups, 1):
                short_url = url[:60] + "..." if len(url) > 60 else url
                self.console.print(f"  {i}. {short_url} ({count}回使用)")

    def reset(self):
        """ダウンロード状態をリセット"""
        self.download_db.reset()
        self.console.print("[bold green]ダウンロード状態をリセットしました[/bold green]")

    def retry_failed(self) -> DownloadReport:
        """失敗したダウンロードを再試行"""
        failed_records = self.download_db.get_failed_records()

        if not failed_records:
            self.console.print("[bold yellow]失敗したダウンロードはありません[/bold yellow]")
            return DownloadReport(
                sheets=[],
                total_tasks=0,
                completed=0,
                failed=0,
                skipped=0,
                deduped=0,
                execution_time=0
            )

        self.console.print(f"[bold]失敗した{len(failed_records)}件を再試行します[/bold]\n")

        start_time = time.time()
        completed = 0
        failed = 0

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            console=self.console
        ) as progress:
            task_id = progress.add_task("[cyan]Retrying failed downloads", total=len(failed_records))

            for record in failed_records:
                # タスク再作成
                task = DownloadTask(
                    url=record.url,
                    file_path=record.file_path,
                    sheet_name=record.sheet_name,
                    row_index=record.row_index,
                    column_name=record.column_name,
                    url_type=DownloadExecutor.detect_url_type(record.url)
                )

                result = self.executor.download(task)

                if result.success:
                    completed += 1
                    self.download_db.mark_completed(
                        task.url, result.file_path or "", result.file_size or 0,
                        task.sheet_name, task.row_index, task.column_name
                    )
                else:
                    failed += 1

                progress.update(task_id, advance=1)

        execution_time = time.time() - start_time

        return DownloadReport(
            sheets=[],
            total_tasks=len(failed_records),
            completed=completed,
            failed=failed,
            skipped=0,
            deduped=0,
            execution_time=execution_time
        )


# ============================================================================
# メイン処理
# ============================================================================

def parse_arguments() -> argparse.Namespace:
    """コマンドライン引数をパース"""
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


def main():
    """メインエントリーポイント"""
    args = parse_arguments()

    # マネージャー初期化
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
        report.print_summary(manager.console)
    else:
        report = manager.run()
        report.print_summary(manager.console)


if __name__ == "__main__":
    main()
