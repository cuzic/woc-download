"""実際のダウンロード処理の実行"""

import os
from pathlib import Path
from typing import Optional

from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import yt_dlp
import gdown

from .models import URLType, DownloadTask, DownloadResult
from .logger import Logger
from .utils import format_file_size


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
        yt-dlpを使って動画の字幕をダウンロード

        Args:
            url: 動画URL
            output_path: 出力パス（拡張子なし）

        Returns:
            DownloadResult: ダウンロード結果
        """
        if self.dry_run:
            self.logger.console_print(f"[DRY RUN] Would download subtitles: {url}", "yellow")
            return DownloadResult(
                success=True,
                file_path=output_path + ".ja.srt",
                file_size=0,
                error_message=None
            )

        try:
            ydl_opts = {
                'writesubtitles': True,          # 字幕をダウンロード
                'writeautomaticsub': True,       # 自動生成字幕をダウンロード
                'subtitleslangs': ['ja', 'en'],  # 日本語と英語の字幕
                'skip_download': True,           # 動画はダウンロードしない
                'outtmpl': output_path + '.%(ext)s',
                'quiet': True,
                'no_warnings': True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])

            # ダウンロードされた字幕ファイルを探す
            parent_dir = Path(output_path).parent
            base_name = Path(output_path).name

            # .ja.srt, .en.srt, .ja.vtt, .en.vtt などのファイルを探す
            subtitle_patterns = [
                f"{base_name}.ja.*",
                f"{base_name}.en.*",
                f"{base_name}.*.*"  # その他の字幕ファイル
            ]

            downloaded_files = []
            for pattern in subtitle_patterns:
                downloaded_files.extend(list(parent_dir.glob(pattern)))

            # 重複を除去
            downloaded_files = list(set(downloaded_files))

            if downloaded_files:
                # 最初のファイルをメインとして扱う
                file_path = str(downloaded_files[0])
                total_size = sum(self.get_file_size(str(f)) for f in downloaded_files)

                subtitle_list = ', '.join([f.name for f in downloaded_files])
                self.logger.info(
                    f"Downloaded subtitles: {subtitle_list} ({format_file_size(total_size)})",
                    "success"
                )

                return DownloadResult(
                    success=True,
                    file_path=file_path,
                    file_size=total_size,
                    error_message=None
                )
            else:
                raise FileNotFoundError("Subtitle files not found")

        except Exception as e:
            error_msg = f"Failed to download subtitles {url}: {str(e)}"
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
