"""実際のダウンロード処理の実行"""

import os
import re
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

    @staticmethod
    def extract_utage_m3u8_url(url: str) -> Optional[str]:
        """
        utage-system の URL から m3u8 URL を抽出

        Args:
            url: utage-system の動画URL

        Returns:
            m3u8 URL、取得できない場合は None
        """
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()

                # ページを開く
                page.goto(url, wait_until='domcontentloaded', timeout=60000)

                # HTML を取得
                html = page.content()
                browser.close()

                # config オブジェクトから m3u8 URL を抽出
                # Looking for: const config = {...src: "https://...video.m3u8"...};
                match = re.search(r'src:\s*"([^"]+\.m3u8)"', html)

                if match:
                    m3u8_url = match.group(1)
                    return m3u8_url

                return None

        except ImportError:
            raise ImportError("playwright is required for utage-system downloads. Install with: pip install playwright && playwright install chromium")
        except Exception as e:
            return None

    def transcribe_video_with_whisper(self, video_url: str, output_path: str) -> DownloadResult:
        """
        OpenAI Whisper API を使って動画から字幕を生成

        Args:
            video_url: 動画URL（m3u8 または mp4）
            output_path: 出力パス（拡張子なし）

        Returns:
            DownloadResult: 文字起こし結果
        """
        temp_video = None
        temp_audio = None

        try:
            from openai import OpenAI

            # OpenAI API キーを環境変数から取得
            api_key = os.environ.get("OPENAI_API_KEY")
            if not api_key:
                raise ValueError("OPENAI_API_KEY environment variable is not set")

            client = OpenAI(api_key=api_key)

            # 一時的に動画をダウンロード
            temp_video = output_path + "_temp.mp4"
            temp_audio = output_path + "_temp.mp3"

            self.logger.info(f"Downloading video for transcription: {video_url}", "info")

            # yt-dlp で音声のみをダウンロード（軽量化のため）
            ydl_opts = {
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '192',
                }],
                'outtmpl': output_path + "_temp",
                'quiet': True,
                'no_warnings': True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([video_url])

            if not os.path.exists(temp_audio):
                raise FileNotFoundError(f"Failed to download audio from: {video_url}")

            # ファイルサイズをチェック（OpenAI API の制限: 25MB）
            file_size_mb = os.path.getsize(temp_audio) / (1024 * 1024)
            if file_size_mb > 25:
                self.logger.warning(f"Audio file is {file_size_mb:.2f}MB, which exceeds OpenAI's 25MB limit. Consider splitting the file.", "warning")

            self.logger.info(f"Transcribing audio with OpenAI Whisper API ({file_size_mb:.2f}MB)...", "info")

            # OpenAI Whisper API で文字起こし
            with open(temp_audio, "rb") as audio_file:
                transcript = client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    language="ja",  # 日本語
                    response_format="verbose_json",  # タイムスタンプ付き
                    timestamp_granularities=["segment"]
                )

            self.logger.info(f"Detected language: {transcript.language}", "success")

            # SRT 形式で保存
            srt_path = output_path + ".ja.srt"
            with open(srt_path, 'w', encoding='utf-8') as f:
                if hasattr(transcript, 'segments') and transcript.segments:
                    for i, segment in enumerate(transcript.segments, start=1):
                        start_time = self._format_timestamp(segment.start)
                        end_time = self._format_timestamp(segment.end)
                        f.write(f"{i}\n")
                        f.write(f"{start_time} --> {end_time}\n")
                        f.write(f"{segment.text.strip()}\n\n")
                else:
                    # segments がない場合は、全体のテキストを1つのエントリとして保存
                    f.write("1\n")
                    f.write("00:00:00,000 --> 99:99:99,999\n")
                    f.write(f"{transcript.text}\n\n")

            # 一時ファイルを削除
            if temp_video and os.path.exists(temp_video):
                os.remove(temp_video)
            if temp_audio and os.path.exists(temp_audio):
                os.remove(temp_audio)

            file_size = self.get_file_size(srt_path)
            self.logger.info(
                f"Transcription completed: {srt_path} ({format_file_size(file_size)})",
                "success"
            )

            return DownloadResult(
                success=True,
                file_path=srt_path,
                file_size=file_size,
                error_message=None
            )

        except ImportError:
            error_msg = "openai package is required for transcription. Install with: pip install openai"
            self.logger.error(error_msg)
            return DownloadResult(
                success=False,
                file_path=None,
                file_size=None,
                error_message=error_msg
            )
        except Exception as e:
            error_msg = f"Failed to transcribe video: {str(e)}"
            self.logger.error(error_msg)

            # クリーンアップ
            if temp_video and os.path.exists(temp_video):
                os.remove(temp_video)
            if temp_audio and os.path.exists(temp_audio):
                os.remove(temp_audio)

            return DownloadResult(
                success=False,
                file_path=None,
                file_size=None,
                error_message=error_msg
            )

    @staticmethod
    def _format_timestamp(seconds: float) -> str:
        """秒数を SRT タイムスタンプ形式に変換"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        millis = int((seconds % 1) * 1000)
        return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"

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
            # utage-system の場合は m3u8 URL を抽出して whisper で文字起こし
            original_url = url
            if 'utage-system.com' in url.lower():
                self.logger.info(f"Extracting m3u8 URL from utage-system: {url}", "info")
                m3u8_url = self.extract_utage_m3u8_url(url)
                if m3u8_url:
                    self.logger.info(f"Found m3u8 URL: {m3u8_url}", "success")
                    # utage-system には字幕がないので、whisper で文字起こし
                    self.logger.info("utage-system videos don't have subtitles. Using OpenAI Whisper API for transcription...", "info")
                    return self.transcribe_video_with_whisper(m3u8_url, output_path)
                else:
                    raise ValueError(f"Failed to extract m3u8 URL from {original_url}")

            # Cookie ファイルのパスを確認（現在の作業ディレクトリから探す）
            base_path = Path.cwd()

            # URL に応じて適切な cookie ファイルを選択
            cookie_path = None
            if 'youtube.com' in original_url.lower() or 'youtu.be' in original_url.lower():
                youtube_cookies = base_path / "youtube_cookies.txt"
                if youtube_cookies.exists():
                    cookie_path = youtube_cookies
            elif 'vimeo.com' in original_url.lower():
                vimeo_cookies = base_path / "vimeo_cookies.txt"
                if vimeo_cookies.exists():
                    cookie_path = vimeo_cookies

            ydl_opts = {
                'writesubtitles': True,          # 字幕をダウンロード
                'writeautomaticsub': True,       # 自動生成字幕をダウンロード
                'subtitleslangs': ['ja-x-autogen', 'ja', 'en'],  # 日本語と英語の字幕（Vimeo対応）
                'skip_download': True,           # 動画はダウンロードしない
                'outtmpl': output_path + '.%(ext)s',
                'quiet': True,
                'no_warnings': True,
            }

            # Cookie ファイルが存在する場合は使用
            if cookie_path:
                ydl_opts['cookiefile'] = str(cookie_path)

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
