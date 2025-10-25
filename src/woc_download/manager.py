"""システム全体の統括"""

import sys
import time
from pathlib import Path
from typing import List, Optional

import pandas as pd
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.table import Table

from .models import (
    DownloadTask, DownloadResult, SheetReport, DownloadReport
)
from .logger import Logger
from .executor import DownloadExecutor
from .database import DownloadDB
from .dedup import URLDedup
from .filename import FileNameGenerator
from .utils import format_file_size


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
