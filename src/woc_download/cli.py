"""コマンドラインインターフェース"""

import argparse

from .manager import WOCDownloadManager


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
