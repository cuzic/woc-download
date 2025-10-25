"""ログ出力の管理"""

import logging
from pathlib import Path
from typing import Dict

from rich.console import Console


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
