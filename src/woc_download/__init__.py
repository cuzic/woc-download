"""
WOC講義資料ダウンロードシステム

Excelファイルに記載された講義録画URLと資料URLから、
動画と資料を自動的にダウンロードし、適切なファイル名で保存します。

Features:
- Resume機能: 中断しても続きから再開
- Dedup機能: 重複URLを自動検出して容量節約
- 進捗表示: ダウンロード状況をリアルタイム表示
"""

from .manager import WOCDownloadManager
from .cli import main

__version__ = "1.0.0"
__author__ = "Your Name"
__license__ = "MIT"

__all__ = ["WOCDownloadManager", "main"]
