"""ダウンロード状態の管理（Resume機能）"""

import json
import os
import uuid
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .models import DownloadRecord, DownloadStatistics


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
