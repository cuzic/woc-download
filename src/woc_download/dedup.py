"""URL重複排除管理（Dedup機能）"""

import hashlib
import json
import os
import shutil
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

from .models import DedupRecord, DedupStatistics


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
