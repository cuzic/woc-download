"""ユーティリティ関数"""

import re
from typing import Tuple


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
