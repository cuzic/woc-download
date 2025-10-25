"""ファイル名の生成とサニタイズ"""

import re
import pandas as pd

from .utils import parse_japanese_date


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
