# 推奨ライブラリ一覧

## 必須ライブラリ

### 1. データ処理

#### pandas
```bash
uv pip install pandas
```
**目的**: Excelファイルの読み込みと処理
**使用箇所**:
- Excelファイルの読み込み (`pd.read_excel()`)
- DataFrameの操作
- NaN値のハンドリング

**主な機能**:
```python
import pandas as pd

# Excelファイル読み込み
df = pd.read_excel('file.xlsx', sheet_name='シート名')

# NaN判定
pd.notna(value)
pd.isna(value)

# 行の反復処理
for index, row in df.iterrows():
    pass
```

#### openpyxl
```bash
uv pip install openpyxl
```
**目的**: Excelファイル（.xlsx）のバックエンド
**使用箇所**:
- pandasがExcelファイルを読み込む際に内部で使用
- 必須依存関係

---

### 2. ダウンロード機能

#### yt-dlp
```bash
uv pip install yt-dlp
```
**目的**: 動画ダウンロード（YouTube, Vimeo, Loom等）
**使用箇所**:
- 動画URLからのダウンロード
- m3u8ストリーミング動画のダウンロード

**主な機能**:
```python
import yt_dlp

ydl_opts = {
    'format': 'bestvideo+bestaudio/best',
    'outtmpl': '%(title)s.%(ext)s',
    'quiet': False,
    'no_warnings': False,
}

with yt_dlp.YoutubeDL(ydl_opts) as ydl:
    ydl.download([url])
```

**対応サイト**:
- YouTube
- Vimeo
- Loom
- m3u8ストリーム
- その他1000以上のサイト

#### gdown
```bash
uv pip install gdown
```
**目的**: Google Drive/Docsからのダウンロード
**使用箇所**:
- Google Driveファイルのダウンロード
- Google Docs/Slides/Sheetsのダウンロード（PDF/XLSX変換）

**主な機能**:
```python
import gdown

# ファイルダウンロード
gdown.download(url, output='file.pdf', quiet=False, fuzzy=True)

# フォルダダウンロード
gdown.download_folder(url, output='folder/', quiet=False)

# Google DocsをPDFで
url = 'https://docs.google.com/document/d/xxxxx/edit'
gdown.download(url, output='doc.pdf', quiet=False, fuzzy=True, format='pdf')

# Google SheetsをXLSXで
url = 'https://docs.google.com/spreadsheets/d/xxxxx/edit'
gdown.download(url, output='sheet.xlsx', quiet=False, fuzzy=True, format='xlsx')
```

**注意事項**:
- 共有設定が「リンクを知っている全員」になっている必要がある
- 大容量ファイルは警告が出る場合がある

---

### 3. UI/UX

#### tqdm
```bash
uv pip install tqdm
```
**目的**: プログレスバーの表示
**使用箇所**:
- ダウンロード進捗の表示
- ファイル処理進捗の表示

**主な機能**:
```python
from tqdm import tqdm
import time

# 基本的な使用
for i in tqdm(range(100), desc="Processing"):
    time.sleep(0.1)

# ファイル数のカウント
files = ['file1', 'file2', 'file3']
for file in tqdm(files, desc="Downloading"):
    download(file)

# 手動更新
pbar = tqdm(total=100, desc="Download")
pbar.update(10)  # 10進める
pbar.close()
```

**表示例**:
```
Processing: 100%|██████████| 100/100 [00:10<00:00,  9.99it/s]
```

#### rich (オプション - 推奨)
```bash
uv pip install rich
```
**目的**: より美しいコンソール出力
**使用箇所**:
- テーブル形式のレポート表示
- カラフルなログ出力
- より高機能なプログレスバー

**主な機能**:
```python
from rich.console import Console
from rich.table import Table
from rich.progress import Progress

console = Console()

# カラフルな出力
console.print("[bold green]Success![/bold green]")
console.print("[bold red]Error![/bold red]")

# テーブル表示
table = Table(title="Download Report")
table.add_column("File", style="cyan")
table.add_column("Status", style="magenta")
table.add_row("file1.mp4", "✓ Completed")
console.print(table)

# プログレスバー
with Progress() as progress:
    task = progress.add_task("[cyan]Downloading...", total=100)
    for i in range(100):
        progress.update(task, advance=1)
```

---

### 4. リトライ・エラーハンドリング

#### tenacity
```bash
uv pip install tenacity
```
**目的**: リトライ処理の実装
**使用箇所**:
- ダウンロード失敗時の自動リトライ
- ネットワークエラー時の指数バックオフ

**主な機能**:
```python
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

@retry(
    stop=stop_after_attempt(3),  # 最大3回リトライ
    wait=wait_exponential(multiplier=1, min=1, max=10),  # 1秒→2秒→4秒
    retry=retry_if_exception_type(ConnectionError)
)
def download_with_retry(url):
    # ダウンロード処理
    pass
```

**リトライ戦略**:
- 最大リトライ回数: 3回
- 待機時間: 指数バックオフ（1秒 → 2秒 → 4秒）
- リトライ対象: ネットワークエラー、タイムアウト

---

### 5. データバリデーション (オプション)

#### pydantic
```bash
uv pip install pydantic
```
**目的**: データクラスのバリデーションとシリアライゼーション
**使用箇所**:
- データクラスの型チェック
- JSONのバリデーション

**主な機能**:
```python
from pydantic import BaseModel, HttpUrl, validator
from typing import Optional
from datetime import datetime

class DownloadRecord(BaseModel):
    url: HttpUrl
    file_path: str
    status: str
    file_size: Optional[int] = None
    downloaded_at: Optional[datetime] = None

    @validator('status')
    def validate_status(cls, v):
        if v not in ['completed', 'failed', 'in_progress']:
            raise ValueError('Invalid status')
        return v

# 使用例
record = DownloadRecord(
    url='https://example.com/video.mp4',
    file_path='/path/to/file.mp4',
    status='completed'
)

# JSONシリアライゼーション
json_str = record.json()
# JSONデシリアライゼーション
record = DownloadRecord.parse_raw(json_str)
```

**メリット**:
- 自動型チェック
- JSON変換が簡単
- バリデーションエラーが分かりやすい

**注意**: 標準ライブラリの`dataclasses`でも十分な場合が多い

---

## 標準ライブラリ（インストール不要）

以下は Python 3.8+ に含まれる標準ライブラリです。

### 1. ファイル・パス操作
```python
from pathlib import Path
import shutil
import os
```

### 2. データ処理
```python
import json
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any, Tuple
from enum import Enum
import re
```

### 3. 日時処理
```python
from datetime import datetime, timedelta
import time
```

### 4. ハッシュ・UUID
```python
import hashlib
import uuid
```

### 5. URL処理
```python
from urllib.parse import urlparse, parse_qs, urlencode
```

### 6. ログ
```python
import logging
from logging.handlers import RotatingFileHandler
```

### 7. コマンドライン
```python
import argparse
import sys
```

### 8. 並列処理
```python
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
```

---

## インストール方法

### 必須パッケージ一括インストール
```bash
uv pip install pandas openpyxl yt-dlp gdown tqdm
```

### 推奨パッケージ追加
```bash
uv pip install rich tenacity
```

### オプション（データバリデーション）
```bash
uv pip install pydantic
```

### pyproject.toml を使用する場合
```toml
[project]
name = "woc-download"
version = "1.0.0"
dependencies = [
    "pandas>=2.0.0",
    "openpyxl>=3.1.0",
    "yt-dlp>=2024.0.0",
    "gdown>=5.0.0",
    "tqdm>=4.66.0",
    "rich>=13.0.0",
    "tenacity>=8.0.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.0.0",
    "black>=23.0.0",
    "mypy>=1.0.0",
]
```

インストール:
```bash
uv pip install -e .
# または開発依存も含めて
uv pip install -e ".[dev]"
```

---

## ライブラリ選定理由

### なぜ yt-dlp か？
- **youtube-dl の後継**: より活発に開発されている
- **対応サイトが豊富**: 1000以上のサイトに対応
- **高機能**: プレイリスト、字幕、品質選択など
- **メンテナンス**: 頻繁に更新されている

**代替案**:
- `youtube-dl`: 更新が遅い（非推奨）
- `pytube`: YouTube専用（機能が限定的）

### なぜ gdown か？
- **Google Drive特化**: Google Driveに最適化
- **シンプルAPI**: 使いやすい
- **フォルダダウンロード対応**: 複数ファイルの一括取得可能

**代替案**:
- `google-api-python-client`: 複雑、認証が必要
- `PyDrive2`: 設定が複雑

### なぜ rich か（オプション）？
- **美しい出力**: テーブル、カラー、プログレスバー
- **デバッグ支援**: `console.log()` で変数を見やすく表示
- **パフォーマンス**: 高速

**代替案**:
- `colorama`: カラー出力のみ（機能が限定的）
- `termcolor`: カラー出力のみ

### なぜ tenacity か？
- **柔軟なリトライ戦略**: 指数バックオフ、条件付きリトライ
- **デコレータベース**: コードが読みやすい
- **エラーハンドリング**: 詳細なログ

**代替案**:
- `backoff`: 似た機能だが tenacity の方が柔軟
- 手動実装: 複雑になりがち

---

## パフォーマンス最適化のためのライブラリ（高度）

### aiohttp + asyncio（非同期処理）
```bash
uv pip install aiohttp aiofiles
```
**目的**: 非同期ダウンロードでさらに高速化
**使用タイミング**: 並列処理でも遅い場合

**メリット**:
- I/O待機時間を削減
- 数百ファイルの並列ダウンロードが可能

**デメリット**:
- コードが複雑になる
- デバッグが難しい

**推奨**: まずは `concurrent.futures` で実装し、必要に応じて移行

---

## 最小構成（必須のみ）
```bash
uv pip install pandas openpyxl yt-dlp gdown
```

## 推奨構成（快適な開発）
```bash
uv pip install pandas openpyxl yt-dlp gdown tqdm rich tenacity
```

## フル構成（すべての機能）
```bash
uv pip install pandas openpyxl yt-dlp gdown tqdm rich tenacity pydantic
```

---

## requirements.txt
```txt
# 必須
pandas>=2.0.0
openpyxl>=3.1.0
yt-dlp>=2024.0.0
gdown>=5.0.0

# UI/UX
tqdm>=4.66.0
rich>=13.0.0

# エラーハンドリング
tenacity>=8.0.0

# オプション
# pydantic>=2.0.0
```

インストール:
```bash
uv pip install -r requirements.txt
```
