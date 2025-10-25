# WOC講義資料ダウンロードシステム 使い方ガイド

## クイックスタート

### 1. 基本的な使用方法

```bash
# すべてのシートをダウンロード
.venv/bin/python download_woc_materials.py

# 特定のシートのみダウンロード
.venv/bin/python download_woc_materials.py --sheet 講義録画・資料

# 複数のシートを指定
.venv/bin/python download_woc_materials.py --sheet 講義録画・資料 グルコン
```

### 2. ドライランモード（実際にダウンロードせずテスト）

```bash
# 何がダウンロードされるか確認
.venv/bin/python download_woc_materials.py --dry-run

# 特定シートをドライラン
.venv/bin/python download_woc_materials.py --dry-run --sheet 講義録画・資料
```

## 機能別使用方法

### Resume機能（中断・再開）

ダウンロード中に中断しても、次回実行時は完了済みのファイルをスキップして続きから再開します。

```bash
# 普通に実行（自動的にResume機能が働く）
.venv/bin/python download_woc_materials.py

# ダウンロード状態を確認
.venv/bin/python download_woc_materials.py --status

# 失敗したダウンロードのみ再試行
.venv/bin/python download_woc_materials.py --retry-failed

# ダウンロード状態をリセット（最初からやり直す）
.venv/bin/python download_woc_materials.py --reset
```

**ステータス表示例:**
```
     ダウンロード状態
┏━━━━━━━━━━━━━━━━━━┳━━━━━━┓
┃ 項目             ┃ 件数 ┃
┡━━━━━━━━━━━━━━━━━━╇━━━━━━┩
│ 総ダウンロード数 │   72 │
│ ✓ 完了           │   68 │
│ ⊗ 失敗           │    4 │
│ ⏳ 進行中        │    0 │
└──────────────────┴──────┘
```

### Dedup機能（重複排除）

同じURLが複数の講義で使用されている場合、1回だけダウンロードして残りはシンボリックリンクを作成します。

```bash
# デフォルト（シンボリックリンクモード）
.venv/bin/python download_woc_materials.py

# コピーモード（実ファイルをコピー）
.venv/bin/python download_woc_materials.py --dedup-mode copy

# 参照のみモード（DBに記録だけ、リンク作成なし）
.venv/bin/python download_woc_materials.py --dedup-mode reference

# Dedup機能を無効化（すべて個別ダウンロード）
.venv/bin/python download_woc_materials.py --no-dedup

# 重複排除の統計を表示
.venv/bin/python download_woc_materials.py --dedup-stats
```

**重複排除統計の例:**
```
重複排除統計
┏━━━━━━━━━━━━━━━┳━━━━━━━━━┓
┃ 項目          ┃      値 ┃
┡━━━━━━━━━━━━━━━╇━━━━━━━━━┩
│ ユニークURL数 │     200 │
│ 参照数        │     300 │
│ 節約容量      │ 1.5 GB  │
└───────────────┴─────────┘

重複が多いURL Top 5:
  1. https://docs.google.com/presentation/d/xxxxx (15回使用)
  2. https://docs.google.com/spreadsheets/d/yyyyy (12回使用)
  ...
```

### その他のオプション

```bash
# 既存ファイルを上書き
.venv/bin/python download_woc_materials.py --overwrite

# ダウンロード先ディレクトリを変更
.venv/bin/python download_woc_materials.py --download-dir /path/to/downloads
```

## 実行例

### 例1: 初回実行（全シートダウンロード）

```bash
.venv/bin/python download_woc_materials.py
```

**出力:**
```
WOC講義資料ダウンロードシステム
Excel: draft_【WOC】AIチャットボット開発データベース.xlsx

Processing 5 sheets: 講義録画・資料, グルコン, 合宿, コンテンツ, loom

Processing sheet: 講義録画・資料
  講義録画・資料 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 100%
  Completed: 73, Failed: 0, Skipped: 0, Deduped: 0

Processing sheet: グルコン
  グルコン ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 100%
  Completed: 18, Failed: 0, Skipped: 0, Deduped: 5

...

===== ダウンロード完了レポート =====

        全体統計
┏━━━━━━━━━━━━━┳━━━━━━━━┓
┃ 項目        ┃   件数 ┃
┡━━━━━━━━━━━━━╇━━━━━━━━┩
│ 総タスク数  │    250 │
│ ✓ 完了      │    240 │
│ ⊗ 失敗      │      5 │
│ ⊘ スキップ  │      0 │
│ 🔗 重複排除 │     15 │
│ 実行時間    │ 1250秒 │
└─────────────┴────────┘
```

### 例2: 中断後の再開

ダウンロード中に Ctrl+C で中断した後：

```bash
.venv/bin/python download_woc_materials.py
```

完了済みのファイルは自動的にスキップされ、未完了分のみダウンロードされます。

### 例3: 失敗分のみ再試行

```bash
# 失敗したダウンロードを確認
.venv/bin/python download_woc_materials.py --status

# 失敗分のみ再試行
.venv/bin/python download_woc_materials.py --retry-failed
```

## ファイル構成

ダウンロード後のディレクトリ構造：

```
downloads/
├── 講義録画・資料/
│   ├── 20250515_オンライン_【2025年最新版】スタッフ雇用を守り続ける人事評価制度の要諦_video.mp4
│   ├── 20250515_オンライン_【2025年最新版】スタッフ雇用を守り続ける人事評価制度の要諦_資料1.pdf
│   └── ...
├── グルコン/
│   ├── 20250502_オンライン_グルコン_video.mp4
│   └── ...
├── 合宿/
│   └── ...
├── コンテンツ/
│   ├── 1-1_高収益ビジネスモデル構築の全体像_video.mp4
│   └── ...
├── loom/
│   └── ...
└── .download_state/
    ├── download_db.json      # Resume用DB
    └── url_dedup.json        # Dedup用DB
```

## ログファイル

実行ログは `logs/` ディレクトリに保存されます：

```
logs/
├── download_success.log  # 成功ログ
├── download_error.log    # エラーログ
├── download_skip.log     # スキップログ
├── download_resume.log   # Resume関連ログ
└── download_dedup.log    # Dedup関連ログ
```

**ログ例:**
```
[2025-10-25 10:30:00] [INFO] Downloaded video: downloads/講義録画・資料/20250515_..._video.mp4 (125.5 MB)
[2025-10-25 10:30:05] [INFO] [RESUME] Skipped (already completed): downloads/...
[2025-10-25 10:30:10] [ERROR] Failed to download video https://...: Connection timeout
[2025-10-25 10:30:15] [INFO] [DEDUP] Symlink created: downloads/... -> downloads/...
```

## トラブルシューティング

### 問題1: Google Driveからダウンロードできない

**原因**: 共有設定が適切でない

**解決策**:
1. Google Driveのファイルを開く
2. 「共有」ボタンをクリック
3. 「リンクを知っている全員が閲覧可能」に変更

### 問題2: yt-dlpでダウンロードできない

**原因**: yt-dlpのバージョンが古い

**解決策**:
```bash
uv pip install --upgrade yt-dlp
```

### 問題3: ファイル名が長すぎる

**原因**: ファイル名が200文字を超えている

**解決策**: 自動的に200文字に切り詰められますが、さらに短くする場合は `FileNameGenerator.sanitize_filename()` の `max_length` を調整してください。

### 問題4: シンボリックリンクが機能しない（Windows）

**原因**: Windowsではシンボリックリンクに管理者権限が必要な場合がある

**解決策**:
```bash
# コピーモードを使用
.venv/bin/python download_woc_materials.py --dedup-mode copy
```

## パフォーマンスチューニング

### 並列ダウンロード（将来実装予定）

```bash
# 3並列でダウンロード
.venv/bin/python download_woc_materials.py --parallel 3
```

現在のバージョンでは並列処理は未実装ですが、`--parallel` オプションは受け付けます（将来のバージョンで有効化予定）。

## コマンドライン引数一覧

| 引数 | 説明 | デフォルト |
|------|------|-----------|
| `excel_path` | Excelファイルのパス | `draft_【WOC】AIチャットボット開発データベース.xlsx` |
| `--download-dir` | ダウンロード先ディレクトリ | `downloads` |
| `--sheet` | 処理対象シート名（複数指定可） | すべてのシート |
| `--dry-run` | ドライランモード | False |
| `--overwrite` | 既存ファイルを上書き | False |
| `--parallel` | 並列ダウンロード数 | 1（未実装） |
| `--status` | ダウンロード状態を表示 | - |
| `--reset` | ダウンロード状態をリセット | - |
| `--retry-failed` | 失敗したダウンロードを再試行 | - |
| `--no-dedup` | 重複排除機能を無効化 | False |
| `--dedup-mode` | 重複時の動作（symlink/copy/reference） | symlink |
| `--dedup-stats` | 重複排除の統計情報を表示 | - |

## ヘルプ表示

```bash
.venv/bin/python download_woc_materials.py --help
```
