# WOC講義資料ダウンロードシステム 仕様書

## 1. システム概要

### 1.1 目的
Excelファイルに記載された講義録画URLと資料URLから、動画と資料を自動的にダウンロードし、適切なファイル名で保存するシステムを構築する。

### 1.2 対象ファイル
- **入力**: `draft_【WOC】AIチャットボット開発データベース.xlsx`
- **シート**: 講義録画・資料、グルコン、合宿、コンテンツ、loom

## 2. データ構造

### 2.1 シート別データ構造

#### 講義録画・資料シート
| 列名 | データ型 | 説明 |
|------|---------|------|
| 実施年 | 文字列 | 例: "2025年" |
| 実施日 | 文字列 | 例: "5月15日" |
| 開催種別 | 文字列 | "オンライン" または "リアル" |
| 講義タイトル | 文字列 | 講義の名称 |
| 録画（動画視聴リンク） | URL | Vimeo/YouTube等の動画URL |
| 録画（動画DLリンク） | URL | 動画ダウンロード用URL |
| 資料1 | URL | Google Docs/Drive URL (オプション) |
| 資料2 | URL | Google Docs/Drive URL (オプション) |
| 資料3 | URL | Google Docs/Drive URL (オプション) |
| 資料4 | URL | Google Docs/Drive URL (オプション) |

#### グルコンシート
| 列名 | データ型 | 説明 |
|------|---------|------|
| 実施年 | 文字列 | 例: "2025年" |
| 実施日 | 文字列 | 例: "2025-05-02" (日付形式) |
| 開催種別 | 文字列 | "オンライン" |
| 講義タイトル | 文字列 | "グルコン" (固定) |
| 録画（動画視聴リンク） | URL | Vimeo動画URL |
| 録画（動画DLリンク） | URL | 動画ダウンロード用URL |

#### 合宿シート
構造は「講義録画・資料」と同様

#### コンテンツシート
| 列名 | データ型 | 説明 |
|------|---------|------|
| コンテンツタイトル | 文字列 | 章番号付きタイトル (例: "1-1.高収益ビジネスモデル構築の全体像") |
| Unnamed: 1 | 文字列 | 詳細タイトル |
| 動画リンク | URL | Utage/YouTube/Vimeo/Loom URL |
| 動画DLリンク | URL | ダウンロード用URL (.m3u8含む) |
| 資料1 | URL | Google Docs/Drive URL (オプション) |
| 資料2 | URL | Google Docs/Drive URL (オプション) |

#### loomシート
構造は調査中

## 3. ファイル命名規則

### 3.1 基本原則
- **同じ講義の動画と資料は共通の接頭辞を持つ**
- **ファイルの種類を接尾辞で区別する**
- **特殊文字は除去または置換する**
- **ファイル名の最大長は200文字とする**

### 3.2 シート別命名規則

#### 講義録画・資料 / グルコン / 合宿
```
{実施年4桁}{実施月2桁}{実施日2桁}_{開催種別}_{講義タイトル}_{種別}.{拡張子}
```

**例**:
```
20250515_オンライン_【2025年最新版】スタッフ雇用を守り続ける人事評価制度の要諦_video.mp4
20250515_オンライン_【2025年最新版】スタッフ雇用を守り続ける人事評価制度の要諦_資料1.pdf
20250515_オンライン_【2025年最新版】スタッフ雇用を守り続ける人事評価制度の要諦_資料2.pdf
```

**種別の命名**:
- 動画（視聴リンク列）: `_video_view`
- 動画（DLリンク列）: `_video`
- 資料1: `_資料1`
- 資料2: `_資料2`
- 資料3: `_資料3`
- 資料4: `_資料4`

**注意事項**:
- 「前半」「後半」がタイトルに含まれる場合はそのまま使用
- 実施年・実施日がNaNの場合は、直前の行の値を継承する

#### コンテンツシート
```
{章番号}_{コンテンツタイトル}_{種別}.{拡張子}
```

**例**:
```
1-1_高収益ビジネスモデル構築の全体像_video.mp4
2-2_PDCAサイクル構築ステップ_資料1.pdf
```

**章番号の抽出**:
- コンテンツタイトルから正規表現で抽出: `(\d+-\d+|\d+)`
- 抽出できない場合は行番号を使用

### 3.3 ファイル名サニタイズルール
以下の文字を置換または除去:
```python
置換ルール:
  / → ／ (全角スラッシュ)
  \ → ＼ (全角バックスラッシュ)
  : → ： (全角コロン)
  * → ＊ (全角アスタリスク)
  ? → ？ (全角疑問符)
  " → " (全角ダブルクォート)
  < → ＜ (全角小なり)
  > → ＞ (全角大なり)
  | → ｜ (全角パイプ)

除去文字:
  改行、タブ、連続する空白
```

## 4. ディレクトリ構造

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
│   ├── 20241214_リアル_採用＆教育の仕組み構築合宿（12_14）_video.mp4
│   ├── 20241214_リアル_採用＆教育の仕組み構築合宿（12_14）_資料1.pdf
│   └── ...
├── コンテンツ/
│   ├── 1-1_高収益ビジネスモデル構築の全体像_video.mp4
│   ├── 2-2_PDCAサイクル構築ステップ_資料1.xlsx
│   └── ...
├── loom/
│   └── ...
└── .download_state/
    ├── download_db.json      # ダウンロード状態管理DB
    └── url_dedup.json        # URL重複排除DB
```

## 5. URL種別とダウンローダー

### 5.1 動画URL
| URLパターン | ツール | 備考 |
|------------|--------|------|
| `player.vimeo.com` | yt-dlp | Vimeo直接再生URL |
| `vimeo.com` | yt-dlp | Vimeo共有URL |
| `youtu.be`, `youtube.com` | yt-dlp | YouTube |
| `loom.com` | yt-dlp | Loom録画 |
| `utage-system.com/video/` | yt-dlp | Utage動画 |
| `*.m3u8` | yt-dlp | HLS動画ストリーム |

### 5.2 資料URL
| URLパターン | ツール | 備考 |
|------------|--------|------|
| `docs.google.com/presentation` | gdown | Google Slides (PDF出力) |
| `docs.google.com/spreadsheets` | gdown | Google Sheets (XLSX出力) |
| `docs.google.com/document` | gdown | Google Docs (PDF出力) |
| `drive.google.com/file` | gdown | Google Drive ファイル |
| `drive.google.com/drive/folders` | gdown | Google Drive フォルダ (全ファイル) |

### 5.3 yt-dlpオプション
```python
yt_dlp_options = {
    'format': 'bestvideo+bestaudio/best',
    'outtmpl': '{保存先パス}',
    'quiet': False,
    'no_warnings': False,
    'extract_flat': False,
    'cookiefile': None,  # 必要に応じて設定
}
```

### 5.4 gdownオプション
```python
# Google Docsの場合
gdown.download(url, output='{保存先パス}', quiet=False, fuzzy=True, format='pdf')

# Google Sheetsの場合
gdown.download(url, output='{保存先パス}', quiet=False, fuzzy=True, format='xlsx')

# Google Driveファイルの場合
gdown.download(url, output='{保存先パス}', quiet=False, fuzzy=True)

# フォルダの場合
gdown.download_folder(url, output='{フォルダパス}', quiet=False)
```

## 6. 処理フロー

### 6.1 メイン処理
```
1. Excelファイルを読み込む
2. 各シートを順次処理
   a. シート名を取得
   b. 出力ディレクトリを作成
   c. 各行を処理
      i. メタデータを抽出（実施年、実施日、タイトル等）
      ii. 共通接頭辞を生成
      iii. 動画URLを処理
      iv. 資料URLを処理
   d. 処理結果をログに記録
3. 完了レポートを出力
```

### 6.2 URL処理フロー
```
1. URLの有効性チェック (NaN, 空文字, "-" をスキップ)
2. URLパターンマッチング
3. ファイル名生成
4. 既存ファイルチェック
   - 存在する場合: スキップ
   - 存在しない場合: ダウンロード実行
5. エラーハンドリング
   - 成功: ログに記録
   - 失敗: エラーログに記録、処理継続
```

### 6.3 メタデータ継承ロジック
```python
# 実施年・実施日がNaNの場合、直前の有効な値を使用
previous_year = None
previous_date = None

for row in rows:
    if pd.notna(row['実施年']):
        previous_year = row['実施年']
    else:
        row['実施年'] = previous_year

    if pd.notna(row['実施日']):
        previous_date = row['実施日']
    else:
        row['実施日'] = previous_date
```

### 6.4 Resume機能（中断・再開機能）

#### 目的
ダウンロード中断時に、次回実行時に途中から再開できるようにする。

#### データベース構造（download_db.json）
```json
{
  "downloads": [
    {
      "id": "uuid-xxxxx",
      "url": "https://player.vimeo.com/...",
      "file_path": "downloads/講義録画・資料/20250515_オンライン_..._video.mp4",
      "status": "completed|failed|in_progress",
      "file_size": 123456789,
      "downloaded_at": "2025-10-25T10:30:00",
      "error_message": null,
      "sheet_name": "講義録画・資料",
      "row_index": 0,
      "column_name": "録画（動画DLリンク）"
    }
  ],
  "metadata": {
    "last_updated": "2025-10-25T10:30:00",
    "total_downloads": 150,
    "completed": 100,
    "failed": 5,
    "skipped": 45
  }
}
```

#### Resume処理フロー
```
1. プログラム起動時にdownload_db.jsonを読み込む
2. URLごとに以下をチェック:
   a. DBにレコードが存在するか？
   b. status が "completed" か？
   c. ファイルが実際に存在するか？
   d. ファイルサイズが0より大きいか？
3. 上記すべて満たす場合 → スキップ
4. それ以外 → ダウンロード実行
5. ダウンロード成功時:
   - status を "completed" に更新
   - file_size, downloaded_at を記録
6. ダウンロード失敗時:
   - status を "failed" に更新
   - error_message を記録
```

#### コマンドラインオプション
```bash
# ステータスをリセットして全て再ダウンロード
python download_woc_materials.py --reset

# 失敗したものだけ再試行
python download_woc_materials.py --retry-failed

# ダウンロード状態を表示
python download_woc_materials.py --status
```

### 6.5 Dedup機能（重複排除機能）

#### 目的
同じURLが複数の講義・資料で使用されている場合、実際のダウンロードは1回だけ行い、2回目以降はシンボリックリンクまたは既存ファイルの情報を参照する。

#### データベース構造（url_dedup.json）
```json
{
  "url_hash_to_file": {
    "sha256_hash_of_url_1": {
      "url": "https://docs.google.com/presentation/d/xxxxx/edit",
      "original_file_path": "downloads/講義録画・資料/20250515_..._資料1.pdf",
      "file_size": 2048576,
      "downloaded_at": "2025-10-25T10:30:00",
      "references": [
        {
          "file_path": "downloads/グルコン/20250502_..._資料1.pdf",
          "link_type": "symlink|copy|reference",
          "created_at": "2025-10-25T10:31:00"
        }
      ]
    }
  },
  "metadata": {
    "total_unique_urls": 200,
    "total_references": 250,
    "space_saved_bytes": 52428800
  }
}
```

#### URLハッシュ生成
```python
import hashlib

def get_url_hash(url: str) -> str:
    """URLを正規化してSHA256ハッシュを生成"""
    # URLを正規化（クエリパラメータの順序を統一、末尾スラッシュ除去等）
    normalized_url = normalize_url(url)
    return hashlib.sha256(normalized_url.encode('utf-8')).hexdigest()

def normalize_url(url: str) -> str:
    """URLを正規化"""
    # Google DocsのURL正規化例:
    # /edit?usp=sharing と /edit を同一視
    # クエリパラメータを除去または並び替え
    from urllib.parse import urlparse, parse_qs, urlencode

    parsed = urlparse(url)
    # クエリパラメータを無視する場合
    normalized = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
    return normalized.rstrip('/')
```

#### Dedup処理フロー
```
1. URLからハッシュを生成
2. url_dedup.jsonでハッシュを検索
3. ハッシュが存在する場合:
   a. 元ファイルが存在するか確認
   b. 存在する場合:
      - 新しいファイルパスにシンボリックリンクを作成（Linux/Mac）
      - Windowsの場合はコピーまたはハードリンク
      - referencesに追加
      - ダウンロードスキップ
   c. 存在しない場合:
      - ハッシュエントリを削除
      - 通常のダウンロードを実行
4. ハッシュが存在しない場合:
   a. 通常のダウンロードを実行
   b. 成功時にurl_dedup.jsonに登録
```

#### シンボリックリンク vs コピー
**デフォルト: シンボリックリンク**
- 利点: ディスク容量節約
- 欠点: 元ファイル削除時にリンク切れ

**コピーモード (--dedup-mode copy)**
- 利点: ファイル間の独立性
- 欠点: ディスク容量を消費

**参照のみモード (--dedup-mode reference)**
- ファイルを作成せず、DBに記録のみ
- ユーザーが手動で参照先を確認

#### コマンドラインオプション
```bash
# Dedup機能を無効化（全て個別ダウンロード）
python download_woc_materials.py --no-dedup

# Dedupモードを指定（symlink|copy|reference）
python download_woc_materials.py --dedup-mode symlink

# 重複URL統計を表示
python download_woc_materials.py --dedup-stats
```

#### Dedup統計レポート例
```
=== URL重複排除統計 ===
総URL数: 500
ユニークURL数: 200
重複URL数: 300
節約ディスク容量: 1.5 GB

重複が多いURL Top 5:
1. https://docs.google.com/presentation/d/xxxxx (15回使用)
2. https://docs.google.com/spreadsheets/d/yyyyy (12回使用)
...
```

## 7. エラーハンドリング

### 7.1 エラーレベル
- **INFO**: 正常終了、スキップ（既存ファイル）
- **WARNING**: URLが無効、フォーマット不明
- **ERROR**: ダウンロード失敗、ネットワークエラー

### 7.2 ログ出力
```
logs/
├── download_success.log  # 成功ログ
├── download_error.log    # エラーログ
├── download_skip.log     # スキップログ
├── download_resume.log   # Resume関連ログ
└── download_dedup.log    # Dedup関連ログ
```

**ログフォーマット例**:
```
[2025-10-25 10:30:00] [INFO] [SUCCESS] Downloaded: 20250515_オンライン_..._video.mp4 (125.5 MB)
[2025-10-25 10:30:05] [INFO] [SKIP] Already exists: 20250502_オンライン_グルコン_video.mp4
[2025-10-25 10:30:10] [ERROR] [FAILED] URL: https://... Error: Connection timeout
[2025-10-25 10:30:15] [INFO] [DEDUP] Symlink created: 20250515_..._資料1.pdf -> 20250502_..._資料1.pdf
[2025-10-25 10:30:20] [INFO] [RESUME] Skipped (already completed): 20250515_..._video.mp4
```

### 7.3 エラー時の動作
- ダウンロード失敗時は処理をスキップし、次のURLへ進む
- 致命的エラー（ファイル読み込み失敗等）の場合は処理を中断

## 8. 実行環境

### 8.1 必要なパッケージ
```
pandas
openpyxl
yt-dlp
gdown
```

### 8.2 インストール方法（uv使用）
```bash
uv pip install pandas openpyxl yt-dlp gdown
```

または

```bash
uv add pandas openpyxl yt-dlp gdown
```

### 8.3 実行コマンド
```bash
python download_woc_materials.py
```

### 8.4 オプション引数
```bash
# 特定のシートのみ処理
python download_woc_materials.py --sheet 講義録画・資料

# ドライランモード（ダウンロードせず、処理内容のみ表示）
python download_woc_materials.py --dry-run

# 既存ファイルを上書き
python download_woc_materials.py --overwrite

# 並列ダウンロード数を指定
python download_woc_materials.py --parallel 3

# === Resume機能関連 ===
# ステータスをリセットして全て再ダウンロード
python download_woc_materials.py --reset

# 失敗したものだけ再試行
python download_woc_materials.py --retry-failed

# ダウンロード状態を表示
python download_woc_materials.py --status

# === Dedup機能関連 ===
# Dedup機能を無効化（全て個別ダウンロード）
python download_woc_materials.py --no-dedup

# Dedupモードを指定（symlink|copy|reference）
python download_woc_materials.py --dedup-mode symlink

# 重複URL統計を表示
python download_woc_materials.py --dedup-stats
```

## 9. パフォーマンス要件

### 9.1 処理速度
- 1動画あたり: ネットワーク速度に依存（通常1-5分）
- 1資料あたり: 5-30秒

### 9.2 並列処理
- デフォルト: 1並列（順次処理）
- オプション: 最大5並列まで対応

### 9.3 レート制限対策
- 連続ダウンロード間に1秒の待機時間を設定
- エラー発生時は指数バックオフ（1秒 → 2秒 → 4秒 → 8秒）

## 10. 制約事項

### 10.1 Google Drive制限
- 共有設定が「リンクを知っている全員が閲覧可能」になっている必要がある
- 大容量ファイルはGoogle Driveのウイルススキャン警告が出る場合がある

### 10.2 ファイルサイズ
- 個別ファイルのサイズ制限なし（ディスク容量に依存）

### 10.3 その他
- yt-dlpで対応していない動画形式は手動ダウンロードが必要
- 認証が必要なURLには対応しない

## 11. テスト項目

### 11.1 機能テスト
- [ ] Vimeo動画のダウンロード
- [ ] YouTube動画のダウンロード
- [ ] Utage動画のダウンロード
- [ ] .m3u8形式のダウンロード
- [ ] Google Slidesのダウンロード（PDF変換）
- [ ] Google Sheetsのダウンロード（XLSX変換）
- [ ] Google Driveファイルのダウンロード
- [ ] Google Driveフォルダのダウンロード
- [ ] ファイル名サニタイズ
- [ ] 既存ファイルのスキップ
- [ ] メタデータ継承（NaN処理）

#### Resume機能テスト
- [ ] 中断後の再開（DBから状態を復元）
- [ ] 完了済みファイルのスキップ
- [ ] 失敗ファイルの再試行（--retry-failed）
- [ ] ステータスリセット（--reset）
- [ ] ステータス表示（--status）
- [ ] ファイルが削除された場合の再ダウンロード

#### Dedup機能テスト
- [ ] 同一URLの検出
- [ ] シンボリックリンクの作成
- [ ] コピーモードでの動作
- [ ] 参照のみモードでの動作
- [ ] 元ファイル削除時の再ダウンロード
- [ ] URL正規化（クエリパラメータ違いの同一URL検出）
- [ ] 重複統計の表示（--dedup-stats）

### 11.2 エラーハンドリングテスト
- [ ] 無効なURLのスキップ
- [ ] ダウンロード失敗時の継続処理
- [ ] ネットワークエラー時のリトライ

### 11.3 パフォーマンステスト
- [ ] 大量ファイル処理時のメモリ使用量
- [ ] 並列ダウンロード時の安定性

## 12. 今後の拡張案

- [ ] 進捗バーの表示（tqdm）
- [x] ダウンロード履歴のデータベース化（Resume機能として実装済み）
- [x] 差分ダウンロード（Resume機能として実装済み）
- [x] 重複排除機能（Dedup機能として実装済み）
- [ ] Web UIの追加
- [ ] クラウドストレージへの直接アップロード（Google Drive、S3等）
- [ ] ダウンロード完了時の通知機能（Slack、Email等）
- [ ] スケジューラー機能（定期的な自動ダウンロード）
- [ ] 動画のトランスコード機能（形式変換、圧縮）
- [ ] メタデータの自動抽出（動画の長さ、解像度等）
- [ ] 全文検索機能（資料の内容を検索）
