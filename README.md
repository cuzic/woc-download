# WOC講義資料ダウンロードシステム

Excelファイルに記載された講義録画URLと資料URLから、動画と資料を自動的にダウンロードするシステムです。

## 機能

- **自動ダウンロード**: YouTube, Vimeo, Loom, Google Drive等からの一括ダウンロード
- **Resume機能**: 中断しても続きから再開可能
- **Dedup機能**: 重複URLを自動検出して容量を節約
- **進捗表示**: ダウンロード状況をリアルタイム表示

## インストール

```bash
uv sync
```

## 使用方法

```bash
# 基本実行
python download_woc_materials.py

# 特定のシートのみ処理
python download_woc_materials.py --sheet 講義録画・資料

# ドライランモード
python download_woc_materials.py --dry-run

# ステータス確認
python download_woc_materials.py --status
```

## ドキュメント

- [仕様書](SPECIFICATION.md)
- [設計ドキュメント](DESIGN.md)
- [ライブラリ一覧](LIBRARIES.md)
