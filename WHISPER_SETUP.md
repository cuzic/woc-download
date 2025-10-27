# OpenAI Whisper API を使った文字起こし設定

## 概要

utage-system の動画には字幕が含まれていないため、OpenAI の Whisper API を使って自動的に文字起こしを行います。

## 必要な準備

### 1. OpenAI API キーの取得

1. https://platform.openai.com/ にアクセス
2. アカウントを作成またはログイン
3. API Keys ページで新しい API キーを作成
4. API キーをコピー（一度しか表示されません）

### 2. 環境変数の設定

#### Linux/Mac の場合

```bash
export OPENAI_API_KEY="sk-..."
```

または、`.bashrc` / `.zshrc` に追加：

```bash
echo 'export OPENAI_API_KEY="sk-..."' >> ~/.bashrc
source ~/.bashrc
```

#### Windows の場合

PowerShell:
```powershell
$env:OPENAI_API_KEY="sk-..."
```

または、システム環境変数として設定

### 3. ffmpeg のインストール（音声抽出に必要）

#### Linux (Debian/Ubuntu)
```bash
sudo apt-get install ffmpeg
```

#### Mac
```bash
brew install ffmpeg
```

#### Windows
https://ffmpeg.org/download.html からダウンロードしてインストール

## 使用方法

### 基本的な使用

```bash
# 環境変数を設定
export OPENAI_API_KEY="sk-..."

# ダウンロードを実行
.venv/bin/python -m woc_download --sheet コンテンツ
```

### utage-system 動画の処理フロー

1. **m3u8 URL の抽出**: playwright で HTML を解析して動画 URL を取得
2. **音声のダウンロード**: yt-dlp で音声を MP3 形式でダウンロード
3. **文字起こし**: OpenAI Whisper API で日本語の文字起こしを実行
4. **字幕ファイル生成**: SRT 形式で字幕ファイルを保存

## 料金について

OpenAI Whisper API の料金（2025年1月時点）：
- **$0.006 / 分**

例：
- 10分の動画 → $0.06
- 60分の動画 → $0.36

詳細: https://openai.com/pricing

## 制限事項

1. **ファイルサイズ制限**: 25MB まで
   - 長い動画の場合、自動的に音声品質を調整します
   - それでも超える場合はエラーになります

2. **処理時間**: 動画の長さに応じて数分かかる場合があります

3. **言語**: 現在は日本語固定
   - 必要に応じて `language="ja"` を変更可能

## トラブルシューティング

### エラー: "OPENAI_API_KEY environment variable is not set"

環境変数が設定されていません。上記の手順で設定してください。

### エラー: "File size exceeds 25MB limit"

動画が長すぎます。以下の対処法：
1. より短い区間に分割
2. 音声品質を下げる（実装済み: 192kbps）

### エラー: "FFmpeg not found"

ffmpeg がインストールされていません。上記の手順でインストールしてください。

## テスト

単一の utage-system URL でテスト：

```python
from src.woc_download.executor import DownloadExecutor
import os

# API キーを設定
os.environ["OPENAI_API_KEY"] = "sk-..."

executor = DownloadExecutor()

# utage-system URL から m3u8 を抽出
url = "https://utage-system.com/video/zwNkkYvDGkMN"
m3u8_url = executor.extract_utage_m3u8_url(url)
print(f"M3U8 URL: {m3u8_url}")

# 文字起こしを実行
result = executor.transcribe_video_with_whisper(m3u8_url, "./test_output")
print(f"Success: {result.success}")
print(f"File: {result.file_path}")
```

## 参考リンク

- [OpenAI Whisper API Documentation](https://platform.openai.com/docs/guides/speech-to-text)
- [OpenAI Pricing](https://openai.com/pricing)
- [yt-dlp Documentation](https://github.com/yt-dlp/yt-dlp)
