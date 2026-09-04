# Academic Paper Editor v18

Word原稿を段落・画像・表・Word数式（OMML）単位で読み込み、編集・翻訳・出版社書式で再出力するWebアプリである。

## v18 の主な変更

### 1. ローカル翻訳を刷新

v16のHelsinki OPUS-MTを既定から外し、**FuguMT（日英専用）を標準**に変更した。

選択可能なローカルモデル:

- **FuguMT（推奨）**: `staka/fugumt-ja-en` / `staka/fugumt-en-ja`
- **M2M100 418M（大容量）**: `facebook/m2m100_418M`
- **OPUS-MT（軽量・旧方式）**: v16互換用

いずれもAPIキーは不要である。初回だけモデル取得が必要で、キャッシュ後はオフラインで利用できる。

FuguMTだけを事前取得:

```bash
```

M2M100を取得:

```bash
```

すべて取得:

```bash
```

完全オフラインで実行する場合:

```bash
LOCAL_TRANSLATION_OFFLINE=1 python app.py
```

Windows PowerShell:

```powershell
$env:LOCAL_TRANSLATION_OFFLINE="1"
python app.py
```

数式、引用番号、DOI/URL、Figure/Table/Equation番号、大文字略語、ユーザー用語辞書は翻訳時に保護される。日本語出力は常体・である調へ正規化される。

### 2. 「素材」アップロードを初回インポートと統一

素材タブのWordアップロードも、初回アップロードと同じ構造解析器を使用する。

- 段落単位で抽出
- 見出し構造を保持
- 埋込画像を抽出・表示
- Word数式（OMML）を保持
- 表を表ブロックとして保持
- 現在の編集位置へ挿入
- 元ファイルも素材一覧へ保存

素材一覧に残った文書は「構造を再解析して挿入」から再利用できる。

### 3. 既存機能

- 新規プロジェクト作成
- `.paperproj.zip` によるプロジェクト保存・再読込
- DOCXのタイトル・著者・所属・Abstract抽出
- 段落単位編集
- Word数式（OMML）の抽出・保持・再出力
- 文書内画像の表示・挿入・ZIP出力
- 日本語 / English UI
- IEEE / Elsevier / Springer Nature / MDPI / Wiley / Taylor & Francis / Generic 書式
- Word / LaTeX / PDF印刷表示への出力
- 日本語本文の常体・である調への正規化
- HTTP 500を避けるAPIエラーハンドリング

## インストール

```bash
pip install -r requirements.txt
python app.py
```

既定URL:

```text
http://127.0.0.1:5000
```

## ローカル翻訳の状態確認

```text
http://127.0.0.1:5000/api/translation/status
```

各翻訳プロファイルのモデルID、キャッシュ状態、オフラインモードを確認できる。

## 注意

ローカル翻訳は外部APIへ原稿を送信しない一方、翻訳品質はモデル・分野・文章構造に依存する。複雑な学術日本語ではFuguMTとM2M100を比較し、最終投稿前に著者が英文を確認することを推奨する。

出版社プリセットは出版社レベルの一般的な著者原稿書式である。最終投稿時には、対象ジャーナル固有の最新テンプレートおよびInstructions for Authorsを優先する。

## v18 changes

- Mobile-first UI with bottom navigation and a touch-friendly action sheet.
- Online translation only. No local translation model or API key is required by the app; translation uses `deep-translator`'s Google Translate integration. Availability can be affected by service-side rate limits.
- Browser-scoped workspace isolation for public deployment. Manuscripts and uploaded assets from different browsers are stored in separate runtime folders.
- Japanese translations remain normalized to academic plain style (常体・である調).
- Publisher presets, project save/open, structured Word material import, Word OMML equations, images, and exports remain available.
- Vercel deployment configuration is included.
