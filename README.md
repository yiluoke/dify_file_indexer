# dify_file_indexer (PoC)

社内ファイルサーバー上の設計書をスキャンして、**Difyナレッジに登録するための「検索用インデックスMarkdown」**を生成する PoC 用ツールです。

## できること
- 指定フォルダを再帰スキャン
- 対象拡張子（docx/xlsx/pptx/pdf/txt/md/sql など）から **見出し・シート名・スライドタイトル等を抽出**
- **要約（Pythonのみの簡易抽出）**と **キーワード** を生成
- **ファイルのフルパス**・更新日時・推定版数を含む **1ファイル=1 Markdown** を出力
- (任意) **最新版マップ**（system, screen_id, doc_type ごとの latest_path）を出力
- 差分スキャン（state.json により変更がないファイルはスキップ）
- Windows実行時に **ショートカット（.lnk）を解決してリンク先もスキャン**（configでON/OFF可能）

## セキュリティの考え方（PoC）
- Dify側へは **原本を送らず、インデックスのみ**送る前提
- 抽出テキストは「見出し＋冒頭少量」に限定
- 既定でメール/電話/IP/パスワードっぽい文字列をマスク（configで調整）

## 使い方

### 1) セットアップ
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

> Windowsで `.lnk` 追跡を使う場合は `pywin32` が必要です。
> `requirements.txt` に含まれているので通常は自動で入ります（Windowsのみ）。

### 2) 設定ファイルを作成
`config.example.yml` をコピーして `config.yml` を作ってください。

#### フォルダ除外（キーワード）
ルート配下に `old` / `backup` / `削除` などが含まれるフォルダがある場合、
`exclude_dir_keywords` に指定すると **部分一致**で除外できます。

例:
```yml
exclude_dir_keywords:
  - "old"
  - "backup"
  - "削除"
```

#### Windowsショートカット（.lnk）追跡
`config.yml` の `shortcuts` を設定してください。

- `enabled: true` で `.lnk` を解決し、リンク先が対象拡張子ならインデックス化します。
- リンク先がディレクトリの場合、`follow_dir_targets: true` でそのディレクトリも再帰スキャンします。
- 安全のためデフォルトでは `allow_outside_roots: false`（roots外は追跡しない）です。

ショートカット経由で見つかった場合、生成Markdownに `ALIASES` セクションとして
「どのショートカットから辿れたか」も併記します。

### 3) 実行
```bash
python -m src.scan_kb --config config.yml --out out
```

- 出力:
  - `out/docs/*.md` : Difyに登録するMarkdown
  - `out/latest_map.md` : 最新版マップ（有効化した場合）
  - `out/state.json` : 差分用状態ファイル

## Dify側（手動アップロード PoC）
1. Dataset（ナレッジ）を新規作成（Chunk modeは後から変更できないので注意）
2. `out/docs/*.md` をアップロード
3. アプリ側で Knowledge Retrieval をONにして回答プロンプトを設定（README末尾参照）

## 推奨プロンプト（例）
- 「PATH:」で始まる行を必ず含めて回答
- 最新版は `version_key` と `updated_at` を根拠に選ぶ
- 候補は最大3件、関連資料は最大5件まで

---
"# dify_file_indexer" 
