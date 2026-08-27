# ExceltoMarkdown

Excel・Word・Webブラウザからコピーした内容を、GitHub・Markdown文書・LLMで扱いやすい **GitHub Flavored Markdown（GFM）** へ変換する軽量なクリップボードツールです。

このツールは、ExcelをMarkdownで置き換えるものではありません。Excelでの計算・編集・共有はそのまま活かし、結果の一部をGit、文書、AIへ渡す際の摩擦を減らす「変換ブリッジ」です。

> [!IMPORTANT]
> Excelファイルを完全・可逆に変換することは目的としていません。数式、セル結合、色、罫線、列幅、図形など、GFMで表現できない情報は変換時に失われます。

## 主な機能

- ExcelからコピーしたTSVをGFM tableへ変換
- Excelのquoted TSVを解釈し、セル内改行・タブ・二重引用符を保持
- WordやブラウザのClipboard HTML Formatから、段落・太字・斜体・リンク・簡易リストを変換
- Windowsのタスクトレイ、ホットキー、ダブルクリックから変換
- 任意のExcel COM連携による表示文字列・太字・斜体・ハイパーリンクの取得
- Windows以外でも利用できる標準入力（`--stdin`）変換
- 日本語／英語UI

## 変換契約

変換時に保持する情報と、失われる・正規化される情報は次のとおりです。

| 入力 | 保持する情報 | 失われる／正規化される情報 |
| --- | --- | --- |
| Excel clipboard TSV | セル文字列、行列構造、quoted field内の改行・タブ・引用符 | 書式、数式そのもの、結合セル、色、罫線、列幅、図形 |
| Excel COM（任意） | 表示文字列、太字、斜体、ハイパーリンク | 色、罫線、フォントサイズ、数式そのもの、図形 |
| Word / Web HTML | 段落、改行、太字、斜体、リンク、簡易リスト | レイアウト、画像、複雑なCSS、未対応のHTML構造 |
| セル内改行 | 改行位置 | GFM table内で `<br>` へ正規化 |
| GFM table | 行列 | **コピー範囲の先頭行をheader rowとして扱う** |

GFM tableにはheader rowとdelimiter rowが必須です。そのため、元データに見出しがなくても先頭行へheaderの意味が付与されます。これは表示上の変換だけでなく、意味上の変換でもあります。

## 変換例

Excelのコピー結果が次のTSVの場合:

```text
品名	数量	備考
A	10	通常品
B	20	"2行の
備考"
```

セル内改行がExcelのquoted fieldとして渡された場合、次のGFM tableになります。

```markdown
| 品名 | 数量 | 備考 |
| --- | --- | --- |
| A | 10 | 通常品 |
| B | 20 | 2行の<br>備考 |
```

## 説明動画

### Excel表をGFM tableへ変換

<img width="779" height="360" alt="Excel table conversion demo" src="https://github.com/user-attachments/assets/a7dd24ac-0f1e-45a5-bbb6-f704d6e33150" />

### Word・ブラウザの装飾をGFMへ変換

<img width="779" height="247" alt="Rich text conversion demo" src="https://github.com/user-attachments/assets/f801dad8-5648-405b-8845-85faef3e4e9c" />

## インストール

### Windows x64版（推奨）

[GitHub Releases](https://github.com/Akihiko-Fuji/ExceltoMarkdown/releases/latest) から `ExcelToMarkdown-windows-x64.zip` をダウンロードし、ZIPを展開して `ExcelToMarkdown.exe` を起動します。

Nuitkaのstandalone配布にはEXEが利用するDLLなどが含まれます。`ExcelToMarkdown.exe` だけを取り出さず、展開したフォルダー一式を同じ場所に置いてください。設定はEXEと同じ場所の `config.ini` を編集します。

配布ZIPにはSHA-256 checksumの `ExcelToMarkdown-windows-x64.zip.sha256` も添付します。

### Pythonから実行

Python 3.11以降を使用します。

```powershell
git clone https://github.com/Akihiko-Fuji/ExceltoMarkdown.git
cd ExceltoMarkdown
python -m pip install .
```

Excel COM連携も利用する場合は、Windowsで任意dependencyを追加します。

```powershell
python -m pip install ".[excel]"
```

source checkoutから従来どおり `python excel_to_markdown.py` で起動することもできます。

## 使い方

### Windowsタスクトレイアプリ

```powershell
exceltomarkdown
```

または:

```powershell
python -m exceltomarkdown
```

1. Excelで表の範囲をコピーします。
2. 初期値 `Ctrl+Alt+M` のホットキー、タスクトレイメニュー、またはトレイアイコンのダブルクリックで変換します。
3. クリップボードがGFMへ置き換わるので、`.md` ファイルなどへ貼り付けます。

右クリックメニューでは「Markdown表に変換」と「リッチテキストをMarkdown化」を明示的に選べます。

### 1回だけ変換（Windows）

```powershell
exceltomarkdown --once
exceltomarkdown --once --mode rich_text
```

### 標準入力から変換（OS共通）

```powershell
Get-Content sample.tsv -Raw | exceltomarkdown --stdin
```

Linux / Unix / macOSでは、Windows APIを使わない `--stdin` のみ利用できます。

## 設定

source checkoutではリポジトリ直下、EXE配布時はEXEと同じ場所にある `config.ini` を読み込みます。

```ini
[shortcut]
key = Ctrl+Alt+M

[conversion]
# table または rich_text
default_mode = table

# trueの場合、pywin32経由のExcel選択範囲をclipboard TSVより優先
prefer_excel = false

[ui]
# auto / ja / en
language = auto
```

`prefer_excel` は `1 / yes / true / on / enabled` または `0 / no / false / off / disabled` を指定できます。

指定できる修飾キーは `Ctrl / Alt / Shift / Win`、通常キーは英数字1文字、`F1`〜`F24`、`Enter`、`Esc`、`Space`、`Tab`、矢印キーなどです。`Ctrl+Shift+V` はWindowsやOfficeの「テキストのみ貼り付け」と競合しやすいため、既定値にはしていません。

### Excel COM連携の注意

`prefer_excel = true` の場合、コピー済みclipboardではなく、変換時点のExcel選択範囲が対象になります。選択範囲が変わっていると別の表を変換するため、書式が不要なら既定値の `false` を推奨します。

タスクトレイからの変換はworker threadで動作します。COM利用時はそのthread内でCOM apartmentを初期化・解放します。

## HTML→GFM変換の安全策

HTMLの文字列が意図せず見出し、引用、リスト、水平線などのGFM block syntaxへ変わらないよう、行頭記号をエスケープします。リンク先の空白・括弧・制御文字もpercent encodingし、Markdownリンクの境界を壊さないようにします。

これは汎用HTML変換器ではありません。画像、複雑なtable、CSSレイアウトなどの忠実な変換が必要な用途は対象外です。

## コード構成

```text
exceltomarkdown/
├─ core.py        # quoted TSV、Cell、GFM table
├─ rich_text.py   # Clipboard HTML→GFM
├─ config.py      # config.ini、hotkey、i18n
├─ windows.py     # Win32 clipboard、Excel COM、tray
└─ cli.py         # command line entry point
```

`excel_to_markdown.py` は従来の起動方法・importを維持する互換entry pointです。

## 開発

```powershell
python -m unittest discover -v
python -m pip wheel . --no-deps --no-build-isolation
```

### Windows x64 standalone build

64-bit版Python 3.11を使い、dependencyとNuitkaをインストールしてからbuild scriptを実行します。

```batch
python -m pip install ".[excel]" "Nuitka==4.1.3"
build_windows.cmd
```

生成物は `excel_to_markdown.dist` フォルダーです。動作確認・配布ともフォルダー一式を使用します。

CIはUbuntuとWindowsの両方でunit testを実行します。package versionと同じ `v*` tag（例: `v1.1.0`）をpushすると、次のファイルを [ExceltoMarkdownのGitHub Releases](https://github.com/Akihiko-Fuji/ExceltoMarkdown/releases) へ添付します。

- `ExcelToMarkdown-windows-x64.zip`
- `ExcelToMarkdown-windows-x64.zip.sha256`
- Python wheel（`.whl`）
- source distribution（`.tar.gz`）

## ライセンス

[Apache License 2.0](LICENSE)
