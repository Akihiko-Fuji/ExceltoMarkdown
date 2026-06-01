
Excelでコピーした表をMarkdownテーブルに変換するシンプルなツールです。Windowsではクリップボード常駐アプリとして使え、Linux/Unix/macOSなどでも標準入力のTSV変換を利用できます。

## 簡単な説明
<img width="779" height="360" alt="demo" src="https://github.com/user-attachments/assets/a7dd24ac-0f1e-45a5-bbb6-f704d6e33150" />

## できること

- Windowsでは、Excelで範囲をコピーした後、タスクトレイ常駐アプリからMarkdownテーブルへ変換します。
- ホットキーは `config.ini` の `key` で指定できます（初期値は `Ctrl+Alt+M`）。タスクトレイの右クリックメニュー、またはトレイアイコンのダブルクリックでも変換できます。
-  `pywin32` ライブラリが導入済みの環境で、`config.ini` の `prefer_excel = true` を指定すると、起動中のExcelの選択範囲から太字、イタリック、ハイパーリンクを読み取り、Markdownへ反映します。
- タスクトレイと実行ファイルのアイコンには、同梱の `e2m_ico.ico` を利用します。
- クリップボードTSV変換では、セル内タブ・セル内改行を含むデータは正しく復元できない場合があります。書式やセル単位の取得を重視する場合は、Windows上でExcel選択範囲変換を有効にしてください。

## 変換例

Excelのコピー結果が次のテキストの場合:

```text
A1	B1	C1
A2	B2	C2
A3	B3	C3
```

出力は次のMarkdownになります。

```markdown
| A1 | B1 | C1 |
| --- | --- | --- |
| A2 | B2 | C2 |
| A3 | B3 | C3 |
```

## 使い方

### タスクトレイ常駐アプリとして起動

```powershell
python excel_to_markdown.py
```

1. Excelで表の範囲をコピーします。
2. `config.ini` に設定したショートカットキー（初期値は `Ctrl+Alt+M`）を押すか、タスクトレイアイコンのメニューから「Markdownに変換」を選びます。
3. クリップボードの内容がMarkdownテーブルに置き換わるので、`.md` ファイルへ貼り付けます。

### ショートカットキーの設定

`excel_to_markdown.py` と同じフォルダー（EXE配布時はEXEと同じフォルダー）にある `config.ini` で、変換に使うショートカットキーを指定します。

```ini
[shortcut]
key = Ctrl+Alt+M

[conversion]
# 通常はコピー済みクリップボードTSVを優先します。
# trueにすると、pywin32利用時だけExcelの現在選択範囲を先に読み取ります。
prefer_excel = false
```

`prefer_excel` の真偽値は `1` / `yes` / `true` / `on` / `enabled` と `0` / `no` / `false` / `off` / `disabled` を指定できます。

指定できる修飾キーは `Ctrl` / `Alt` / `Shift` / `Win` です。通常キーは英数字1文字、`F1`〜`F24`、`Enter`、`Esc`、`Space`、`Tab`、矢印キーなどを指定できます。例: `Ctrl+Shift+M`、`Alt+F12`。 `Ctrl+Shift+V` はWindowsやOfficeで「テキストのみ貼り付け」に使われるため、既定値にはしていません。

### 1回だけ変換して終了（Windowsのみ）

```powershell
python excel_to_markdown.py --once
```

`--once` はWindowsのクリップボード（および設定に応じてExcel選択範囲）を1回だけ変換して終了します。Windows APIを使うため、Linux/Unix/macOSでは利用できません。

### 標準入力で変換を確認（OSを問わず利用可）

```powershell
Get-Content sample.tsv -Raw | python excel_to_markdown.py --stdin
```

`--stdin` はWindows APIを使わず、標準入力のTSVをMarkdownへ変換して標準出力へ書き出します。

## Excel書式の反映（任意）

太字、イタリック、リンクを反映したい場合のみ、軽量な標準機能の範囲を超えるため `pywin32` を追加してください。

```powershell
python -m pip install pywin32
```

`pywin32` がない環境、または `prefer_excel = false` の環境では、Excelがクリップボードに置いたプレーンテキスト（TSV）だけを使って変換します。`prefer_excel = true` は太字・イタリック・リンクを反映したい場合に有効ですが、Excelの現在選択範囲がコピー済みクリップボード内容と異なる場合は、現在選択範囲が変換対象になります。

## Windows Executableとして配布する例

配布用EXEを作る場合は、ビルド環境だけにPyInstallerを入れる構成にすると、実行時コード側の依存を増やさずに済みます。`e2m_ico.ico` はEXEアイコンおよび実行時のタスクトレイアイコンとして同梱してください。`config.ini` はショートカット変更用にEXEと同じフォルダーへ配置してください。

```powershell
python -m pip install pyinstaller
pyinstaller --onefile --noconsole --name ExceltoMarkdown --icon e2m_ico.ico --add-data "e2m_ico.ico;." excel_to_markdown.py
```

書式反映も含めて配布したい場合は、ビルド環境に `pywin32` もインストールしてからPyInstallerを実行してください。

## 開発・テスト

```bash
python -m unittest
```

変換ロジックの単体テストと `--stdin` 変換はLinux/Unix上でも実行できます。クリップボード操作、Excel COM連携、タスクトレイ常駐はWindows APIを使うためWindows上で利用してください。
