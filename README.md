
Excelでコピーした表をMarkdownテーブルに変換し、クリップボードへ戻すWindows向けツールです。

## できること

- Windows専用です。
- Excelで範囲をコピーした後、タスクトレイ常駐アプリからMarkdownテーブルへ変換します。
- ホットキー `Ctrl+Alt+M`、タスクトレイの右クリックメニュー、またはトレイアイコンのダブルクリックで変換できます。
- 標準ライブラリのみでクリップボード内のTSVをMarkdown化できます。
- 任意で `pywin32` を入れると、起動中のExcelの選択範囲から太字、イタリック、ハイパーリンクを読み取り、Markdownへ反映します。
- タスクトレイと実行ファイルのアイコンには、同梱の `e2m_ico.ico を利用します。

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
2. `Ctrl+Alt+M` を押すか、タスクトレイアイコンのメニューから「Markdownに変換」を選びます。
3. クリップボードの内容がMarkdownテーブルに置き換わるので、`.md` ファイルへ貼り付けます。

### 1回だけ変換して終了

```powershell
python excel_to_markdown.py --once
```

### 標準入力で変換を確認

```powershell
Get-Content sample.tsv -Raw | python excel_to_markdown.py --stdin
```

## Excel書式の反映（任意）

太字、イタリック、リンクを反映したい場合のみ、軽量な標準機能の範囲を超えるため `pywin32` を追加してください。

```powershell
python -m pip install pywin32
```

`pywin32` がない環境では、Excelがクリップボードに置いたプレーンテキスト（TSV）だけを使って変換します。

## Windows Executableとして配布する例

配布用EXEを作る場合は、ビルド環境だけにPyInstallerを入れる構成にすると、実行時コード側の依存を増やさずに済みます。`E2M.ico` はEXEアイコンおよび実行時のタスクトレイアイコンとして同梱してください。

```powershell
python -m pip install pyinstaller
pyinstaller --onefile --noconsole --name ExceltoMarkdown --icon E2M.ico --add-data "E2M.ico;." excel_to_markdown.py
```

書式反映も含めて配布したい場合は、ビルド環境に `pywin32` もインストールしてからPyInstallerを実行してください。

## 開発・テスト

```bash
python -m unittest
```

変換ロジックの単体テストはLinux/Unix上でも実行できますが、アプリ本体（CLIの `main`、クリップボード操作、タスクトレイ常駐）はWindows以外では動作しません。
