"""OSに依存しないGitHub Flavored Markdown表変換です。"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from typing import Iterable, Sequence
from urllib.parse import quote


@dataclass(frozen=True)
class Cell:
    """セルの値とGFMへ反映する簡易書式情報を保持します。"""

    value: object = ""
    bold: bool = False
    italic: bool = False
    href: str | None = None


def escape_link_destination(value: object) -> str:
    """Markdownリンクの行先を、括弧や空白で構文が壊れない形へ変換します。"""

    if value is None:
        return ""
    # URLとして意味を持つ区切り文字と既存のpercent encodingは保ちつつ、
    # Markdownのリンク終端と衝突する括弧、空白、制御文字などをpercent encodeします。
    return quote(
        str(value).strip(),
        safe=":/?#[]@!$&'*,;=+%~-._",
    )


def escape_markdown_cell(value: object) -> str:
    """GFM tableの列・リンク構文を壊す文字をセル内でエスケープします。"""

    text = "" if value is None else str(value)
    # セル内改行はGFM tableの行区切りと衝突するため、HTMLの改行タグへ寄せます。
    text = text.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    text = text.replace("\\", "\\\\")
    text = text.replace("|", "\\|")
    text = text.replace("[", "\\[").replace("]", "\\]")
    text = text.replace("(", "\\(").replace(")", "\\)")
    return text.strip()


def format_cell(cell: Cell | object) -> str:
    """セル値にリンク・イタリック・太字を適用したGFM文字列を返します。"""

    if not isinstance(cell, Cell):
        cell = Cell(cell)
    text = escape_markdown_cell(cell.value)
    if cell.href and text:
        # リンクテキスト内の丸括弧はリンク構文と衝突しないため、表示用に元へ戻します。
        link_text = text.replace("\\(", "(").replace("\\)", ")")
        if cell.bold and cell.italic:
            link_text = f"***{link_text}***"
        elif cell.italic:
            link_text = f"*{link_text}*"
        elif cell.bold:
            link_text = f"**{link_text}**"
        return f"[{link_text}]({escape_link_destination(cell.href)})"
    if cell.bold and cell.italic and text:
        return f"***{text}***"
    if cell.italic and text:
        text = f"*{text}*"
    if cell.bold and text:
        text = f"**{text}**"
    return text


def normalize_rows(rows: Iterable[Sequence[Cell | object]]) -> list[list[Cell | object]]:
    """行ごとの列数をそろえ、GFM tableとして崩れない矩形データにします。"""

    normalized = [list(row) for row in rows]
    if not normalized:
        return []
    width = max(len(row) for row in normalized)
    return [row + [Cell()] * (width - len(row)) for row in normalized]


def tsv_to_rows(text: str) -> list[list[Cell]]:
    """Excel由来のquoted TSVを行・セルへ分解します。

    Excelがセル内のタブ、改行、二重引用符を保護するために使うquoted fieldを
    Python標準のCSVパーサーで解釈します。
    """

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Excelコピーでは末尾に改行が付くため、余分な空行として出力しないよう除去します。
    # quoted field内の末尾改行は閉じ引用符より前にあるので、この処理では失われません。
    text = text.rstrip("\n")
    if not text:
        return []
    reader = csv.reader(
        io.StringIO(text),
        delimiter="\t",
        quotechar='"',
        doublequote=True,
        skipinitialspace=False,
    )
    return [[Cell(value) for value in row] for row in reader]


def rows_to_markdown(rows: Iterable[Sequence[Cell | object]]) -> str:
    """行データをGitHub Flavored Markdown (GFM) tableへ変換します。

    GFM tableの仕様に従い、入力の先頭行は必ずheader rowとして扱います。
    """

    table = normalize_rows(rows)
    if not table:
        return ""
    formatted = [[format_cell(cell) for cell in row] for row in table]
    width = len(formatted[0])
    separator = ["---"] * width
    lines = [
        "| " + " | ".join(formatted[0]) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in formatted[1:])
    return "\n".join(lines) + "\n"


def convert_text_to_markdown(text: str) -> str:
    """Excel由来のTSVテキストをGFM tableへ変換します。"""

    return rows_to_markdown(tsv_to_rows(text))


def convert_table_to_markdown(text: str) -> str:
    """表変換経路を明示するための別名です。"""

    return convert_text_to_markdown(text)
