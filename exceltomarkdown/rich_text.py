"""HTML clipboard断片を読みやすいGitHub Flavored Markdownへ変換します。"""

from __future__ import annotations

import re
from html.parser import HTMLParser

from .core import escape_link_destination


HTML_VOID_TAGS = {
    "area",
    "base",
    "br",
    "col",
    "embed",
    "hr",
    "img",
    "input",
    "link",
    "meta",
    "param",
    "source",
    "track",
    "wbr",
}


def _escape_block_prefix(text: str) -> str:
    """文章の先頭文字が意図しないGFM blockとして解釈されることを防ぎます。"""

    text = re.sub(r"^( {0,3})(#{1,6})(?=\s|$)", r"\1\\\2", text)
    text = re.sub(r"^( {0,3})(>)(?=\s|$)", r"\1\\\2", text)
    text = re.sub(r"^( {0,3})([-+*])(?=\s)", r"\1\\\2", text)
    text = re.sub(r"^( {0,3}\d{1,9})([.)])(?=\s)", r"\1\\\2", text)
    if re.fullmatch(r" {0,3}(-\s*){3,}", text):
        text = text.replace("-", "\\-", 1)
    return text


def escape_markdown_text(value: object, *, at_line_start: bool = False) -> str:
    """通常文章用のGFMエスケープを行います。

    inline構文に加え、行頭では見出し、引用、リスト、水平線への意図しない
    semantic transformationも防ぎます。
    """

    text = "" if value is None else str(value)
    for character in ("\\", "*", "_", "`", "[", "]", "<", ">", "~"):
        text = text.replace(character, "\\" + character)
    if at_line_start:
        text = _escape_block_prefix(text)
    return text


class _RichHtmlToMarkdownParser(HTMLParser):
    """ブラウザやOffice由来の簡易HTML断片をGFM文章へ変換するパーサーです。"""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._parts: list[str] = []
        self._list_stack: list[dict[str, int | str]] = []
        self._link_stack: list[str] = []
        self._style_marker_stack: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "br":
            self._ensure_line_break()
            return
        if tag in HTML_VOID_TAGS:
            return
        if tag in {"p", "div"}:
            self._ensure_paragraph_break()
        elif tag in {"ul", "ol"}:
            self._ensure_line_break()
            self._list_stack.append({"tag": tag, "index": 1})
        elif tag == "li":
            self._start_list_item()

        if tag in {"strong", "b"}:
            self._parts.append("**")
        elif tag in {"em", "i"}:
            self._parts.append("*")
        elif tag == "a":
            href = ""
            for name, value in attrs:
                if name and name.lower() == "href" and value:
                    href = escape_link_destination(value)
                    break
            self._link_stack.append(href)
            self._parts.append("[")

        style_markers = self._style_markers(
            attrs,
            skip_bold=tag in {"strong", "b"},
            skip_italic=tag in {"em", "i"},
        )
        self._style_marker_stack.append(style_markers)
        self._parts.extend(style_markers)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in HTML_VOID_TAGS:
            return
        style_markers = self._style_marker_stack.pop() if self._style_marker_stack else []
        self._parts.extend(reversed(style_markers))
        if tag in {"strong", "b"}:
            self._parts.append("**")
        elif tag in {"em", "i"}:
            self._parts.append("*")
        elif tag == "a":
            href = self._link_stack.pop() if self._link_stack else ""
            self._parts.append(f"]({href})" if href else "]")
        elif tag in {"p", "div"}:
            self._ensure_paragraph_break()
        elif tag == "li":
            self._ensure_line_break()
        elif tag in {"ul", "ol"}:
            if self._list_stack:
                self._list_stack.pop()
            self._ensure_paragraph_break()

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "br":
            self._ensure_line_break()
        elif tag in HTML_VOID_TAGS:
            return
        else:
            self.handle_starttag(tag, attrs)
            self.handle_endtag(tag)

    def handle_data(self, data: str) -> None:
        if not data:
            return
        text = re.sub(r"\s+", " ", data)
        if not text.strip():
            if self._parts and not self._parts[-1].endswith((" ", "\n", "[")):
                self._parts.append(" ")
            return
        current = "".join(self._parts)
        self._parts.append(escape_markdown_text(text, at_line_start=not current or current.endswith("\n")))

    def markdown(self) -> str:
        text = "".join(self._parts)
        lines = [line.rstrip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        compacted: list[str] = []
        blank_count = 0
        for line in lines:
            if line.strip():
                compacted.append(line.rstrip())
                blank_count = 0
            else:
                blank_count += 1
                if blank_count <= 1 and compacted:
                    compacted.append("")
        while compacted and not compacted[-1]:
            compacted.pop()
        return "\n".join(compacted) + ("\n" if compacted else "")

    def _ensure_line_break(self) -> None:
        if not self._parts:
            return
        current = "".join(self._parts)
        if not current.endswith("\n"):
            self._parts.append("\n")

    def _ensure_paragraph_break(self) -> None:
        if not self._parts:
            return
        current = "".join(self._parts).rstrip(" ")
        if not current or current.endswith("\n\n"):
            return
        self._parts.append("\n" if current.endswith("\n") else "\n\n")

    def _start_list_item(self) -> None:
        self._ensure_line_break()
        indent = "  " * max(len(self._list_stack) - 1, 0)
        marker = "- "
        if self._list_stack and self._list_stack[-1]["tag"] == "ol":
            index = int(self._list_stack[-1]["index"])
            marker = f"{index}. "
            self._list_stack[-1]["index"] = index + 1
        self._parts.append(f"{indent}{marker}")

    def _style_markers(
        self,
        attrs: list[tuple[str, str | None]],
        *,
        skip_bold: bool = False,
        skip_italic: bool = False,
    ) -> list[str]:
        markers: list[str] = []
        style = ""
        for name, value in attrs:
            if name and name.lower() == "style" and value:
                style = value
                break
        if not style:
            return markers
        declarations: dict[str, str] = {}
        for declaration in style.split(";"):
            if ":" not in declaration:
                continue
            property_name, property_value = declaration.split(":", 1)
            declarations[property_name.strip().lower()] = property_value.strip().lower()
        if not skip_bold and _is_bold_weight(declarations.get("font-weight", "")):
            markers.append("**")
        if not skip_italic and declarations.get("font-style", "") == "italic":
            markers.append("*")
        return markers


def _is_bold_weight(value: str) -> bool:
    """CSSのfont-weight値がGFM太字として扱う重みかを返します。"""

    normalized = value.strip().lower()
    if normalized in {"bold", "bolder"}:
        return True
    if normalized.isdigit():
        return int(normalized) >= 600
    return False


def convert_html_fragment_to_markdown(html_fragment: str) -> str:
    """HTML断片を依存ライブラリなしでGFM文章へ変換します。"""

    parser = _RichHtmlToMarkdownParser()
    parser.feed(html_fragment)
    parser.close()
    return parser.markdown()


def convert_html_to_markdown(text: str) -> str:
    """HTML文字列をGFM文章へ変換します。"""

    return convert_html_fragment_to_markdown(text)


def convert_rich_text_to_markdown(text: str) -> str:
    """後方互換用のHTML文字列変換別名です。"""

    return convert_html_to_markdown(text)
