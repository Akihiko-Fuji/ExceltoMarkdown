"""Excel由来のTSVテキストやExcel選択範囲をMarkdown表へ変換します。

標準入力のTSV変換はOSを問わず利用できます。Windowsではクリップボード、
ホットキー、タスクトレイを使った常駐アプリとしても動作します。任意の
``pywin32`` がインストールされている場合のみ、COM経由でExcelのアクティブ
な選択範囲を読み取り、太字・イタリック・ハイパーリンクなどの簡単な書式
もMarkdownへ反映します。
"""

from __future__ import annotations

import argparse
import configparser
import ctypes
import importlib
import importlib.util
import io
import re
import sys
import threading
import time
import traceback
from pathlib import Path
from dataclasses import dataclass
from ctypes import wintypes
from html.parser import HTMLParser
from typing import Iterable, Sequence

# WindowsではWINFUNCTYPE、非Windowsのテスト環境ではCFUNCTYPEを使います。
WINDOW_CALLBACK = getattr(ctypes, "WINFUNCTYPE", ctypes.CFUNCTYPE)
LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)
# Python 3.11 の非Windows環境では一部のWin32ハンドル型が未定義のため、
# 構造体定義をimport時に評価できるよう公開されているHANDLE型へフォールバックします。
HCURSOR = getattr(wintypes, "HCURSOR", wintypes.HANDLE)

# Win32メッセージ処理とタスクトレイ操作に使う定数です。
HOTKEY_ID = 0x4D44
WM_APP = 0x8000
WM_TRAYICON = WM_APP + 1
WM_COMMAND = 0x0111
WM_DESTROY = 0x0002
WM_HOTKEY = 0x0312
WM_RBUTTONUP = 0x0205
WM_LBUTTONDBLCLK = 0x0203
NIM_ADD = 0x00000000
NIM_DELETE = 0x00000002
NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
TPM_RIGHTBUTTON = 0x0002
TPM_RETURNCMD = 0x0100
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x00000010
LR_DEFAULTSIZE = 0x00000040
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
HTML_CLIPBOARD_FORMAT_NAME = "HTML Format"
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

MENU_CONVERT_TABLE = 1001
MENU_CONVERT_RICH_TEXT = 1002
MENU_EXIT = 1003
# 後方互換のため、既存テストや外部利用のMENU_CONVERT名は表変換コマンドとして残します。
MENU_CONVERT = MENU_CONVERT_TABLE
APP_ICON_FILENAME = "e2m_ico.ico"
CONFIG_FILENAME = "config.ini"
DEFAULT_HOTKEY = "Ctrl+Alt+M"
DEFAULT_PREFER_EXCEL = False
DEFAULT_CONVERSION_MODE = "table"
CONVERSION_MODE_TABLE = "table"
CONVERSION_MODE_RICH_TEXT = "rich_text"
CONVERSION_MODE_AUTO = "auto"
SUPPORTED_CONVERSION_MODES = {CONVERSION_MODE_TABLE, CONVERSION_MODE_RICH_TEXT}
INTERNAL_CONVERSION_MODES = {*SUPPORTED_CONVERSION_MODES, CONVERSION_MODE_AUTO}
MAX_EXCEL_SELECTION_CELLS = 5000
CLIPBOARD_OPEN_RETRIES = 10
CLIPBOARD_OPEN_DELAY_SECONDS = 0.05

# 非Windows環境でも変換ロジックの単体テストを実行できるよう、DLL参照を遅延させます。
if sys.platform == "win32":
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    shell32 = ctypes.WinDLL("shell32", use_last_error=True)
else:
    user32 = kernel32 = shell32 = None


class POINT(ctypes.Structure):
    """右クリックメニューを表示するカーソル座標を保持するWin32構造体です。"""

    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSG(ctypes.Structure):
    """Windowsのメッセージループで受け取るイベント情報の構造体です。"""

    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


class WNDCLASS(ctypes.Structure):
    """非表示ウィンドウを登録するためのWin32ウィンドウクラス構造体です。"""

    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WINDOW_CALLBACK(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", HCURSOR),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class NOTIFYICONDATA(ctypes.Structure):
    """タスクトレイアイコンの登録・削除に使う通知領域データ構造体です。"""

    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
    ]


def _configure_windows_api() -> None:
    """Win32 API呼び出しの引数型・戻り値型を明示して安全に扱います。"""

    if sys.platform != "win32":
        return
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
    user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user32.RegisterClipboardFormatW.restype = wintypes.UINT
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.LoadImageW.argtypes = [
        wintypes.HINSTANCE,
        wintypes.LPCWSTR,
        wintypes.UINT,
        ctypes.c_int,
        ctypes.c_int,
        wintypes.UINT,
    ]
    user32.LoadImageW.restype = wintypes.HANDLE
    user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASS)]
    user32.RegisterClassW.restype = wintypes.ATOM
    user32.CreateWindowExW.argtypes = [
        wintypes.DWORD,
        wintypes.LPCWSTR,
        wintypes.LPCWSTR,
        wintypes.DWORD,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        wintypes.HWND,
        wintypes.HMENU,
        wintypes.HINSTANCE,
        wintypes.LPVOID,
    ]
    user32.CreateWindowExW.restype = wintypes.HWND
    user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.DefWindowProcW.restype = LRESULT
    user32.DestroyWindow.argtypes = [wintypes.HWND]
    user32.DestroyWindow.restype = wintypes.BOOL
    user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
    user32.RegisterHotKey.restype = wintypes.BOOL
    user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.UnregisterHotKey.restype = wintypes.BOOL
    user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
    user32.GetMessageW.restype = wintypes.BOOL
    user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]
    user32.TranslateMessage.restype = wintypes.BOOL
    user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]
    user32.DispatchMessageW.restype = LRESULT
    user32.PostQuitMessage.argtypes = [ctypes.c_int]
    user32.PostQuitMessage.restype = None
    user32.CreatePopupMenu.argtypes = []
    user32.CreatePopupMenu.restype = wintypes.HMENU
    user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT, wintypes.WPARAM, wintypes.LPCWSTR]
    user32.AppendMenuW.restype = wintypes.BOOL
    user32.TrackPopupMenu.argtypes = [
        wintypes.HMENU,
        wintypes.UINT,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        wintypes.HWND,
        wintypes.LPVOID,
    ]
    user32.TrackPopupMenu.restype = wintypes.UINT
    user32.DestroyMenu.argtypes = [wintypes.HMENU]
    user32.DestroyMenu.restype = wintypes.BOOL
    user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
    user32.GetCursorPos.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.MessageBeep.argtypes = [wintypes.UINT]
    user32.MessageBeep.restype = wintypes.BOOL
    user32.MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
    user32.MessageBoxW.restype = ctypes.c_int
    shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATA)]
    shell32.Shell_NotifyIconW.restype = wintypes.BOOL
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalSize.restype = ctypes.c_size_t
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL


def is_windows() -> bool:
    """このアプリを動作対象にしているWindows環境かどうかを返します。"""

    return sys.platform == "win32"


def resource_path(filename: str) -> Path:
    """通常実行時とPyInstaller実行時の両方で同梱リソースのパスを解決します。"""

    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / filename


def icon_path() -> Path:
    """タスクトレイと実行ファイルで使うe2m_ico.icoのパスを返します。"""

    return resource_path(APP_ICON_FILENAME)


def app_base_dir() -> Path:
    """実行中アプリと同じ場所にある設定ファイルの基準ディレクトリを返します。"""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    """ショートカット設定を読み取るconfig.iniのパスを返します。"""

    return app_base_dir() / CONFIG_FILENAME


def load_application_icon() -> wintypes.HANDLE:
    """e2m_ico.icoをWin32アイコンとして読み込みます。"""

    _require_windows()
    path = icon_path()
    if not path.exists():
        raise FileNotFoundError(f"アイコンファイルが見つかりません: {path}")
    icon = user32.LoadImageW(
        None,
        str(path),
        IMAGE_ICON,
        0,
        0,
        LR_LOADFROMFILE | LR_DEFAULTSIZE,
    )
    if not icon:
        raise ctypes.WinError(ctypes.get_last_error())
    return icon


_configure_windows_api()


@dataclass(frozen=True)
class HotkeyConfig:
    """WindowsのRegisterHotKeyへ渡すショートカット設定です。"""

    label: str
    modifiers: int
    virtual_key: int


@dataclass(frozen=True)
class Cell:
    """セルの値とMarkdownへ反映したい簡易書式情報を保持します。"""

    value: object = ""
    bold: bool = False
    italic: bool = False
    href: str | None = None


_MODIFIER_ALIASES = {
    "ctrl": ("Ctrl", MOD_CONTROL),
    "control": ("Ctrl", MOD_CONTROL),
    "alt": ("Alt", MOD_ALT),
    "shift": ("Shift", MOD_SHIFT),
    "win": ("Win", MOD_WIN),
    "windows": ("Win", MOD_WIN),
}

_KEY_ALIASES = {
    "enter": 0x0D,
    "return": 0x0D,
    "esc": 0x1B,
    "escape": 0x1B,
    "space": 0x20,
    "tab": 0x09,
    "backspace": 0x08,
    "delete": 0x2E,
    "del": 0x2E,
    "insert": 0x2D,
    "ins": 0x2D,
    "home": 0x24,
    "end": 0x23,
    "pageup": 0x21,
    "pagedown": 0x22,
    "pgup": 0x21,
    "pgdn": 0x22,
    "up": 0x26,
    "down": 0x28,
    "left": 0x25,
    "right": 0x27,
}

_FUNCTION_KEY_ALIASES = {f"f{_number}": 0x70 + _number - 1 for _number in range(1, 25)}
_KEY_ALIASES.update(_FUNCTION_KEY_ALIASES)


def _parse_virtual_key(key_name: str) -> tuple[str, int]:
    """設定文字列のキー名をWin32仮想キーコードへ変換します。"""

    normalized = key_name.strip().lower()
    if len(normalized) == 1 and normalized.isascii() and normalized.isalnum():
        # Win32の仮想キーコードとASCII英数字のコード値が一致する範囲だけをord()で扱います。
        return normalized.upper(), ord(normalized.upper())
    if normalized in _FUNCTION_KEY_ALIASES:
        return f"F{normalized[1:]}", _FUNCTION_KEY_ALIASES[normalized]
    if normalized in _KEY_ALIASES:
        return key_name.strip(), _KEY_ALIASES[normalized]
    raise ValueError(f"未対応のショートカットキーです: {key_name}")


def parse_hotkey(value: str) -> HotkeyConfig:
    """Ctrl+Alt+M形式の設定値をRegisterHotKey用の値へ変換します。"""

    parts = [part.strip() for part in value.replace("-", "+").split("+") if part.strip()]
    if not parts:
        raise ValueError("ショートカットキーが空です。")
    modifiers = 0
    modifier_labels: list[str] = []
    key_label: str | None = None
    virtual_key: int | None = None
    for part in parts:
        normalized = part.lower()
        if normalized in _MODIFIER_ALIASES:
            label, flag = _MODIFIER_ALIASES[normalized]
            if modifiers & flag:
                continue
            modifiers |= flag
            modifier_labels.append(label)
            continue
        if virtual_key is not None:
            raise ValueError(f"ショートカットの通常キーは1つだけ指定してください: {value}")
        key_label, virtual_key = _parse_virtual_key(part)
    if virtual_key is None or key_label is None:
        raise ValueError(f"ショートカットには通常キーを1つ指定してください: {value}")
    if modifiers == 0:
        raise ValueError(f"ショートカットにはCtrl/Alt/Shift/Winのいずれかを指定してください: {value}")
    return HotkeyConfig("+".join([*modifier_labels, key_label]), modifiers, virtual_key)


def _get_section_own_options(parser: configparser.ConfigParser, section: str) -> set[str]:
    """ConfigParserの公開APIだけで、指定セクション直下のオプション名を取得します。"""

    buffer = io.StringIO()
    parser.write(buffer)
    current_section = None
    own_options: set[str] = set()
    for line in buffer.getvalue().splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith(("#", ";")):
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            current_section = stripped[1:-1]
            continue
        if current_section != section or line[:1].isspace():
            continue

        delimiters = [index for index in (line.find("="), line.find(":")) if index != -1]
        if not delimiters:
            continue
        option_name = line[: min(delimiters)].strip()
        own_options.add(parser.optionxform(option_name))
    return own_options


def _get_explicit_config_option(parser: configparser.ConfigParser, section: str, option: str) -> str | None:
    """DEFAULTから継承された値ではなく、セクション直下に書かれた値だけを返します。"""

    option_norm = parser.optionxform(option)
    if section == parser.default_section:
        return parser.defaults().get(option_norm)
    if not parser.has_section(section):
        return None

    section_own_keys = _get_section_own_options(parser, section)
    if option_norm not in section_own_keys:
        return None
    return parser.get(section, option)


def load_hotkey_config(path: Path | None = None) -> HotkeyConfig:
    """config.iniからショートカットキー設定を読み込みます。"""

    config_file = path or config_path()
    parser = configparser.ConfigParser()
    if config_file.exists():
        parser.read(config_file, encoding="utf-8")

    value = DEFAULT_HOTKEY
    explicit_shortcut_key = _get_explicit_config_option(parser, "shortcut", "key")
    explicit_hotkey_key = _get_explicit_config_option(parser, "hotkey", "key")
    if explicit_shortcut_key is not None:
        value = explicit_shortcut_key
    elif explicit_hotkey_key is not None:
        value = explicit_hotkey_key
    else:
        explicit_default_shortcut = _get_explicit_config_option(parser, parser.default_section, "shortcut")
        if explicit_default_shortcut is not None:
            value = explicit_default_shortcut
    return parse_hotkey(value)


def _parse_config_bool(value: str) -> bool:
    """config.iniの真偽値文字列をboolへ変換します。

    標準的な 1/yes/true/on と 0/no/false/off に加え、
    enabled/disabled も受け付けます。
    """

    normalized = value.strip().lower()
    if normalized in {"1", "yes", "true", "on", "enabled"}:
        return True
    if normalized in {"0", "no", "false", "off", "disabled"}:
        return False
    raise ValueError(f"真偽値として解釈できない設定値です: {value}")


def load_prefer_excel_config(path: Path | None = None) -> bool:
    """Excel選択範囲をクリップボードTSVより優先するかをconfig.iniから読み込みます。"""

    config_file = path or config_path()
    parser = configparser.ConfigParser()
    if config_file.exists():
        parser.read(config_file, encoding="utf-8")

    explicit_value = _get_explicit_config_option(parser, "conversion", "prefer_excel")
    if explicit_value is None:
        explicit_value = _get_explicit_config_option(parser, parser.default_section, "prefer_excel")
    if explicit_value is None:
        return DEFAULT_PREFER_EXCEL
    return _parse_config_bool(explicit_value)


def normalize_conversion_mode(value: str) -> str:
    """config.iniやCLIから受け取った変換モード名を正規化します。"""

    normalized = value.strip().lower().replace("-", "_")
    if normalized not in SUPPORTED_CONVERSION_MODES:
        raise ValueError(f"未対応の変換モードです: {value}")
    return normalized


def load_default_conversion_mode_config(path: Path | None = None) -> str:
    """ホットキーやダブルクリックで使う既定変換モードをconfig.iniから読み込みます。"""

    config_file = path or config_path()
    parser = configparser.ConfigParser()
    if config_file.exists():
        parser.read(config_file, encoding="utf-8")

    explicit_value = _get_explicit_config_option(parser, "conversion", "default_mode")
    if explicit_value is None:
        explicit_value = _get_explicit_config_option(parser, parser.default_section, "default_mode")
    if explicit_value is None:
        return DEFAULT_CONVERSION_MODE
    return normalize_conversion_mode(explicit_value)


def escape_markdown_cell(value: object) -> str:
    """Markdown表の列区切りを壊す文字をセル内で安全にエスケープします。"""

    text = "" if value is None else str(value)
    # セル内改行はMarkdown表の行区切りと衝突するため、HTMLの改行タグへ寄せます。
    text = text.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    text = text.replace("\\", "\\\\")
    text = text.replace("|", "\\|")
    text = text.replace("[", "\\[").replace("]", "\\]")
    text = text.replace("(", "\\(").replace(")", "\\)")
    return text.strip()


def format_cell(cell: Cell | object) -> str:
    """セル値にリンク・イタリック・太字を適用したMarkdown文字列を返します。"""

    if not isinstance(cell, Cell):
        cell = Cell(cell)
    text = escape_markdown_cell(cell.value)
    if cell.href and text:
        # リンクテキスト内の丸括弧はMarkdownリンク構文を壊さないため、表示用に元へ戻します。
        link_text = text.replace("\\(", "(").replace("\\)", ")")
        if cell.bold and cell.italic:
            link_text = f"***{link_text}***"
        elif cell.italic:
            link_text = f"*{link_text}*"
        elif cell.bold:
            link_text = f"**{link_text}**"
        # URL内の閉じ括弧はMarkdownリンク構文を壊すため、最低限エンコードします。
        href = str(cell.href).replace(")", "%29")
        return f"[{link_text}]({href})"
    if cell.bold and cell.italic and text:
        return f"***{text}***"
    if cell.italic and text:
        text = f"*{text}*"
    if cell.bold and text:
        text = f"**{text}**"
    return text


def normalize_rows(rows: Iterable[Sequence[Cell | object]]) -> list[list[Cell | object]]:
    """行ごとの列数をそろえ、Markdown表として崩れない矩形データにします。"""

    normalized = [list(row) for row in rows]
    if not normalized:
        return []
    width = max(len(row) for row in normalized)
    return [row + [Cell()] * (width - len(row)) for row in normalized]


def tsv_to_rows(text: str) -> list[list[Cell]]:
    """Excelがクリップボードへ置くタブ区切りテキストを行・セルへ分解します。"""

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Excelコピーでは末尾に改行が付くことが多いため、末尾の連続空行として扱わないよう除去します。
    text = text.rstrip("\n")
    if not text:
        return []
    return [[Cell(value) for value in line.split("\t")] for line in text.split("\n")]


def rows_to_markdown(rows: Iterable[Sequence[Cell | object]]) -> str:
    """行データをGitHub Flavored Markdown互換の表へ変換します。"""

    table = normalize_rows(rows)
    if not table:
        return ""
    formatted = [[format_cell(cell) for cell in row] for row in table]
    # Markdown表ではヘッダー直下に区切り行が必須です。
    width = len(formatted[0])
    separator = ["---"] * width
    lines = [
        "| " + " | ".join(formatted[0]) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in formatted[1:])
    return "\n".join(lines) + "\n"


def convert_text_to_markdown(text: str) -> str:
    """Excel由来のTSVテキストをMarkdown表へ変換します。"""

    return rows_to_markdown(tsv_to_rows(text))


def convert_table_to_markdown(text: str) -> str:
    """表変換経路を明示するための別名です。既存TSV変換ロジックをそのまま使います。"""

    return convert_text_to_markdown(text)


def escape_markdown_text(value: object) -> str:
    """通常文章用のMarkdownエスケープを行います。表セル用とは独立させます。"""

    text = "" if value is None else str(value)
    for character in ("\\", "*", "_", "`", "[", "]"):
        text = text.replace(character, "\\" + character)
    return text


class _RichHtmlToMarkdownParser(HTMLParser):
    """ブラウザやOffice由来の簡易HTML断片をMarkdown文章へ変換するパーサーです。"""

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
                    href = value.replace(")", "%29")
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
            self.handle_starttag(tag, attrs)
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
        self._parts.append(escape_markdown_text(text))

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
        if not current:
            return
        if current.endswith("\n\n"):
            return
        if current.endswith("\n"):
            self._parts.append("\n")
        else:
            self._parts.append("\n\n")

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
        font_weight = declarations.get("font-weight", "")
        font_style = declarations.get("font-style", "")
        if not skip_bold and _is_bold_weight(font_weight):
            markers.append("**")
        if not skip_italic and font_style == "italic":
            markers.append("*")
        return markers


def _is_bold_weight(value: str) -> bool:
    """CSSのfont-weight値がMarkdown太字として扱う重みかを返します。"""

    normalized = value.strip().lower()
    if normalized in {"bold", "bolder"}:
        return True
    if normalized.isdigit():
        return int(normalized) >= 600
    return False


def convert_html_fragment_to_markdown(html_fragment: str) -> str:
    """HTML断片を依存ライブラリなしでMarkdown文章へ変換します。"""

    parser = _RichHtmlToMarkdownParser()
    parser.feed(html_fragment)
    parser.close()
    return parser.markdown()


def convert_html_to_markdown(text: str) -> str:
    """HTML文字列をMarkdown文章へ変換します。"""

    return convert_html_fragment_to_markdown(text)


def convert_rich_text_to_markdown(text: str) -> str:
    """後方互換用のHTML文字列変換別名です。"""

    return convert_html_to_markdown(text)


def _require_windows() -> None:
    """Win32 APIが必要な処理をWindows以外で実行しないようにします。"""

    if not is_windows():
        # 例外はログに残る可能性があるため、開発者向けメッセージとして英語に統一します。
        raise RuntimeError("Windows clipboard/task tray APIs are required for this operation.")


def open_clipboard_with_retry(retries: int = CLIPBOARD_OPEN_RETRIES, delay: float = CLIPBOARD_OPEN_DELAY_SECONDS) -> None:
    """他アプリが一瞬クリップボードを掴んでいる場合に備えてOpenClipboardを短く再試行します。

    成功時はクリップボードを開いたまま返すため、呼び出し元は必ず
    ``user32.CloseClipboard()`` で解放してください。
    """

    _require_windows()
    for _ in range(retries):
        if user32.OpenClipboard(None):
            return
        time.sleep(delay)
    raise OSError("Could not open clipboard.")


def read_clipboard_text() -> str:
    """WindowsクリップボードからUnicodeテキストを読み取ります。"""

    open_clipboard_with_retry()
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        pointer = kernel32.GlobalLock(handle)
        if not pointer:
            raise MemoryError("Could not lock clipboard memory.")
        try:
            return ctypes.wstring_at(pointer)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def register_clipboard_format(format_name: str) -> int:
    """Windowsクリップボードの名前付き形式を登録し、形式IDを返します。"""

    _require_windows()
    format_id = user32.RegisterClipboardFormatW(format_name)
    if not format_id:
        raise ctypes.WinError(ctypes.get_last_error())
    return int(format_id)


def clipboard_has_format(format_name: str) -> bool:
    """指定した名前付きクリップボード形式が現在利用可能かを返します。"""

    format_id = register_clipboard_format(format_name)
    return bool(user32.IsClipboardFormatAvailable(format_id))


def _read_clipboard_format_bytes(format_id: int) -> bytes:
    """開いているクリップボードから任意形式のバイト列を読み取ります。"""

    handle = user32.GetClipboardData(format_id)
    if not handle:
        return b""
    size = int(kernel32.GlobalSize(handle))
    pointer = kernel32.GlobalLock(handle)
    if not pointer:
        raise MemoryError("Could not lock clipboard memory.")
    try:
        if size <= 0:
            return ctypes.string_at(pointer).rstrip(b"\0")
        return ctypes.string_at(pointer, size).rstrip(b"\0")
    finally:
        kernel32.GlobalUnlock(handle)


def extract_clipboard_html_fragment(clipboard_html: bytes | str) -> str:
    """Clipboard HTML Format全体からStartFragment〜EndFragmentのHTML断片を抽出します。"""

    raw = clipboard_html.encode("utf-8") if isinstance(clipboard_html, str) else clipboard_html
    header_text = raw[: min(len(raw), 4096)].decode("ascii", errors="ignore")
    offsets: dict[str, int] = {}
    for name in ("StartHTML", "EndHTML", "StartFragment", "EndFragment"):
        match = re.search(rf"{name}:\s*(\d+)", header_text, flags=re.IGNORECASE)
        if match:
            offsets[name] = int(match.group(1))
    start = offsets.get("StartFragment")
    end = offsets.get("EndFragment")
    if start is None or end is None:
        start = offsets.get("StartHTML")
        end = offsets.get("EndHTML")
    if start is None or end is None:
        raise ValueError("Clipboard HTML Format header does not contain fragment offsets.")
    if start < 0 or end < start or end > len(raw):
        raise ValueError(f"Clipboard HTML Format offsets are invalid: start={start}, end={end}, length={len(raw)}")
    return raw[start:end].decode("utf-8")


def read_clipboard_html_fragment() -> str:
    """Windows ClipboardのHTML FormatからHTML断片を読み取ります。"""

    format_id = register_clipboard_format(HTML_CLIPBOARD_FORMAT_NAME)
    open_clipboard_with_retry()
    try:
        if not user32.IsClipboardFormatAvailable(format_id):
            return ""
        data = _read_clipboard_format_bytes(format_id)
    finally:
        user32.CloseClipboard()
    if not data:
        return ""
    return extract_clipboard_html_fragment(data)


def read_clipboard_plain_text() -> str:
    """将来の複数形式分岐から呼び出しやすいプレーンテキスト読取別名です。"""

    return read_clipboard_text()


def _log_rich_text_html_fallback(error: BaseException) -> None:
    """HTML Formatの読取・解析失敗時にプレーンテキストへ戻る理由をログへ残します。"""

    def write_fallback_reason(log_file) -> None:
        print("HTML Formatの取得または解析に失敗したため、プレーンテキストへフォールバックします。", file=log_file)
        traceback.print_exception(type(error), error, error.__traceback__, file=log_file)

    write_error_log(write_fallback_reason)


def convert_clipboard_rich_text_to_markdown() -> str:
    """HTML Formatを優先し、なければプレーンテキストを返すリッチテキスト変換経路です。"""

    markdown = ""
    try:
        if clipboard_has_format(HTML_CLIPBOARD_FORMAT_NAME):
            html_fragment = read_clipboard_html_fragment()
            if html_fragment:
                markdown = convert_html_fragment_to_markdown(html_fragment)
    except Exception as error:
        _log_rich_text_html_fallback(error)
    if not markdown:
        markdown = read_clipboard_plain_text()
    if markdown:
        write_clipboard_text(markdown)
    return markdown


def write_clipboard_text(text: str) -> None:
    """UnicodeテキストをWindowsクリップボードへ書き込みます。"""

    _require_windows()
    encoded = (text + "\0").encode("utf-16-le")
    byte_count = len(encoded)
    handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, byte_count)
    if not handle:
        raise MemoryError("Could not allocate clipboard memory.")
    pointer = kernel32.GlobalLock(handle)
    if not pointer:
        kernel32.GlobalFree(handle)
        raise MemoryError("Could not lock clipboard memory.")
    try:
        ctypes.memmove(pointer, encoded, byte_count)
    finally:
        kernel32.GlobalUnlock(handle)
    try:
        open_clipboard_with_retry()
    except Exception:
        kernel32.GlobalFree(handle)
        raise
    try:
        user32.EmptyClipboard()
        result = user32.SetClipboardData(CF_UNICODETEXT, handle)
        if result:
            handle = None
        else:
            raise OSError("Could not set clipboard data.")
    finally:
        user32.CloseClipboard()
        if handle is not None:
            kernel32.GlobalFree(handle)


def pywin32_available() -> bool:
    """任意機能であるExcel COM連携（pywin32）が利用可能かを返します。"""

    return (
        importlib.util.find_spec("win32com") is not None
        and importlib.util.find_spec("win32com.client") is not None
    )


def _excel_selection_to_rows() -> list[list[Cell]]:
    """pywin32経由でExcelの選択範囲を読み取り、簡易書式もCellへ格納します。"""

    win32com_client = importlib.import_module("win32com.client")
    excel = win32com_client.GetActiveObject("Excel.Application")
    selection = excel.Selection
    row_count = int(selection.Rows.Count)
    col_count = int(selection.Columns.Count)
    cell_count = row_count * col_count
    if cell_count > MAX_EXCEL_SELECTION_CELLS:
        raise ValueError(f"選択範囲が大きすぎます: {row_count} x {col_count}")
    rows: list[list[Cell]] = []
    for row_index in range(1, row_count + 1):
        row: list[Cell] = []
        for col_index in range(1, col_count + 1):
            com_cell = selection.Cells(row_index, col_index)
            value = com_cell.Text
            href = None
            if int(com_cell.Hyperlinks.Count) > 0:
                link = com_cell.Hyperlinks(1)
                address = link.Address or ""
                sub_address = link.SubAddress or ""
                if address and sub_address:
                    href = f"{address}#{sub_address}"
                elif address:
                    href = address
                elif sub_address:
                    href = f"#{sub_address}"
            row.append(
                Cell(
                    value=value,
                    bold=bool(com_cell.Font.Bold),
                    italic=bool(com_cell.Font.Italic),
                    href=href,
                )
            )
        rows.append(row)
    return rows


def write_error_log(callback) -> Path | None:
    """エラーログを書き込み、書き込み先に権限がない場合は静かに諦めます。"""

    log_path = app_base_dir() / "excel_to_markdown_error.log"
    try:
        with open(log_path, "a", encoding="utf-8") as log_file:
            callback(log_file)
    except Exception:
        return None
    return log_path


def _log_excel_selection_fallback(error: BaseException) -> None:
    """Excel選択範囲の取得に失敗してクリップボード変換へ戻る理由をログへ残します。"""

    def write_fallback_reason(log_file) -> None:
        print("Excel選択範囲の取得に失敗したため、クリップボードTSVへフォールバックします。", file=log_file)
        traceback.print_exception(type(error), error, error.__traceback__, file=log_file)

    write_error_log(write_fallback_reason)


def convert_clipboard_or_excel_selection(prefer_excel: bool = DEFAULT_PREFER_EXCEL) -> str:
    """設定に応じてExcel選択範囲、またはクリップボードTSVをMarkdown化します。"""

    if prefer_excel and pywin32_available():
        # pywin32があれば、クリップボードのプレーンテキストより先にExcel本体の選択範囲を試します。
        try:
            markdown = rows_to_markdown(_excel_selection_to_rows())
            if markdown:
                write_clipboard_text(markdown)
                return markdown
        except ValueError:
            raise
        except Exception as error:
            # Excelが起動していない、または選択範囲がRangeではない場合があります。
            # その場合でも、理由をログに残してプレーンテキスト変換へフォールバックします。
            _log_excel_selection_fallback(error)
    markdown = convert_table_to_markdown(read_clipboard_text())
    if markdown:
        write_clipboard_text(markdown)
    return markdown


def convert_clipboard_to_markdown(mode: str = DEFAULT_CONVERSION_MODE, prefer_excel: bool = DEFAULT_PREFER_EXCEL) -> str:
    """指定モードに応じてクリップボード内容をMarkdownへ変換します。"""

    mode = mode.strip().lower().replace("-", "_")
    if mode not in INTERNAL_CONVERSION_MODES:
        raise ValueError(f"未対応の変換モードです: {mode}")
    if mode == CONVERSION_MODE_TABLE:
        return convert_clipboard_or_excel_selection(prefer_excel=prefer_excel)
    if mode == CONVERSION_MODE_RICH_TEXT:
        return convert_clipboard_rich_text_to_markdown()
    # autoは今後の判定強化用です。現時点ではHTML Formatがある場合だけ文章変換を選びます。
    try:
        if clipboard_has_format(HTML_CLIPBOARD_FORMAT_NAME):
            return convert_clipboard_rich_text_to_markdown()
    except Exception as error:
        _log_rich_text_html_fallback(error)
    return convert_clipboard_or_excel_selection(prefer_excel=prefer_excel)


class TrayApplication:
    """ctypesだけで実装した小さなWindows通知領域アプリです。"""

    def __init__(self) -> None:
        """非表示ウィンドウとメッセージ処理コールバックの準備を行います。"""

        _require_windows()
        self.hotkey = load_hotkey_config()
        self.prefer_excel = load_prefer_excel_config()
        self.default_mode = load_default_conversion_mode_config()
        self._convert_lock = threading.Lock()
        self.hinstance = kernel32.GetModuleHandleW(None)
        self.class_name = "ExcelToMarkdownTrayWindow"
        self.hwnd = None
        self._icon = load_application_icon()
        self._wndproc = WINDOW_CALLBACK(
            LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
        )(self._window_proc)

    def run(self) -> None:
        """タスクトレイアイコンとホットキーを登録し、メッセージループを開始します。"""

        self._register_window_class()
        self._create_window()
        if not user32.RegisterHotKey(self.hwnd, HOTKEY_ID, self.hotkey.modifiers, self.hotkey.virtual_key):
            raise ctypes.WinError(ctypes.get_last_error())
        self._add_tray_icon()
        msg = MSG()
        while True:
            result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if result == -1:
                raise ctypes.WinError(ctypes.get_last_error())
            if result == 0:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def _register_window_class(self) -> None:
        """ホットキーやトレイ通知を受けるための非表示ウィンドウクラスを登録します。"""

        wc = WNDCLASS()
        wc.lpfnWndProc = self._wndproc
        wc.hInstance = self.hinstance
        wc.lpszClassName = self.class_name
        wc.hIcon = self._icon
        atom = user32.RegisterClassW(ctypes.byref(wc))
        if not atom and ctypes.get_last_error() != 1410:
            raise ctypes.WinError(ctypes.get_last_error())

    def _create_window(self) -> None:
        """画面には表示しないメッセージ受信用ウィンドウを作成します。"""

        self.hwnd = user32.CreateWindowExW(
            0,
            self.class_name,
            "Excel to Markdown",
            0,
            0,
            0,
            0,
            0,
            None,
            None,
            self.hinstance,
            None,
        )
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())

    def _notify_data(self) -> NOTIFYICONDATA:
        """Shell_NotifyIconWに渡す通知領域アイコン情報を組み立てます。"""

        nid = NOTIFYICONDATA()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATA)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid.uCallbackMessage = WM_TRAYICON
        nid.hIcon = self._icon
        nid.szTip = f"Excel to Markdown ({self.hotkey.label})"
        return nid

    def _add_tray_icon(self) -> None:
        """タスクトレイへアイコンを追加します。"""

        nid = self._notify_data()
        if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
            raise ctypes.WinError(ctypes.get_last_error())

    def _remove_tray_icon(self) -> None:
        """アプリ終了時にタスクトレイアイコンを削除します。"""

        nid = self._notify_data()
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))

    def _show_menu(self) -> None:
        """タスクトレイアイコンの右クリックメニューを表示します。"""

        menu = user32.CreatePopupMenu()
        if not menu:
            return
        user32.AppendMenuW(menu, 0, MENU_CONVERT_TABLE, "Markdown表に変換")
        user32.AppendMenuW(menu, 0, MENU_CONVERT_RICH_TEXT, "リッチテキストをMarkdown化")
        user32.AppendMenuW(menu, 0, MENU_EXIT, "終了 (&X)")
        point = POINT()
        user32.GetCursorPos(ctypes.byref(point))
        user32.SetForegroundWindow(self.hwnd)
        command = user32.TrackPopupMenu(
            menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, point.x, point.y, 0, self.hwnd, None
        )
        user32.DestroyMenu(menu)
        if command == MENU_CONVERT_TABLE:
            self._convert_async(CONVERSION_MODE_TABLE)
        elif command == MENU_CONVERT_RICH_TEXT:
            self._convert_async(CONVERSION_MODE_RICH_TEXT)
        elif command == MENU_EXIT:
            user32.DestroyWindow(self.hwnd)

    def _convert_async(self, mode: str | None = None) -> None:
        """UIのメッセージループを止めないよう、変換処理を別スレッドで実行します。"""

        threading.Thread(target=self._convert_safely, args=(mode,), daemon=True).start()

    def _convert_safely(self, mode: str | None = None) -> None:
        """例外をログへ残しつつ、クリップボード内容をMarkdownへ変換します。"""

        if not self._convert_lock.acquire(blocking=False):
            user32.MessageBeep(0xFFFFFFFF)
            return
        try:
            markdown = convert_clipboard_to_markdown(
                mode or getattr(self, "default_mode", DEFAULT_CONVERSION_MODE),
                prefer_excel=self.prefer_excel,
            )
            if markdown:
                user32.MessageBeep(0xFFFFFFFF)
            else:
                user32.MessageBoxW(self.hwnd, "変換対象のテキストがありません。", "Excel to Markdown", 0x40)
        except Exception:
            log_path = write_error_log(lambda log_file: traceback.print_exc(file=log_file))
            if log_path is None:
                message = "変換に失敗しました。\nログファイルへ書き込めませんでした。"
            else:
                message = f"変換に失敗しました。\n{log_path}"
            user32.MessageBoxW(self.hwnd, message, "Excel to Markdown", 0x10)
        finally:
            self._convert_lock.release()

    def _window_proc(
        self,
        hwnd: wintypes.HWND,
        message: wintypes.UINT,
        wparam: wintypes.WPARAM,
        lparam: wintypes.LPARAM,
    ) -> int:
        """トレイ操作、メニュー選択、ホットキー、終了通知を処理します。"""

        # トレイアイコンの右クリックはメニュー表示、ダブルクリックは即時変換です。
        if message == WM_TRAYICON and lparam in (WM_RBUTTONUP, WM_LBUTTONDBLCLK):
            if lparam == WM_LBUTTONDBLCLK:
                self._convert_async()
            else:
                self._show_menu()
            return 0
        if message == WM_COMMAND:
            command = wparam & 0xFFFF
            if command == MENU_CONVERT_TABLE:
                self._convert_async(CONVERSION_MODE_TABLE)
                return 0
            if command == MENU_CONVERT_RICH_TEXT:
                self._convert_async(CONVERSION_MODE_RICH_TEXT)
                return 0
            if command == MENU_EXIT:
                user32.DestroyWindow(hwnd)
                return 0
            return int(user32.DefWindowProcW(hwnd, message, wparam, lparam))
        # config.iniで指定したホットキーで変換を実行します。
        if message == WM_HOTKEY and wparam == HOTKEY_ID:
            self._convert_async()
            return 0
        if message == WM_DESTROY:
            user32.UnregisterHotKey(hwnd, HOTKEY_ID)
            self._remove_tray_icon()
            user32.PostQuitMessage(0)
            return 0
        return int(user32.DefWindowProcW(hwnd, message, wparam, lparam))


def main(argv: Sequence[str] | None = None) -> int:
    """CLI引数を解釈し、標準入力・1回変換・常駐起動の各モードを実行します。"""

    parser = argparse.ArgumentParser(description="Convert Excel TSV clipboard text to Markdown.")
    parser.add_argument("--stdin", action="store_true", help="read TSV from stdin and write Markdown table to stdout")
    parser.add_argument("--once", action="store_true", help="convert clipboard once and exit")
    parser.add_argument(
        "--mode",
        choices=sorted(SUPPORTED_CONVERSION_MODES),
        default=None,
        help="conversion mode for --once (default: config.ini conversion.default_mode)",
    )
    args = parser.parse_args(argv)

    if args.stdin:
        sys.stdout.write(convert_text_to_markdown(sys.stdin.read()))
        return 0
    if not is_windows():
        print("ExceltoMarkdownはWindows専用です。--stdinのみ非Windowsでも利用できます。", file=sys.stderr)
        return 1
    if args.once:
        convert_clipboard_to_markdown(
            mode=args.mode or load_default_conversion_mode_config(),
            prefer_excel=load_prefer_excel_config(),
        )
        return 0
    TrayApplication().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
