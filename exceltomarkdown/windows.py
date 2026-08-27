"""Windows clipboard、Excel COM、タスクトレイの実装です。

標準入力のTSV変換はOSを問わず利用できます。Windowsではクリップボード、
ホットキー、タスクトレイを使った常駐アプリとしても動作します。任意の
``pywin32`` がインストールされている場合のみ、COM経由でExcelのアクティブ
な選択範囲を読み取り、太字・イタリック・ハイパーリンクなどの簡単な書式
もMarkdownへ反映します。
"""

from __future__ import annotations

import ctypes
import importlib
import importlib.util
import re
import sys
import threading
import time
import traceback
from ctypes import wintypes
from pathlib import Path

from .config import (
    APP_ICON_FILENAME,
    CONVERSION_MODE_AUTO,
    CONVERSION_MODE_RICH_TEXT,
    CONVERSION_MODE_TABLE,
    DEFAULT_CONVERSION_MODE,
    DEFAULT_PREFER_EXCEL,
    INTERNAL_CONVERSION_MODES,
    MOD_ALT,
    MOD_CONTROL,
    MOD_SHIFT,
    MOD_WIN,
    app_base_dir,
    icon_path,
    load_default_conversion_mode_config,
    load_hotkey_config,
    load_prefer_excel_config,
    load_ui_language_config,
    tr,
)
from .core import Cell, convert_table_to_markdown, rows_to_markdown
from .rich_text import convert_html_fragment_to_markdown

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
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
HTML_CLIPBOARD_FORMAT_NAME = "HTML Format"

MENU_CONVERT_TABLE = 1001
MENU_CONVERT_RICH_TEXT = 1002
MENU_EXIT = 1003
# 後方互換のため、既存テストや外部利用のMENU_CONVERT名は表変換コマンドとして残します。
MENU_CONVERT = MENU_CONVERT_TABLE
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
    if hasattr(kernel32, "GetUserDefaultUILanguage"):
        kernel32.GetUserDefaultUILanguage.argtypes = []
        kernel32.GetUserDefaultUILanguage.restype = wintypes.LANGID
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
    if hasattr(ctypes, "set_last_error"):
        ctypes.set_last_error(0)
    size = int(kernel32.GlobalSize(handle))
    if size == 0:
        get_last_error = getattr(ctypes, "get_last_error", None)
        last_error = get_last_error() if get_last_error else 0
        if last_error:
            raise ctypes.WinError(last_error)
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
        print(tr("log_rich_text_html_fallback"), file=log_file)
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
        importlib.util.find_spec("pythoncom") is not None
        and importlib.util.find_spec("win32com") is not None
        and importlib.util.find_spec("win32com.client") is not None
    )


def _excel_selection_to_rows() -> list[list[Cell]]:
    """pywin32経由でExcelの選択範囲を読み取り、簡易書式もCellへ格納します。"""

    pythoncom = importlib.import_module("pythoncom")
    win32com_client = importlib.import_module("win32com.client")
    initialized = False
    try:
        # TrayApplicationは変換をworker threadで実行するため、そのthreadごとに
        # COM apartmentを初期化する必要があります。
        pythoncom.CoInitialize()
        initialized = True
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
    finally:
        if initialized:
            pythoncom.CoUninitialize()


def write_error_log(callback) -> Path | None:
    """エラーログを書き込み、書き込み先に権限がない場合は静かに諦めます。"""

    log_directory = app_base_dir()
    log_path = log_directory / "excel_to_markdown_error.log"
    try:
        log_directory.mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as log_file:
            callback(log_file)
    except Exception:
        return None
    return log_path


def _log_excel_selection_fallback(error: BaseException) -> None:
    """Excel選択範囲の取得に失敗してクリップボード変換へ戻る理由をログへ残します。"""

    def write_fallback_reason(log_file) -> None:
        print(tr("log_excel_selection_fallback"), file=log_file)
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
        self.language = load_ui_language_config()
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
        nid.szTip = tr("tray_tooltip", getattr(self, "language", "en"), hotkey=self.hotkey.label)
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
        language = getattr(self, "language", "en")
        user32.AppendMenuW(menu, 0, MENU_CONVERT_TABLE, tr("tray_convert_table", language))
        user32.AppendMenuW(menu, 0, MENU_CONVERT_RICH_TEXT, tr("tray_convert_rich_text", language))
        user32.AppendMenuW(menu, 0, MENU_EXIT, tr("tray_exit", language))
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
                mode or self.default_mode,
                prefer_excel=self.prefer_excel,
            )
            # convert_clipboard_to_markdown() は成功時に各変換経路内でクリップボードへ書き込みます。
            # ここでは結果の有無に応じて通知だけ行い、書き込み処理を重複させません。
            if markdown:
                user32.MessageBeep(0xFFFFFFFF)
            else:
                user32.MessageBoxW(
                    self.hwnd,
                    tr("no_text_to_convert", getattr(self, "language", "en")),
                    "Excel to Markdown",
                    0x40,
                )
        except Exception:
            log_path = write_error_log(lambda log_file: traceback.print_exc(file=log_file))
            language = getattr(self, "language", "en")
            if log_path is None:
                message = f"{tr('conversion_failed', language)}\n{tr('log_write_failed', language)}"
            else:
                message = f"{tr('conversion_failed', language)}\n{log_path}"
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
