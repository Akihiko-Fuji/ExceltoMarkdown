"""設定ファイル、ホットキー、UI言語の処理です。"""

from __future__ import annotations

import configparser
import ctypes
import io
import locale
import os
import sys
from dataclasses import dataclass
from pathlib import Path


APP_ICON_FILENAME = "e2m_ico.ico"
CONFIG_FILENAME = "config.ini"
DEFAULT_HOTKEY = "Ctrl+Alt+M"
DEFAULT_PREFER_EXCEL = False
DEFAULT_CONVERSION_MODE = "table"
DEFAULT_UI_LANGUAGE = "auto"
UI_LANGUAGE_AUTO = "auto"
SUPPORTED_UI_LANGUAGES = {"ja", "en"}
JAPANESE_PRIMARY_LANGUAGE_ID = 0x11
CONVERSION_MODE_TABLE = "table"
CONVERSION_MODE_RICH_TEXT = "rich_text"
CONVERSION_MODE_AUTO = "auto"
SUPPORTED_CONVERSION_MODES = {CONVERSION_MODE_TABLE, CONVERSION_MODE_RICH_TEXT}
INTERNAL_CONVERSION_MODES = {*SUPPORTED_CONVERSION_MODES, CONVERSION_MODE_AUTO}

# RegisterHotKeyで使う修飾キーです。Windows以外でも設定テストを実行できます。
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008


MESSAGES = {
    "ja": {
        "tray_convert_table": "Markdown表に変換",
        "tray_convert_rich_text": "リッチテキストをMarkdown化",
        "tray_exit": "終了 (&X)",
        "tray_tooltip": "Excel to Markdown ({hotkey})",
        "no_text_to_convert": "変換対象のテキストがありません。",
        "conversion_failed": "変換に失敗しました。",
        "log_write_failed": "ログファイルへ書き込めませんでした。",
        "cli_description": "ExcelのTSVクリップボードテキストをGFMへ変換します。",
        "cli_stdin_help": "標準入力からTSVを読み取り、GFM tableを標準出力へ書き出します",
        "cli_once_help": "クリップボードを1回だけ変換して終了します",
        "cli_mode_help": "--onceで使う変換モード (既定: config.ini conversion.default_mode)",
        "non_windows_unavailable": "ExceltoMarkdownはWindows専用です。--stdinのみ非Windowsでも利用できます。",
        "log_excel_selection_fallback": "Excel選択範囲の取得に失敗したため、クリップボードTSVへフォールバックします。",
        "log_rich_text_html_fallback": "HTML Formatの取得または解析に失敗したため、プレーンテキストへフォールバックします。",
    },
    "en": {
        "tray_convert_table": "Convert table to Markdown",
        "tray_convert_rich_text": "Convert rich text to Markdown",
        "tray_exit": "Exit (&X)",
        "tray_tooltip": "Excel to Markdown ({hotkey})",
        "no_text_to_convert": "No text to convert.",
        "conversion_failed": "Conversion failed.",
        "log_write_failed": "Could not write to the log file.",
        "cli_description": "Convert Excel TSV clipboard text to GFM.",
        "cli_stdin_help": "read TSV from stdin and write a GFM table to stdout",
        "cli_once_help": "convert clipboard once and exit",
        "cli_mode_help": "conversion mode for --once (default: config.ini conversion.default_mode)",
        "non_windows_unavailable": "ExceltoMarkdown is Windows-only. Only --stdin is available on non-Windows platforms.",
        "log_excel_selection_fallback": "Failed to read the Excel selection; falling back to clipboard TSV.",
        "log_rich_text_html_fallback": "Failed to read or parse HTML Format; falling back to plain text.",
    },
}


if sys.platform == "win32":
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    if hasattr(kernel32, "GetUserDefaultUILanguage"):
        kernel32.GetUserDefaultUILanguage.argtypes = []
        kernel32.GetUserDefaultUILanguage.restype = ctypes.c_ushort
else:
    kernel32 = None


@dataclass(frozen=True)
class HotkeyConfig:
    """WindowsのRegisterHotKeyへ渡すショートカット設定です。"""

    label: str
    modifiers: int
    virtual_key: int


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
_FUNCTION_KEY_ALIASES = {f"f{number}": 0x70 + number - 1 for number in range(1, 25)}
_KEY_ALIASES.update(_FUNCTION_KEY_ALIASES)


def _package_dir() -> Path:
    return Path(__file__).resolve().parent


def app_base_dir() -> Path:
    """設定・ログを置く、現在の配布形態に合った基準ディレクトリを返します。"""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    working_directory = Path.cwd()
    if (working_directory / CONFIG_FILENAME).is_file():
        return working_directory
    if sys.platform == "win32":
        app_data = os.environ.get("APPDATA")
        base_directory = Path(app_data) if app_data else Path.home() / "AppData" / "Roaming"
        return base_directory / "ExceltoMarkdown"
    config_home = os.environ.get("XDG_CONFIG_HOME")
    configured_directory = Path(config_home) if config_home else None
    base_directory = (
        configured_directory
        if configured_directory is not None and configured_directory.is_absolute()
        else Path.home() / ".config"
    )
    return base_directory / "exceltomarkdown"


def resource_path(filename: str) -> Path:
    """source checkout、wheel、PyInstallerの同梱リソースを解決します。"""

    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    source_candidate = _package_dir().parent / filename
    if source_candidate.exists():
        return source_candidate
    return _package_dir() / filename


def icon_path() -> Path:
    """タスクトレイ用アイコンのパスを返します。"""

    return resource_path(APP_ICON_FILENAME)


def config_path() -> Path:
    """現在の配布形態で読み込むconfig.iniのパスを返します。"""

    return app_base_dir() / CONFIG_FILENAME


def _parse_virtual_key(key_name: str) -> tuple[str, int]:
    normalized = key_name.strip().lower()
    if len(normalized) == 1 and normalized.isascii() and normalized.isalnum():
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
        own_options.add(parser.optionxform(line[: min(delimiters)].strip()))
    return own_options


def _get_explicit_config_option(
    parser: configparser.ConfigParser,
    section: str,
    option: str,
) -> str | None:
    """DEFAULT継承値ではなく、セクション直下に書かれた値だけを返します。"""

    option_norm = parser.optionxform(option)
    if section == parser.default_section:
        return parser.defaults().get(option_norm)
    if not parser.has_section(section):
        return None
    if option_norm not in _get_section_own_options(parser, section):
        return None
    return parser.get(section, option)


def _read_config(path: Path | None) -> configparser.ConfigParser:
    parser = configparser.ConfigParser()
    config_file = path or config_path()
    if config_file.exists():
        parser.read(config_file, encoding="utf-8")
    return parser


def load_hotkey_config(path: Path | None = None) -> HotkeyConfig:
    parser = _read_config(path)
    value = DEFAULT_HOTKEY
    explicit_shortcut_key = _get_explicit_config_option(parser, "shortcut", "key")
    explicit_hotkey_key = _get_explicit_config_option(parser, "hotkey", "key")
    if explicit_shortcut_key is not None:
        value = explicit_shortcut_key
    elif explicit_hotkey_key is not None:
        value = explicit_hotkey_key
    else:
        explicit_default = _get_explicit_config_option(parser, parser.default_section, "shortcut")
        if explicit_default is not None:
            value = explicit_default
    return parse_hotkey(value)


def _parse_config_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "yes", "true", "on", "enabled"}:
        return True
    if normalized in {"0", "no", "false", "off", "disabled"}:
        return False
    raise ValueError(f"真偽値として解釈できない設定値です: {value}")


def load_prefer_excel_config(path: Path | None = None) -> bool:
    parser = _read_config(path)
    value = _get_explicit_config_option(parser, "conversion", "prefer_excel")
    if value is None:
        value = _get_explicit_config_option(parser, parser.default_section, "prefer_excel")
    return DEFAULT_PREFER_EXCEL if value is None else _parse_config_bool(value)


def normalize_conversion_mode(value: str) -> str:
    normalized = value.strip().lower().replace("-", "_")
    if normalized not in SUPPORTED_CONVERSION_MODES:
        raise ValueError(f"未対応の変換モードです: {value}")
    return normalized


def load_default_conversion_mode_config(path: Path | None = None) -> str:
    parser = _read_config(path)
    value = _get_explicit_config_option(parser, "conversion", "default_mode")
    if value is None:
        value = _get_explicit_config_option(parser, parser.default_section, "default_mode")
    return DEFAULT_CONVERSION_MODE if value is None else normalize_conversion_mode(value)


def _language_from_locale_name(locale_name: str | None) -> str | None:
    if not locale_name:
        return None
    return "ja" if locale_name.strip().lower().startswith("ja") else None


def detect_ui_language() -> str:
    """OSの表示言語またはロケールからUI言語をja/enで判定します。"""

    try:
        if sys.platform == "win32" and kernel32 is not None and hasattr(kernel32, "GetUserDefaultUILanguage"):
            langid = int(kernel32.GetUserDefaultUILanguage())
            primary_language_id = langid & 0x3FF
            if primary_language_id:
                return "ja" if primary_language_id == JAPANESE_PRIMARY_LANGUAGE_ID else "en"
    except Exception:
        pass

    try:
        candidates: list[str | None] = []
        try:
            candidates.append(locale.getlocale()[0])
        except Exception:
            pass
        if hasattr(locale, "LC_MESSAGES"):
            try:
                candidates.append(locale.getlocale(locale.LC_MESSAGES)[0])
            except Exception:
                pass
        try:
            candidates.append(locale.getdefaultlocale()[0])
        except Exception:
            pass
        candidates.extend(os.environ.get(name) for name in ("LC_ALL", "LC_MESSAGES", "LANG", "LANGUAGE"))
        return "ja" if any(_language_from_locale_name(candidate) == "ja" for candidate in candidates) else "en"
    except Exception:
        return "en"


def normalize_ui_language(value: str) -> str:
    normalized = value.strip().lower().replace("-", "_")
    if normalized in {UI_LANGUAGE_AUTO, *SUPPORTED_UI_LANGUAGES}:
        return normalized
    raise ValueError(f"未対応のUI言語です: {value}")


def load_ui_language_config(path: Path | None = None) -> str:
    parser = _read_config(path)
    value = _get_explicit_config_option(parser, "ui", "language")
    if value is None:
        value = _get_explicit_config_option(parser, parser.default_section, "language")
    normalized = normalize_ui_language(value or DEFAULT_UI_LANGUAGE)
    return detect_ui_language() if normalized == UI_LANGUAGE_AUTO else normalized


def get_ui_language() -> str:
    return load_ui_language_config()


def tr(key: str, language: str | None = None, **kwargs) -> str:
    selected_language = language or get_ui_language()
    template = MESSAGES.get(selected_language, {}).get(key)
    if template is None:
        template = MESSAGES["en"].get(key, key)
    return template.format(**kwargs) if kwargs else template
