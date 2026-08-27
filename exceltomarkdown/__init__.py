"""Excel/Word/Webのコピー内容をGFMへ渡す変換ブリッジです。"""

from .config import (
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
    SUPPORTED_CONVERSION_MODES,
    HotkeyConfig,
    app_base_dir,
    config_path,
    detect_ui_language,
    get_ui_language,
    icon_path,
    load_default_conversion_mode_config,
    load_hotkey_config,
    load_prefer_excel_config,
    load_ui_language_config,
    normalize_conversion_mode,
    normalize_ui_language,
    parse_hotkey,
    resource_path,
    tr,
)
from .core import (
    Cell,
    convert_table_to_markdown,
    convert_text_to_markdown,
    escape_link_destination,
    escape_markdown_cell,
    format_cell,
    normalize_rows,
    rows_to_markdown,
    tsv_to_rows,
)
from .rich_text import (
    convert_html_fragment_to_markdown,
    convert_html_to_markdown,
    convert_rich_text_to_markdown,
    escape_markdown_text,
)
from .windows import (
    MENU_CONVERT,
    MENU_CONVERT_RICH_TEXT,
    MENU_CONVERT_TABLE,
    TrayApplication,
    WM_COMMAND,
    clipboard_has_format,
    convert_clipboard_or_excel_selection,
    convert_clipboard_rich_text_to_markdown,
    convert_clipboard_to_markdown,
    extract_clipboard_html_fragment,
    is_windows,
    open_clipboard_with_retry,
    pywin32_available,
    read_clipboard_html_fragment,
    read_clipboard_plain_text,
    read_clipboard_text,
    register_clipboard_format,
    write_clipboard_text,
)

__version__ = "0.2.0"

__all__ = [name for name in globals() if not name.startswith("_")]
