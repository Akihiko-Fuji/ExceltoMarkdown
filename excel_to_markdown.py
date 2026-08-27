"""既存の起動方法とimportを維持する互換entry pointです。

新しいコードでは ``exceltomarkdown`` packageからのimportを推奨します。
"""

from exceltomarkdown import *  # noqa: F401,F403
from exceltomarkdown.cli import main

# 旧単一moduleのprivate helperを参照していたテスト・利用コードへの移行猶予です。
from exceltomarkdown import config as _config
from exceltomarkdown import rich_text as _rich_text
from exceltomarkdown import windows as _windows

_RichHtmlToMarkdownParser = _rich_text._RichHtmlToMarkdownParser
_configure_windows_api = _windows._configure_windows_api
_excel_selection_to_rows = _windows._excel_selection_to_rows
_get_explicit_config_option = _config._get_explicit_config_option
_get_section_own_options = _config._get_section_own_options
_read_clipboard_format_bytes = _windows._read_clipboard_format_bytes


if __name__ == "__main__":
    raise SystemExit(main())
