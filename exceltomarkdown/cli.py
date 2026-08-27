"""ExceltoMarkdownのcommand line entry pointです。"""

from __future__ import annotations

import argparse
import sys
from typing import Sequence

from .config import (
    SUPPORTED_CONVERSION_MODES,
    load_default_conversion_mode_config,
    load_prefer_excel_config,
    load_ui_language_config,
    tr,
)
from .core import convert_text_to_markdown
from .windows import TrayApplication, convert_clipboard_to_markdown, is_windows


def main(argv: Sequence[str] | None = None) -> int:
    """標準入力・1回変換・Windows常駐起動の各モードを実行します。"""

    language = load_ui_language_config()
    parser = argparse.ArgumentParser(description=tr("cli_description", language))
    parser.add_argument("--stdin", action="store_true", help=tr("cli_stdin_help", language))
    parser.add_argument("--once", action="store_true", help=tr("cli_once_help", language))
    parser.add_argument(
        "--mode",
        choices=sorted(SUPPORTED_CONVERSION_MODES),
        default=None,
        help=tr("cli_mode_help", language),
    )
    args = parser.parse_args(argv)

    if args.stdin:
        sys.stdout.write(convert_text_to_markdown(sys.stdin.read()))
        return 0
    if not is_windows():
        print(tr("non_windows_unavailable", language), file=sys.stderr)
        return 1
    if args.once:
        convert_clipboard_to_markdown(
            mode=args.mode or load_default_conversion_mode_config(),
            prefer_excel=load_prefer_excel_config(),
        )
        return 0
    TrayApplication().run()
    return 0
