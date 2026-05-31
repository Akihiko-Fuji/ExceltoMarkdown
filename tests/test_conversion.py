import io
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from excel_to_markdown import (
    MOD_ALT,
    MOD_CONTROL,
    Cell,
    config_path,
    convert_clipboard_or_excel_selection,
    convert_text_to_markdown,
    escape_markdown_cell,
    icon_path,
    load_hotkey_config,
    main,
    parse_hotkey,
    read_clipboard_text,
    rows_to_markdown,
    TrayApplication,
    WM_COMMAND,
)


class ConversionTests(unittest.TestCase):
    """Excel由来データをMarkdownへ変換する中核処理のテストです。"""

    def test_excel_tsv_to_markdown(self):
        """Excelコピー相当のTSVがMarkdown表へ変換されることを確認します。"""
        source = "A1\tB1\tC1\nA2\tB2\tC2\nA3\tB3\tC3\n"
        expected = "| A1 | B1 | C1 |\n| --- | --- | --- |\n| A2 | B2 | C2 |\n| A3 | B3 | C3 |\n"
        self.assertEqual(convert_text_to_markdown(source), expected)

    def test_escapes_markdown_table_delimiters(self):
        """Markdown表を壊しやすい文字がセル内でエスケープされることを確認します。"""

        source = "Header|1\tBack\\slash\nA\tB"
        expected = "| Header\\|1 | Back\\\\slash |\n| --- | --- |\n| A | B |\n"
        self.assertEqual(convert_text_to_markdown(source), expected)

    def test_optional_formatting(self):
        """太字・イタリック・リンクの書式情報がMarkdownへ反映されることを確認します。"""

        rows = [[Cell("Title", bold=True), Cell("Site", italic=True, href="https://example.com/a)b")], ["x", "y"]]
        expected = "| **Title** | *[Site](https://example.com/a%29b)* |\n| --- | --- |\n| x | y |\n"
        self.assertEqual(rows_to_markdown(rows), expected)

    def test_icon_path_uses_e2m_ico(self):
        """タスクトレイ用アイコンとしてe2m_ico.icoを参照することを確認します。"""

        self.assertEqual(icon_path().name, "e2m_ico.ico")
        self.assertTrue(icon_path().exists())

    def test_default_config_file_defines_hotkey(self):
        """config.iniから変換ショートカットを読み込むことを確認します。"""

        hotkey = load_hotkey_config(config_path())
        self.assertEqual(hotkey.label, "Ctrl+Alt+M")
        self.assertEqual(hotkey.modifiers, MOD_CONTROL | MOD_ALT)
        self.assertEqual(hotkey.virtual_key, ord("M"))

    def test_parse_hotkey_supports_function_keys(self):
        """修飾キー付きのファンクションキーも設定できることを確認します。"""

        hotkey = parse_hotkey("Ctrl+Alt+F12")
        self.assertEqual(hotkey.label, "Ctrl+Alt+F12")
        self.assertEqual(hotkey.modifiers, MOD_CONTROL | MOD_ALT)
        self.assertEqual(hotkey.virtual_key, 0x7B)

    def test_bold_italic_formatting_uses_triple_marker(self):
        """太字とイタリックを併用するセルはGFM互換の一括マーカーで整形します。"""

        rows = [[Cell("Title", bold=True, italic=True)], ["x"]]
        expected = "| ***Title*** |\n| --- |\n| x |\n"
        self.assertEqual(rows_to_markdown(rows), expected)

    def test_escape_markdown_cell_escapes_parentheses(self):
        """リンク構文への誤解釈を避けるため丸括弧もエスケープします。"""

        self.assertEqual(escape_markdown_cell("[label](value)"), "\\[label\\]\\(value\\)")

    def test_clipboard_fallback_does_not_overwrite_empty_markdown(self):
        """変換結果が空なら既存クリップボードを空文字列で上書きしません。"""

        with (
            patch("excel_to_markdown.pywin32_available", return_value=False),
            patch("excel_to_markdown.read_clipboard_text", return_value=""),
            patch("excel_to_markdown.write_clipboard_text") as write_clipboard_text,
        ):
            self.assertEqual(convert_clipboard_or_excel_selection(prefer_excel=True), "")
        write_clipboard_text.assert_not_called()

    def test_read_clipboard_text_raises_when_global_lock_fails(self):
        """クリップボードメモリのロック失敗は空文字列ではなく例外として扱います。"""

        fake_user32 = SimpleNamespace(
            OpenClipboard=Mock(return_value=True),
            GetClipboardData=Mock(return_value=123),
            CloseClipboard=Mock(),
        )
        fake_kernel32 = SimpleNamespace(
            GlobalLock=Mock(return_value=0),
            GlobalUnlock=Mock(),
        )
        with (
            patch("sys.platform", "win32"),
            patch("excel_to_markdown.user32", fake_user32),
            patch("excel_to_markdown.kernel32", fake_kernel32),
        ):
            with self.assertRaises(MemoryError):
                read_clipboard_text()
        fake_user32.CloseClipboard.assert_called_once()
        fake_kernel32.GlobalUnlock.assert_not_called()

    def test_show_menu_returns_when_popup_menu_creation_fails(self):
        """ポップアップメニュー作成に失敗したら後続Win32 APIを呼びません。"""

        fake_user32 = SimpleNamespace(
            CreatePopupMenu=Mock(return_value=0),
            AppendMenuW=Mock(),
        )
        app = object.__new__(TrayApplication)
        with patch("excel_to_markdown.user32", fake_user32):
            app._show_menu()
        fake_user32.AppendMenuW.assert_not_called()

    def test_add_tray_icon_only_adds_once(self):
        """タスクトレイ追加時は同一データで不要な更新呼び出しを行いません。"""

        fake_shell32 = SimpleNamespace(Shell_NotifyIconW=Mock(return_value=True))
        app = object.__new__(TrayApplication)
        app.hwnd = 1
        app._icon = 2
        app.hotkey = SimpleNamespace(label="Ctrl+Alt+M")
        with patch("excel_to_markdown.shell32", fake_shell32):
            app._add_tray_icon()
        fake_shell32.Shell_NotifyIconW.assert_called_once()

    def test_unknown_wm_command_delegates_to_default_window_proc(self):
        """未処理のWM_COMMANDは明示的にDefWindowProcWへ委譲します。"""

        fake_user32 = SimpleNamespace(DefWindowProcW=Mock(return_value=77))
        app = object.__new__(TrayApplication)
        with patch("excel_to_markdown.user32", fake_user32):
            self.assertEqual(app._window_proc(1, WM_COMMAND, 9999, 0), 77)
        fake_user32.DefWindowProcW.assert_called_once_with(1, WM_COMMAND, 9999, 0)

    def test_tsv_to_rows_ignores_multiple_trailing_blank_lines(self):
        """末尾に連続する改行は余分な空行としてMarkdownへ出力しません。"""

        source = "A\tB\n1\t2\n\n"
        expected = "| A | B |\n| --- | --- |\n| 1 | 2 |\n"
        self.assertEqual(convert_text_to_markdown(source), expected)

    def test_main_rejects_non_windows(self):
        """Linux/UnixなどWindows以外ではアプリ本体が動作しないことを確認します。"""

        with patch("sys.platform", "linux"), patch("sys.stderr", new_callable=io.StringIO) as stderr:
            self.assertEqual(main(["--stdin"]), 1)
            self.assertIn("Windows専用", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()