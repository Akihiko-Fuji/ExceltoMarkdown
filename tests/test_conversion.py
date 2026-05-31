import configparser
import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import excel_to_markdown as e2m
from excel_to_markdown import (
    MOD_ALT,
    MOD_CONTROL,
    MOD_SHIFT,
    Cell,
    config_path,
    convert_clipboard_or_excel_selection,
    convert_text_to_markdown,
    escape_markdown_cell,
    icon_path,
    load_hotkey_config,
    load_prefer_excel_config,
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
        expected = "| **Title** | [*Site*](https://example.com/a%29b) |\n| --- | --- |\n| x | y |\n"
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

    def test_default_shortcut_does_not_override_inherited_section_key(self):
        """DEFAULT由来のkeyをshortcutセクション直下のkeyとして誤認しないことを確認します。"""

        with tempfile.TemporaryDirectory() as directory:
            config_file = Path(directory) / "config.ini"
            config_file.write_text(
                "[DEFAULT]\nkey = Ctrl+Alt+X\nshortcut = Ctrl+Shift+Z\n[shortcut]\nname = ignored\n",
                encoding="utf-8",
            )
            hotkey = load_hotkey_config(config_file)
        self.assertEqual(hotkey.label, "Ctrl+Shift+Z")
        self.assertEqual(hotkey.modifiers, MOD_CONTROL | MOD_SHIFT)
        self.assertEqual(hotkey.virtual_key, ord("Z"))

    def test_explicit_shortcut_key_overrides_default_key_with_same_name(self):
        """DEFAULTにもkeyがある場合でもshortcutセクション直下のkeyを優先します。"""

        with tempfile.TemporaryDirectory() as directory:
            config_file = Path(directory) / "config.ini"
            config_file.write_text(
                "[DEFAULT]\nkey = Ctrl+Alt+X\n[shortcut]\nkey = Ctrl+Shift+Y\n",
                encoding="utf-8",
            )
            hotkey = load_hotkey_config(config_file)
        self.assertEqual(hotkey.label, "Ctrl+Shift+Y")
        self.assertEqual(hotkey.modifiers, MOD_CONTROL | MOD_SHIFT)
        self.assertEqual(hotkey.virtual_key, ord("Y"))

    def test_get_explicit_config_option_does_not_mutate_defaults(self):
        """明示設定の確認でConfigParserのDEFAULT辞書を破壊しないことを確認します。"""

        parser = configparser.ConfigParser()
        parser.read_string("[DEFAULT]\nkey = Ctrl+Alt+X\n[shortcut]\nname = ignored\n")
        defaults = parser.defaults()
        self.assertIsNone(e2m._get_explicit_config_option(parser, "shortcut", "key"))
        self.assertEqual(defaults, {"key": "Ctrl+Alt+X"})
        self.assertIs(parser.defaults(), defaults)

    def test_explicit_shortcut_key_interpolates_default_value(self):
        """セクション直下のkeyがDEFAULT値を補間している場合も正しく読み込みます。"""

        with tempfile.TemporaryDirectory() as directory:
            config_file = Path(directory) / "config.ini"
            config_file.write_text(
                "[DEFAULT]\nmod = Ctrl+Alt\n[shortcut]\nkey = %(mod)s+M\n",
                encoding="utf-8",
            )
            hotkey = load_hotkey_config(config_file)
        self.assertEqual(hotkey.label, "Ctrl+Alt+M")
        self.assertEqual(hotkey.modifiers, MOD_CONTROL | MOD_ALT)
        self.assertEqual(hotkey.virtual_key, ord("M"))

    def test_bold_italic_formatting_uses_triple_marker(self):
        """太字とイタリックを併用するセルはGFM互換の一括マーカーで整形します。"""

        rows = [[Cell("Title", bold=True, italic=True)], ["x"]]
        expected = "| ***Title*** |\n| --- |\n| x |\n"
        self.assertEqual(rows_to_markdown(rows), expected)

    def test_link_text_keeps_parentheses_readable(self):
        """リンクテキスト内の丸括弧はバックスラッシュ表示にならないことを確認します。"""

        rows = [[Cell("Report (final)", href="https://example.com/a)b")], ["x"]]
        expected = "| [Report (final)](https://example.com/a%29b) |\n| --- |\n| x |\n"
        self.assertEqual(rows_to_markdown(rows), expected)

    def test_bold_italic_link_formats_link_text(self):
        """リンク付き太字イタリックはリンクの表示テキスト側へマーカーを入れます。"""

        rows = [[Cell("Docs", bold=True, italic=True, href="https://example.com/docs")], ["x"]]
        expected = "| [***Docs***](https://example.com/docs) |\n| --- |\n| x |\n"
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

    def test_excel_selection_too_large_does_not_fallback_to_clipboard(self):
        """Excel選択範囲が大きすぎる場合は古いクリップボード内容へフォールバックしません。"""

        with (
            patch("excel_to_markdown.pywin32_available", return_value=True),
            patch("excel_to_markdown._excel_selection_to_rows", side_effect=ValueError("too large")),
            patch("excel_to_markdown.read_clipboard_text") as read_clipboard_text,
        ):
            with self.assertRaises(ValueError):
                convert_clipboard_or_excel_selection(prefer_excel=True)
        read_clipboard_text.assert_not_called()

    def test_prefer_excel_logs_com_failure_before_clipboard_fallback(self):
        """Excel COM取得失敗時は理由をログへ残してからクリップボード変換へ戻ります。"""

        with tempfile.TemporaryDirectory() as directory:
            with (
                patch("excel_to_markdown.pywin32_available", return_value=True),
                patch("excel_to_markdown._excel_selection_to_rows", side_effect=RuntimeError("COM unavailable")),
                patch("excel_to_markdown.app_base_dir", return_value=Path(directory)),
                patch("excel_to_markdown.read_clipboard_text", return_value="A\tB\n1\t2\n"),
                patch("excel_to_markdown.write_clipboard_text"),
            ):
                markdown = convert_clipboard_or_excel_selection(prefer_excel=True)
            log_text = (Path(directory) / "excel_to_markdown_error.log").read_text(encoding="utf-8")
        self.assertIn("COM unavailable", log_text)
        self.assertIn("| A | B |", markdown)

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

    def test_run_registers_hotkey_before_adding_tray_icon(self):
        """ホットキー登録に失敗した場合に死んだトレイアイコンを残さない順序にします。"""

        calls = []
        fake_user32 = SimpleNamespace(
            RegisterHotKey=Mock(side_effect=lambda *args: calls.append("register_hotkey") or False),
        )
        app = object.__new__(TrayApplication)
        app.hotkey = SimpleNamespace(modifiers=MOD_CONTROL | MOD_ALT, virtual_key=ord("M"))
        app.hwnd = 1
        app._register_window_class = Mock(side_effect=lambda: calls.append("register_class"))
        app._create_window = Mock(side_effect=lambda: calls.append("create_window"))
        app._add_tray_icon = Mock(side_effect=lambda: calls.append("add_tray_icon"))

        with patch("excel_to_markdown.user32", fake_user32), self.assertRaises(Exception):
            app.run()

        self.assertEqual(calls, ["register_class", "create_window", "register_hotkey"])
        app._add_tray_icon.assert_not_called()

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


    def test_convert_safely_skips_when_already_running(self):
        """連打時は変換処理を多重起動せず、短い通知音だけで戻ることを確認します。"""

        fake_user32 = SimpleNamespace(MessageBeep=Mock())
        app = object.__new__(TrayApplication)
        app._convert_lock = e2m.threading.Lock()
        app._convert_lock.acquire()
        try:
            with (
                patch("excel_to_markdown.user32", fake_user32),
                patch("excel_to_markdown.convert_clipboard_or_excel_selection") as convert,
            ):
                app._convert_safely()
        finally:
            app._convert_lock.release()
        convert.assert_not_called()
        fake_user32.MessageBeep.assert_called_once_with(0xFFFFFFFF)

    def test_convert_safely_does_not_beep_for_empty_markdown(self):
        """変換結果が空の場合は成功音ではなく案内ダイアログを出すことを確認します。"""

        fake_user32 = SimpleNamespace(MessageBeep=Mock(), MessageBoxW=Mock())
        app = object.__new__(TrayApplication)
        app._convert_lock = e2m.threading.Lock()
        app.prefer_excel = False
        app.hwnd = 100
        with (
            patch("excel_to_markdown.user32", fake_user32),
            patch("excel_to_markdown.convert_clipboard_or_excel_selection", return_value=""),
        ):
            app._convert_safely()
        fake_user32.MessageBeep.assert_not_called()
        fake_user32.MessageBoxW.assert_called_once()

    def test_configure_windows_api_declares_window_and_tray_functions(self):
        """ウィンドウ管理・トレイ関連Win32 APIにもctypes型宣言を付けることを確認します。"""

        user32_names = [
            "OpenClipboard",
            "CloseClipboard",
            "EmptyClipboard",
            "GetClipboardData",
            "SetClipboardData",
            "LoadImageW",
            "RegisterClassW",
            "CreateWindowExW",
            "DefWindowProcW",
            "DestroyWindow",
            "RegisterHotKey",
            "UnregisterHotKey",
            "GetMessageW",
            "TranslateMessage",
            "DispatchMessageW",
            "PostQuitMessage",
            "CreatePopupMenu",
            "AppendMenuW",
            "TrackPopupMenu",
            "DestroyMenu",
            "GetCursorPos",
            "SetForegroundWindow",
            "MessageBeep",
            "MessageBoxW",
        ]
        kernel32_names = [
            "GetModuleHandleW",
            "GlobalAlloc",
            "GlobalLock",
            "GlobalUnlock",
            "GlobalFree",
        ]
        fake_user32 = SimpleNamespace(**{name: Mock() for name in user32_names})
        fake_kernel32 = SimpleNamespace(**{name: Mock() for name in kernel32_names})
        fake_shell32 = SimpleNamespace(Shell_NotifyIconW=Mock())

        with (
            patch("sys.platform", "win32"),
            patch("excel_to_markdown.user32", fake_user32),
            patch("excel_to_markdown.kernel32", fake_kernel32),
            patch("excel_to_markdown.shell32", fake_shell32),
        ):
            e2m._configure_windows_api()

        for api in [*user32_names, *kernel32_names]:
            function = getattr(fake_user32, api, None) or getattr(fake_kernel32, api, None)
            self.assertIn("argtypes", vars(function), api)
            self.assertIn("restype", vars(function), api)
        self.assertIn("argtypes", vars(fake_shell32.Shell_NotifyIconW))
        self.assertIn("restype", vars(fake_shell32.Shell_NotifyIconW))

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

    def test_main_stdin_runs_on_non_windows(self):
        """Linux/UnixなどWindows以外でも標準入力変換はOS判定で終了しないことを確認します。"""

        with (
            patch("sys.platform", "linux"),
            patch("sys.stdin", io.StringIO("A\tB\n1\t2\n")),
            patch("sys.stdout", new_callable=io.StringIO) as stdout,
        ):
            self.assertEqual(main(["--stdin"]), 0)
        self.assertIn("| A | B |", stdout.getvalue())

    def test_main_once_on_non_windows_prints_friendly_error(self):
        """非Windowsの--onceはWin32 API例外ではなく明示メッセージで終了します。"""

        with (
            patch("sys.platform", "linux"),
            patch("sys.stderr", new_callable=io.StringIO) as stderr,
            patch("excel_to_markdown.convert_clipboard_or_excel_selection") as convert,
        ):
            self.assertEqual(main(["--once"]), 1)
        convert.assert_not_called()
        self.assertIn("--stdin", stderr.getvalue())

    def test_prefer_excel_config_defaults_to_clipboard(self):
        """Excel起動中の別選択範囲を誤変換しないよう、既定値はクリップボード優先にします。"""

        with tempfile.TemporaryDirectory() as directory:
            config_file = Path(directory) / "config.ini"
            config_file.write_text("[shortcut]\nkey = Ctrl+Alt+M\n", encoding="utf-8")
            self.assertFalse(load_prefer_excel_config(config_file))

    def test_prefer_excel_config_can_enable_excel_selection(self):
        """config.iniで明示した場合だけExcel選択範囲を優先できることを確認します。"""

        with tempfile.TemporaryDirectory() as directory:
            config_file = Path(directory) / "config.ini"
            config_file.write_text("[conversion]\nprefer_excel = true\n", encoding="utf-8")
            self.assertTrue(load_prefer_excel_config(config_file))


if __name__ == "__main__":
    unittest.main()