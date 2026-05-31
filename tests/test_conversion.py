import io
import unittest
from unittest.mock import patch

from excel_to_markdown import Cell, convert_text_to_markdown, icon_path, main, rows_to_markdown


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
        """タスクトレイ用アイコンとしてE2M.icoを参照することを確認します。"""

        self.assertEqual(icon_path().name, "E2M.ico")
        self.assertTrue(icon_path().exists())

    def test_main_rejects_non_windows(self):
        """Linux/UnixなどWindows以外ではアプリ本体が動作しないことを確認します。"""

        with patch("sys.platform", "linux"), patch("sys.stderr", new_callable=io.StringIO) as stderr:
            self.assertEqual(main(["--stdin"]), 1)
            self.assertIn("Windows専用", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()