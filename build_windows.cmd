@echo off
setlocal

rem GitHub Actions sets this from pyproject.toml. Keep the default convenient
rem for a local Windows x64 build of the v1.1.0 release.
if not defined EXCELTOMARKDOWN_VERSION set "EXCELTOMARKDOWN_VERSION=1.1.0.0"

python -m nuitka ^
    --mingw64 ^
    --standalone ^
    --remove-output ^
    --assume-yes-for-downloads ^
    --include-data-file=e2m_ico.ico=e2m_ico.ico ^
    --include-data-file=config.ini=config.ini ^
    --windows-icon-from-ico=e2m_ico.ico ^
    --windows-company-name="Akihiko Fujita" ^
    --windows-product-name="Excel to Markdown Converter" ^
    --windows-file-version="%EXCELTOMARKDOWN_VERSION%" ^
    --windows-product-version="%EXCELTOMARKDOWN_VERSION%" ^
    --windows-file-description="Excel to Markdown Converter" ^
    --windows-console-mode=disable ^
    --python-flag=no_warnings ^
    --python-flag=no_docstrings ^
    --nofollow-import-to=PIL ^
    --include-module=win32com.client ^
    --include-module=pythoncom ^
    --include-module=pywintypes ^
    --output-filename=ExcelToMarkdown.exe ^
    excel_to_markdown.py

if errorlevel 1 exit /b %errorlevel%

echo Build completed: excel_to_markdown.dist\ExcelToMarkdown.exe
