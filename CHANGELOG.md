# Changelog

All notable changes to this project will be documented in this file.

## Unreleased

## 1.1.0 - 2026-08-27

### Added

- Excel quoted TSV support for embedded newlines, tabs, and escaped quotes.
- Windows CI and tagged GitHub Release builds.
- Nuitka standalone Windows x64 ZIP and SHA-256 checksum release assets.
- Installable Python package and `exceltomarkdown` console command.

### Changed

- Split conversion, rich text, configuration, Windows, and CLI code into modules.
- Clarified the lossy GFM conversion contract and first-row header semantics.
- Hardened HTML block-syntax and link-destination escaping.

### Fixed

- Initialize and release COM in the worker thread that reads the Excel selection.
- Removed unused `[rich_text]` configuration options.
