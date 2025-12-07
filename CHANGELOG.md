# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `column_widths` parameter - set custom column widths by index
  - Dict mapping column index (0-based) to width in characters
  - Example: `column_widths={0: 25, 1: 15, 3: 30}`
- `row_heights` parameter - set custom row heights by index
  - Dict mapping row index (0-based) to height in points
  - Example: `row_heights={0: 22, 5: 30}`
- Both parameters available in `df_to_xlsx()` and `dfs_to_xlsx()`

## [0.3.0] - 2025-12-05

### Added
- `autofit` parameter - automatically adjust column widths to fit content
- `table_style` parameter - apply Excel table formatting with 61 built-in styles
  - Light styles: Light1-Light21
  - Medium styles: Medium1-Medium28
  - Dark styles: Dark1-Dark11
  - Tables include autofilter dropdowns and banded rows
- `freeze_panes` parameter - freeze header row for easier scrolling
- All new parameters available in both `df_to_xlsx()` and `dfs_to_xlsx()`

### Changed
- Updated type stubs with new parameters and documentation

## [0.2.0] - 2025-12-05

### Added
- `df_to_xlsx()` function for direct DataFrame export (pandas and polars)
- `dfs_to_xlsx()` function for writing multiple DataFrames to separate sheets
- `parallel=True` option for `csv_to_xlsx()` using multi-core processing
- Type preservation for DataFrame columns:
  - Python int/float → Excel numbers
  - Python bool → Excel booleans
  - datetime.date → Excel dates with formatting
  - datetime.datetime / pandas.Timestamp → Excel datetimes with formatting
  - None/NaN/NaT → Empty cells
- Type stubs for better IDE support
- rayon dependency for parallel processing

### Changed
- Updated documentation to include DataFrame and parallel processing examples

## [0.1.0] - 2025-12-04

### Added
- Initial release
- Python bindings via PyO3
- `csv_to_xlsx()` function for converting CSV files to Excel format
- Automatic type detection from CSV strings:
  - Integers and floats → Excel numbers
  - Booleans (`true`/`false`, case-insensitive) → Excel booleans
  - Dates (YYYY-MM-DD, DD/MM/YYYY, etc.) → Excel dates with formatting
  - Datetimes (ISO 8601) → Excel datetimes with formatting
  - NaN/Inf → Empty cells
  - Empty strings → Empty cells
- CLI tool for command-line usage
- Support for custom sheet names
- Verbose mode for progress reporting

[0.2.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.2.0
[0.1.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.1.0
