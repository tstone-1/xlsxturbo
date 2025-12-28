# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.7.0] - 2025-12-28

### Added
- **Column formatting with wildcards** - `column_formats` parameter for styling columns by pattern
  - Wildcard patterns: `prefix*`, `*suffix`, `*contains*`, or exact match
  - Format options: `bg_color`, `font_color`, `num_format`, `bold`, `italic`, `underline`, `border`
  - Example: `column_formats={'mcpt_*': {'bg_color': '#D6EAF8', 'num_format': '0.00000', 'border': True}}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()`
  - Per-sheet column formats via options dict in `dfs_to_xlsx()`

## [0.6.0] - 2025-12-08

### Added
- **Global column width cap** - `column_widths={'_all': 50}` to cap all columns at a maximum width
  - Can be combined with specific column widths: `{0: 20, '_all': 50}` (specific overrides '_all')
  - Works with autofit as a cap: `autofit=True, column_widths={'_all': 30}` fits then caps
- **Table name parameter** - `table_name="MyTable"` to set custom Excel table names
  - Invalid characters are automatically sanitized (spaces/special chars become underscores)
  - Names starting with digits get underscore prefix (Excel requirement)
  - Per-sheet table names in `dfs_to_xlsx()` via options dict
- **Header styling** - `header_format={'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'}`
  - Supported options: `bold`, `italic`, `font_color`, `bg_color`, `font_size`, `underline`
  - Colors accept hex (`#RRGGBB`) or named colors (white, black, red, blue, etc.)
  - Per-sheet header formats in `dfs_to_xlsx()` via options dict
- Per-sheet options now support: `table_name`, `header_format`, `column_widths` with '_all'

### Changed
- `column_widths` parameter now accepts both integer keys (`{0: 20}`) and string keys (`{"_all": 50}`)

## [0.5.0] - 2025-12-08

### Added
- **Per-sheet options for `dfs_to_xlsx()`** - override global settings per sheet
  - Each sheet can now be a 3-tuple: `(df, sheet_name, options_dict)`
  - Options dict supports: `header`, `autofit`, `table_style`, `freeze_panes`, `column_widths`, `row_heights`
  - Old 2-tuple API `(df, sheet_name)` still works (backward compatible)
  - Example: `[(df1, "Data", {"table_style": "Medium2"}), (df2, "Instructions", {"header": False})]`
- `SheetOptions` TypedDict for type hints

### Changed
- `dfs_to_xlsx()` signature now accepts mixed tuple formats internally
- Updated type stubs with new `SheetOptions` class and updated `dfs_to_xlsx` signature

## [0.4.1] - 2025-12-07

### Fixed
- Updated type stubs to include v0.4.0 parameters (`column_widths`, `row_heights`, `constant_memory`)
- Cleaned up ROADMAP.md

## [0.4.0] - 2025-12-07

### Added
- `constant_memory` parameter - minimize RAM usage for very large files
  - Uses rust_xlsxwriter's streaming mode to flush rows to disk
  - Ideal for files with millions of rows
  - Note: Disables `table_style`, `freeze_panes`, `row_heights`, and `autofit`
  - Column widths still work in constant memory mode
  - Example: `xlsxturbo.df_to_xlsx(df, "big.xlsx", constant_memory=True)`
- `column_widths` parameter - set custom column widths by index
  - Dict mapping column index (0-based) to width in characters
  - Example: `column_widths={0: 25, 1: 15, 3: 30}`
- `row_heights` parameter - set custom row heights by index
  - Dict mapping row index (0-based) to height in points
  - Example: `row_heights={0: 22, 5: 30}`
- All new parameters available in `df_to_xlsx()` and `dfs_to_xlsx()`

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

[0.7.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.7.0
[0.6.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.6.0
[0.5.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.5.0
[0.4.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.1
[0.4.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.0
[0.3.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.3.0
[0.2.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.2.0
[0.1.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.1.0
