# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.10.1] - 2026-01-16

### Changed
- **Benchmark suite reorganization** - Moved benchmarks to `benchmarks/` directory
  - New `benchmarks/benchmark.py` - comprehensive comparison vs polars, pandas+openpyxl, pandas+xlsxwriter
  - Moved `benchmark_parallel.py` to `benchmarks/`
  - Removed obsolete `benchmark.py` (referenced old Rust binary)
- **README Performance section** - Updated with reproducible benchmark methodology
  - Changed performance claim from "~25x faster" to "~6x faster" (accurate for typical workloads)
  - Added disclaimer that results vary by system
  - Linked to Benchmarking section for running your own tests

## [0.10.0] - 2026-01-16

### Added
- **Comments/Notes** - Add cell annotations with optional author
  - Simple text: `comments={'A1': 'Note text'}`
  - With author: `comments={'A1': {'text': 'Note', 'author': 'John'}}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Data Validation** - Add dropdowns and constraints to columns
  - List (dropdown): `validations={'Status': {'type': 'list', 'values': ['Open', 'Closed']}}`
  - Whole number: `validations={'Score': {'type': 'whole_number', 'min': 0, 'max': 100}}`
  - Decimal: `validations={'Price': {'type': 'decimal', 'min': 0.0, 'max': 999.99}}`
  - Text length: `validations={'Code': {'type': 'text_length', 'min': 3, 'max': 10}}`
  - Supports input/error messages: `input_title`, `input_message`, `error_title`, `error_message`
  - Supports column patterns (like `column_formats`)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Rich Text** - Multiple formats within a single cell
  - Format segments: `rich_text={'A1': [('Bold', {'bold': True}), ' normal text']}`
  - Supports: `bold`, `italic`, `font_color`, `bg_color`, `font_size`, `underline`
  - Mix formatted and plain text segments
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Images** - Embed PNG, JPEG, GIF, BMP images in cells
  - Simple path: `images={'B5': 'logo.png'}`
  - With options: `images={'B5': {'path': 'logo.png', 'scale_width': 0.5, 'scale_height': 0.5}}`
  - Options: `path`, `scale_width`, `scale_height`, `alt_text`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides

### Notes
- All new features are disabled in `constant_memory` mode (they require random access)
- Data validation list values are limited to 255 total characters (Excel limitation)

## [0.9.0] - 2026-01-15

### Added
- **Conditional formatting** - Visual formatting based on cell values
  - `2_color_scale`: Gradient from min_color to max_color
  - `3_color_scale`: Three-color gradient with min/mid/max colors
  - `data_bar`: In-cell bar chart with customizable color, direction, solid fill
  - `icon_set`: Traffic lights, arrows, flags (3/4/5 icons), with reverse and icons_only options
  - Supports column name patterns: `'price_*': {'type': 'data_bar', ...}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `conditional_formats={'score': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'}}`
- **Formula columns** - Add calculated columns with Excel formulas
  - Use `{row}` placeholder for row numbers (1-based)
  - Columns appear after data columns
  - Order preserved (first formula = first new column)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `formula_columns={'Total': '=A{row}+B{row}', 'Percentage': '=C{row}/D{row}*100'}`
- **Merged cells** - Merge cell ranges for headers, titles, and grouped labels
  - Uses Excel notation for ranges (e.g., 'A1:D1')
  - Optional formatting with HeaderFormat options (bold, colors, etc.)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `merged_ranges=[('A1:C1', 'Title'), ('A2:C2', 'Subtitle', {'bold': True})]`
- **Hyperlinks** - Add clickable links to cells
  - Uses Excel notation for cell reference (e.g., 'A1', 'B5')
  - Optional display text (defaults to URL if not provided)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `hyperlinks=[('A2', 'https://example.com'), ('B2', 'https://google.com', 'Google')]`

## [0.8.0] - 2026-01-15

### Added
- **Date order parameter** - `date_order` for `csv_to_xlsx()` to handle ambiguous dates
  - `"auto"` (default): ISO first, then European (DMY), then US (MDY)
  - `"mdy"` or `"us"`: US format where 01-02-2024 = January 2nd
  - `"dmy"` or `"eu"`: European format where 01-02-2024 = February 1st
  - Also available in CLI: `--date-order us`
- **BUILD.md** - Developer guide for building, testing, and releasing

### Fixed
- **Pattern matching order** - `column_formats` patterns now match in definition order (first match wins). Previously, HashMap iteration order was non-deterministic.
- **Empty DataFrame with table_style** - No longer crashes; tables are skipped when DataFrame has no data rows
- **Hex color validation** - Colors like `#FF` now raise descriptive error instead of silently misparsing
- **Invalid table_style validation** - Unknown styles now raise error instead of silently defaulting to Medium9
- **CLI division by zero** - Instant conversions now show "instant rows/sec" instead of "inf"

### Changed
- Uses `indexmap` crate to preserve pattern insertion order
- Updated `pyo3` 0.23 → 0.27, `rust_xlsxwriter` 0.79 → 0.92
- Added Dependabot for automated dependency updates

## [0.7.0] - 2025-12-28

### Added
- **Column formatting with wildcards** - `column_formats` parameter for styling columns by pattern
  - Wildcard patterns: `prefix*`, `*suffix`, `*contains*`, or exact match
  - Format options: `bg_color`, `font_color`, `num_format`, `bold`, `italic`, `underline`, `border`
  - Example: `column_formats={'price_*': {'bg_color': '#D6EAF8', 'num_format': '$#,##0.00', 'border': True}}`
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

[0.10.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.1
[0.10.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.0
[0.9.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.9.0
[0.8.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.8.0
[0.7.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.7.0
[0.6.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.6.0
[0.5.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.5.0
[0.4.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.1
[0.4.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.0
[0.3.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.3.0
[0.2.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.2.0
[0.1.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.1.0
