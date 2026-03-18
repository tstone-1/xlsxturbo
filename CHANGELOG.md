# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.12.3] - 2026-03-17

### Fixed
- **Linux x86_64 wheels now work on Python 3.9+** - Release workflow switched from `manylinux2014` (Python 3.8 only) to `manylinux_2_28` with `--find-interpreter`, producing proper `abi3` wheels instead of `cp38-cp38` wheels
- **Invalid datetime/date values now raise errors** - Previously, invalid dates from Python objects (e.g., month=13) silently fell through to string conversion; now returns a clear error message
- **Improved test assertion** - `test_parse_float` uses `assert!(matches!(...))` instead of bare `panic!`

### Improved
- **Reduced internal parameter counts** - Introduced `WriteConfig` struct to group scalar sheet configuration, reducing `write_sheet_data` from 11 to 5 parameters and `apply_worksheet_features` from 16 to 10
- **Contextual error messages throughout** - All `.map_err(|e| e.to_string())` calls replaced with descriptive `format!("Context: {}", e)` messages
- **CI cargo caching** - Added `actions/cache@v4` for Rust dependencies across all CI jobs
- **CI pip caching** - Added `cache: 'pip'` to all `setup-python` steps
- **CI platform coverage** - Windows and macOS now test Python 3.9 + 3.12 (was only 3.12)

### Documentation
- **README** - Added documentation for `defined_names` and `cells` parameters with usage examples
- **README** - Updated feature list with v0.11.0+ features (defined names, arbitrary cells, borders, alignment)
- **README/type stubs** - Added `formula_columns` and `cells` to `constant_memory` disabled features list
- **Type stubs** - Fixed `column_widths` and `row_heights` value types from `float` to `int | float`
- **Type stubs** - Documented validation `min`/`max` default behavior

### Tests
- **Per-sheet cells** - Added `TestCellsPerSheet` (4 tests) covering 3-tuple SheetOptions with cells
- **Cells formatting** - Added `TestCellsFormatting` (5 tests) covering alignment and wrap_text options

## [0.12.2] - 2026-03-16

### Fixed
- **`autofit=True` + `column_widths={'_all': N}` now caps instead of overriding** - columns are autofit to content then capped at N (`min(content_width, cap)`), matching the documented behavior (#13)

## [0.12.1] - 2026-03-16

### Fixed
- **Table creation crash when `header=False`** - Excel tables require a header row; table creation is now skipped when `include_header` is false, preventing `Table must have at least one row` errors (#12)

## [0.12.0] - 2026-03-16

### Added
- **Per-side border styles** - Fine-grained border control for `column_formats` and `header_format`
  - `border` now accepts string style names: `'border': 'thick'` (all 4 sides)
  - Per-side keys: `border_left`, `border_right`, `border_top`, `border_bottom`
  - Per-side keys accept `True` (= thin) or a style name string
  - `border_color` for setting border color (`'#RRGGBB'` or named color, applies to all sides)
  - 13 border styles: thin, medium, thick, dashed, dotted, double, hair, medium_dashed, dash_dot, medium_dash_dot, dash_dot_dot, medium_dash_dot_dot, slant_dash_dot
  - Works in both `column_formats` and `header_format`
  - Backward compatible: `'border': True` still works (thin, all sides)
  - Example: `column_formats={'col': {'border_right': 'thick'}}`
- **Text alignment** - `align_horizontal`, `align_vertical`, and `wrap_text` formatting options
  - Available in `header_format`, `column_formats`, `merged_ranges`, and `cells`
  - Horizontal: `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'center_across'`, `'distributed'`
  - Vertical: `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`
  - `wrap_text: True` enables text wrapping within cells
  - Example: `column_formats={'description': {'align_horizontal': 'left', 'wrap_text': True}}`
- **Rule-based conditional formatting** - `type: 'cell'` in `conditional_formats` for value-based highlighting
  - Comparison criteria: `equal_to`, `not_equal_to`, `greater_than`, `less_than`, `greater_than_or_equal_to`, `less_than_or_equal_to`
  - Range criteria: `between`, `not_between` (with `min_value`/`max_value`)
  - Text criteria: `containing`, `not_containing`, `begins_with`, `ends_with`
  - Special: `blanks`, `no_blanks`
  - `format` key accepts all column format options (bg_color, font_color, bold, border, etc.)
  - Multiple rules per column: pass a list of config dicts instead of a single dict
  - Example: `conditional_formats={'status': {'type': 'cell', 'criteria': 'equal_to', 'value': 'ERROR', 'format': {'bg_color': '#FF0000'}}}`

## [0.11.0] - 2026-03-15

### Added
- **Defined names** - `defined_names` parameter for workbook-level named ranges
  - Dict mapping name to Excel reference: `defined_names={"MyRange": "=Sheet1!$A$1:$D$100"}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()`
- **Arbitrary cell writes** - `cells` parameter for writing values to specific cells
  - Simple values: `cells={"B9": "Label", "B10": 42}`
  - With number format: `cells={"D6": {"value": "934728173849", "num_format": "@"}}`
  - Cells are written after DataFrame data, allowing overwrite of data cells
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides

## [0.10.6] - 2026-03-12

### Fixed
- **Polars DataFrame detection now checks module name** - `is_polars_dataframe` checks `__module__` instead of duck-typing attributes, preventing misidentification of non-DataFrame objects with `.schema` attribute (e.g., Pydantic models)

### Changed
- **CI uses pytest** - Integration tests now run via `pytest tests/ -v` instead of `python tests/test_features.py`, with proper test discovery and failure reporting
- **CI Python dependencies pinned** - `pandas>=2,<3`, `polars>=1,<2`, `openpyxl>=3,<4`, `pytest>=8,<9`, `maturin>=1.4,<2.0` to prevent unexpected breakage from upstream releases
- **`parse_table_style` uses macro** - Replaced 79-line match statement with `table_style_match!` macro; added version sync comment for `rust_xlsxwriter` 0.94
- **Dependencies** - Updated `rust_xlsxwriter` 0.93 -> 0.94, `actions/upload-artifact` v4 -> v7, `actions/download-artifact` v7 -> v8

### Refactored
- **Extracted `apply_worksheet_features` from `write_sheet_data`** - Feature application (table, formulas, conditional formats, freeze panes, widths, heights, merged ranges, hyperlinks, comments, validations, rich text, images) moved to a dedicated function with a single `constant_memory` early-return gate instead of 12 scattered checks
- **Removed redundant `constant_memory` parameter** from `apply_column_widths_with_autofit_cap` (caller already guards)

### Tests
- Added 22 new Rust unit tests: `parse_cell_ref` (basic, case-insensitive, max column, overflow, Excel max, row zero, empty, no row, no column), `parse_cell_range` (basic, invalid), `parse_color` (hex, named, invalid, whitespace), `sanitize_table_name` (valid, special chars, digit prefix, truncation, empty), `parse_table_style` (valid, invalid), `naive_date_to_excel` (epoch, known date), `DateOrder::parse`

## [0.10.5] - 2026-03-02

### Fixed
- **Formula columns overwrite data when `header=False`** - `apply_formula_columns` no longer hardcodes the formula header to row 0; headers are only written when `include_header=True`, preventing data loss when combining `header=False` with `formula_columns`
- **`parse_cell_ref` overflow on adversarial input** - Column letter fold now uses `checked_mul`/`checked_add` instead of wrapping arithmetic, returning a clear error on pathologically long column strings

### Changed
- **Minimum Python version raised to 3.9** - Type stubs use PEP 585 lowercase generics (`list[str]`, `dict[str, ...]`) which require Python 3.9+. Python 3.8 reached EOL in October 2024. Updated `requires-python`, PyO3 ABI tag (`abi3-py39`), and classifiers accordingly
- **`clap` is now an optional dependency** - CLI argument parser is gated behind a `cli` feature flag (enabled by default), reducing compile time for library-only builds (Python extension)
- **CI Python test matrix** - Integration tests now run on Python 3.9, 3.12, and 3.14 (previously only 3.12)
- **Completed Python docstrings** - `df_to_xlsx` and `dfs_to_xlsx` docstrings now document all parameters including `table_name`, `formula_columns`, `merged_ranges`, `hyperlinks`, `comments`, `validations`, `rich_text`, and `images`
- **`_all` width cap documentation** - Clarified that `_all` sets a uniform width rather than capping autofit results, since `rust_xlsxwriter` does not expose autofitted widths for reading

### Refactored
- **`write_py_value_with_format` reduced from 252 to ~90 lines** - Extracted `write_str`, `write_num`, `write_bool`, `write_int`, `write_float` helpers to eliminate 10x duplicated format/no-format dispatch
- **`extract_sheet_info` reduced from 170 to ~60 lines** - Introduced `extract_scalar!`, `extract_dict_field!`, `extract_list_field!` macros to replace 13 copy-pasted extraction blocks
- **`pydict_to_hashmap` helper** - Replaced 6 duplicated `HashMap<String, Py<PyAny>>` extraction blocks with a single reusable function
- **Explicit imports in `features.rs`** - Replaced `use crate::types::*` glob import with explicit type imports
- **Dependencies** - Updated indirect dependencies via `cargo update` (js-sys, wasm-bindgen, tempfile, zlib-rs)

## [0.10.4] - 2026-02-23

### Fixed
- **Boolean column formatting ignored** - Boolean values now correctly receive column formatting via `write_boolean_with_format` instead of being written without format
- **formula_columns not disabled in constant_memory mode** - Formula columns are now correctly skipped when `constant_memory=True`, matching the documented behavior
- **Cell reference column overflow** - `parse_cell_ref` now validates columns against Excel's maximum (XFD = 16384) using u32 intermediate arithmetic instead of silently wrapping u16
- **Unchecked arithmetic in formula/row operations** - Row and column index calculations now use `checked_add` to prevent silent overflow on extremely large datasets
- **Dead code** - Removed unreachable `extract::<bool>()` fallback in `write_py_value_with_format`

### Changed
- **Deduplicated write logic** - Extracted shared `write_sheet_data` function (~200 lines) used by both `convert_dataframe_to_xlsx` and `dfs_to_xlsx`, eliminating ~300 lines of duplicated code
- **Reference-based option merging** - New `EffectiveOpts` struct uses references instead of cloning, avoiding unnecessary allocations when merging per-sheet and global options in `dfs_to_xlsx`
- **`extract_sheet_info` refactored** - Now delegates to existing `extract_*` functions instead of reimplementing parsing inline
- **constant_memory warnings** - When `constant_memory=True` is used with incompatible options, a Python `warnings.warn()` is now emitted listing the disabled features
- **Dependencies** - Updated `pyo3` 0.28.1 → 0.28.2 (fixes RUSTSEC-2026-0013)
- **Metadata** - Added Python 3.13 and 3.14 classifiers to pyproject.toml

## [0.10.3] - 2026-02-16

### Fixed
- **Large integer precision loss** - Integers exceeding 2^53 are now written as strings instead of silently losing precision when cast to f64
- **Numpy int extraction order** - Numpy integer types (e.g. `numpy.int64`) now go through i64 extraction before f64 fallback, preventing precision loss for large values
- **Unchecked column index casts** - CSV sequential path now uses `u16::try_from` with clear error messages instead of unchecked `as u16` casts
- **CLI branding** - Replaced 4 remaining `fast_xlsx` references in `main.rs` with `xlsxturbo`
- **Undocumented `.unwrap()`** - Changed Excel epoch date `.unwrap()` to `.expect()` with explanation

### Changed
- **DataFrame type detection** - Extracted `is_polars_dataframe()` and `extract_columns()` helpers into `types.rs`, replacing 4 duplicated detection blocks across `convert.rs` and `lib.rs`
- **Documentation** - Added `constant_memory` disabled features list to Python docstrings; added "Known Limitations" section to README (datetime precision, large integers)
- **CI** - Bumped `actions/setup-python` from v5 to v6

### Tests
- Added `TestUnicodeAndSpecialData` class with 10 new tests: Unicode/CJK column names and data, emoji, mixed-type columns, None/NaT/pd.NA handling, all-None columns, large integer precision, CSV with BOM, CSV with CRLF, CSV with quoted delimiters, Polars Unicode
- Total: 93 Python integration tests, 12 Rust unit tests

## [0.10.2] - 2026-02-06

### Fixed
- **Wildcard pattern panic** - `matches_pattern("*")` no longer panics on lone `*` pattern
- **Silent datetime defaults** - Datetime/date attribute extraction now propagates errors instead of silently defaulting to 1900
- **Index overflow safety** - Row count and column count use checked casts (`u32::try_from`, `u16::try_from`) instead of `as` casts
- **PyPI URL in release workflow** - Fixed leftover `fast_xlsx` reference to `xlsxturbo`
- **Validation type aliases** - Added `whole`, `integer`, `number`, `textlength`, `length` aliases to type stubs
- **CHANGELOG accuracy** - `constant_memory` entry now lists all 12 disabled features

### Changed
- **Module split** - Split monolithic `lib.rs` into `convert.rs`, `parse.rs`, `features.rs`, `types.rs`
- **Deduplicated option extraction** - New `ExtractedOptions` struct reduces `convert_dataframe_to_xlsx` from 22 to 12 parameters
- **Merged format parsers** - `parse_header_format` and `parse_column_format` share a single `parse_format_dict` implementation
- **Dependencies** - Updated `pyo3` 0.27 -> 0.28, `rust_xlsxwriter` 0.92 -> 0.93
- **CI** - Added Python integration test job (83 tests with pandas, polars, openpyxl)
- **PEP 561** - Added `py.typed` marker file for type checker support

### Tests
- Added `TestConditionalFormatting` (5 tests), `TestConstantMemoryMode` (3 tests), `TestRowHeights` (3 tests)
- Upgraded ~10 shallow tests with openpyxl content verification (column widths, table names, header formats, validations, comments)
- Total: 83 Python integration tests, 12 Rust unit tests

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
  - Note: Disables `table_style`, `freeze_panes`, `row_heights`, `autofit`, `conditional_formats`, `formula_columns`, `merged_ranges`, `hyperlinks`, `comments`, `validations`, `rich_text`, and `images`
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

[0.10.4]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.4
[0.10.3]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.3
[0.10.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.2
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
