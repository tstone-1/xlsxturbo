# Roadmap

Planned features for xlsxturbo, ordered by priority.

## High Priority

Features that would enable more migrations from pandas/polars write_excel.

- [x] **Column auto-width** - Automatically fit column widths to content (v0.3.0)
- [x] **Custom column widths** - Set specific widths per column (v0.4.0)
- [x] **Custom row heights** - Set specific heights per row (v0.4.0)
- [x] **Per-sheet options in dfs_to_xlsx** - Allow different settings per sheet (v0.5.0)
- [x] **Header styling** - Option for bold/colored headers (`header_format` parameter) (v0.6.0)
- [x] **Freeze panes** - Freeze header row for easier scrolling (v0.3.0)
- [x] **Global column width cap** - `column_widths={'_all': value}` to limit all column widths (v0.6.0)
- [x] **Table name** - `table_name` parameter for named Excel tables (v0.6.0)

## Medium Priority

Power user features for more control over output.

- [x] **Multi-core support** - Parallel CSV parsing with rayon (~7% speedup for large files)
- [x] **Column formatting with wildcards** - `column_formats` with pattern matching (v0.7.0)
  - Supports: `prefix*`, `*suffix`, `*contains*`, exact match
  - Format options: bg_color, font_color, num_format, bold, italic, underline, border
- [ ] **Row-level cell formatting** - Conditional styling based on cell values
- [ ] **Merged cells** - Merge cell ranges for headers/documentation sheets
- [ ] **Conditional formatting** - Color scales, data bars, icon sets
- [x] **Table styles** - Create Excel tables with auto-filters and 61 built-in styles (v0.3.0)
- [ ] **Column type hints** - Override auto-detection for specific columns

## Lower Priority

Niche features for specific use cases.

- [ ] **Formulas** - Add calculated columns with Excel formulas
- [ ] **Data validation** - Dropdowns, input constraints
- [x] **Constant memory mode** - Handle very large datasets with minimal RAM (v0.4.0)
- [ ] **Append mode** - Add sheets to existing workbook

## Completed

- [x] Column formatting with wildcards via `column_formats` parameter (v0.7.0)
- [x] Global column width cap with `column_widths={'_all': value}` (v0.6.0)
- [x] Table name parameter with `table_name` (v0.6.0)
- [x] Header styling with `header_format` (v0.6.0)
- [x] CSV to XLSX conversion with type detection (v0.1.0)
- [x] pandas DataFrame support (v0.2.0)
- [x] polars DataFrame support (v0.2.0)
- [x] Multi-sheet workbooks with `dfs_to_xlsx()` (v0.2.0)
- [x] Parallel CSV parsing with `parallel=True` (v0.2.0)
- [x] Date/datetime formatting (v0.1.0)
- [x] CLI tool (v0.1.0)
- [x] Custom column widths with `column_widths` parameter (v0.4.0)
- [x] Custom row heights with `row_heights` parameter (v0.4.0)
- [x] Constant memory mode with `constant_memory` parameter (v0.4.0)
- [x] Per-sheet options in `dfs_to_xlsx()` (v0.5.0)
