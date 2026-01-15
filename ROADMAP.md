# Roadmap

Planned features for xlsxturbo, ordered by priority.

xlsxturbo's niche is **high-performance DataFrame export** with a simple API. The underlying
rust_xlsxwriter library already supports most advanced Excel features—they just need Python
bindings exposed.

## High Priority

Features that close the gap with XlsxWriter/openpyxl for common use cases.

- [ ] **Conditional formatting** - Color scales, data bars, icon sets, formula-based rules
  - rust_xlsxwriter: supported
  - Enables: heatmaps, progress indicators, threshold highlighting
  - API: `conditional_formats={'column': {'type': '3_color_scale', ...}}`

- [ ] **Formulas** - Add calculated columns with Excel formulas
  - rust_xlsxwriter: supported (including Excel 365 dynamic arrays)
  - Enables: SUM, VLOOKUP, custom calculations that update in Excel
  - API: `formula_columns={'Total': '=B{row}*C{row}'}`

- [ ] **Merged cells** - Merge cell ranges for headers/documentation sheets
  - rust_xlsxwriter: supported
  - Enables: report headers, grouped labels, multi-row titles
  - API: `merged_ranges=[('A1:D1', 'Report Title', format_dict)]`

- [ ] **Hyperlinks** - Add clickable links to cells
  - rust_xlsxwriter: supported
  - Enables: links to URLs, other sheets, or external files
  - API: `hyperlinks={'A1': 'https://example.com'}` or in column_formats

## Medium Priority

Power user features for richer Excel output.

- [ ] **Data validation** - Dropdowns, input constraints
  - rust_xlsxwriter: supported
  - Enables: dropdown lists, numeric ranges, custom error messages
  - API: `validations={'Status': {'type': 'list', 'values': ['Open', 'Closed']}}`

- [ ] **Images** - Embed PNG/JPEG/GIF/BMP images
  - rust_xlsxwriter: supported
  - Enables: logos, charts generated externally, screenshots
  - API: `images={'A1': 'logo.png'}` or `images=[('B5', 'chart.png', options)]`

- [ ] **Comments/notes** - Add cell comments
  - rust_xlsxwriter: supported
  - Enables: documentation, review notes, explanations
  - API: `comments={'A1': 'This is the header'}` or in column_formats

- [ ] **Column type hints** - Override auto-detection for specific columns
  - Enables: force text for ZIP codes, force numbers for string-encoded IDs
  - API: `column_types={'zip_code': 'text', 'amount': 'number'}`

- [ ] **Rich text** - Multiple formats within a single cell
  - rust_xlsxwriter: supported
  - Enables: bold + italic in same cell, colored substrings
  - API: `rich_text={'A1': [('Bold part', {'bold': True}), (' normal')]}`

## Lower Priority

Niche features for specific use cases.

- [ ] **Charts** - Bar, line, pie, scatter, etc.
  - rust_xlsxwriter: supported (all standard chart types)
  - Enables: visual data representation within Excel
  - Complex API, may be better served by matplotlib → image

- [ ] **Sparklines** - Mini inline charts
  - rust_xlsxwriter: supported
  - Enables: trend indicators in cells
  - API: `sparklines={'D2:D10': {'range': 'A2:C10', 'type': 'line'}}`

- [ ] **Append mode** - Add sheets to existing workbook
  - rust_xlsxwriter: not supported (write-only by design)
  - Would require openpyxl hybrid approach

- [ ] **Checkboxes** - Interactive checkboxes in cells
  - rust_xlsxwriter: supported
  - Enables: todo lists, selection interfaces

- [ ] **Textboxes** - Floating text annotations
  - rust_xlsxwriter: supported
  - Enables: callouts, annotations outside cell grid

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
- [x] Column auto-width (v0.3.0)
- [x] Table styles with 61 built-in options (v0.3.0)
- [x] Freeze panes (v0.3.0)
- [x] Multi-core CSV support (v0.2.0)
