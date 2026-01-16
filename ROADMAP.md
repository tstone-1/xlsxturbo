# Roadmap

Planned features for xlsxturbo, ordered by priority.

xlsxturbo's niche is **high-performance DataFrame export** with a simple API. The underlying
rust_xlsxwriter library already supports most advanced Excel features—they just need Python
bindings exposed.

## High Priority

Features that close the gap with XlsxWriter/openpyxl for common use cases.

- [x] **Conditional formatting** - Color scales, data bars, icon sets (v0.9.0)
  - Supports: 2_color_scale, 3_color_scale, data_bar, icon_set
  - API: `conditional_formats={'column': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'}}`

- [x] **Formulas** - Add calculated columns with Excel formulas (v0.9.0)
  - API: `formula_columns={'Total': '=A{row}+B{row}', 'Percentage': '=C{row}/D{row}*100'}`
  - Use `{row}` placeholder for row numbers (1-based)

- [x] **Merged cells** - Merge cell ranges for headers/documentation sheets (v0.9.0)
  - API: `merged_ranges=[('A1:D1', 'Report Title'), ('A2:D2', 'Subtitle', {'bold': True})]`

- [x] **Hyperlinks** - Add clickable links to cells (v0.9.0)
  - API: `hyperlinks=[('A1', 'https://example.com'), ('B1', 'https://google.com', 'Google')]`
  - Optional display text (defaults to URL if omitted)

## Medium Priority

Power user features for richer Excel output.

- [x] **Data validation** - Dropdowns, input constraints (v0.10.0)
  - Types: list (dropdown), whole_number, decimal, text_length
  - API: `validations={'Status': {'type': 'list', 'values': ['Open', 'Closed']}}`

- [x] **Images** - Embed PNG/JPEG/GIF/BMP images (v0.10.0)
  - Options: scale_width, scale_height, alt_text
  - API: `images={'B5': 'logo.png'}` or `images={'B5': {'path': 'logo.png', 'scale_width': 0.5}}`

- [x] **Comments/notes** - Add cell comments (v0.10.0)
  - Simple text or with author
  - API: `comments={'A1': 'Note'}` or `comments={'A1': {'text': 'Note', 'author': 'John'}}`

- [x] **Rich text** - Multiple formats within a single cell (v0.10.0)
  - Supports: bold, italic, font_color, bg_color, font_size, underline
  - API: `rich_text={'A1': [('Bold part', {'bold': True}), ' normal']}`

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

- [x] Data validation via `validations` parameter (v0.10.0)
- [x] Images via `images` parameter (v0.10.0)
- [x] Comments/notes via `comments` parameter (v0.10.0)
- [x] Rich text via `rich_text` parameter (v0.10.0)
- [x] Hyperlinks via `hyperlinks` parameter (v0.9.0)
- [x] Merged cells via `merged_ranges` parameter (v0.9.0)
- [x] Formula columns via `formula_columns` parameter (v0.9.0)
- [x] Conditional formatting via `conditional_formats` parameter (v0.9.0)
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
