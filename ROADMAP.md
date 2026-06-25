# Roadmap

xlsxturbo's niche is **high-performance DataFrame and CSV export** with a simple Python API.
Most commonly requested Excel features are now implemented; this roadmap tracks the remaining
feature gaps first, then records completed milestones.

## Remaining

Niche features for specific use cases.

- [ ] **Append mode** - Add sheets to existing workbook (deferred - not planned as a built-in)
  - rust_xlsxwriter is write-only by design, so this would require an openpyxl hybrid that
    reads and re-materializes the existing workbook. That pulls in a heavy runtime dependency
    and makes the one operation carrying the feature openpyxl-bound (slow) - at odds with the
    library's pure-Rust, high-performance thesis. Recommended path is a documented recipe
    (read with openpyxl, write new sheets with xlsxturbo to a copy) rather than `append=True`.

## Completed

### Common Export Features

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

### Rich Workbook Features

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

### Recent Milestones

- [x] Sparklines via `sparklines` parameter (v0.16.0)
- [x] Native Excel charts via `charts` parameter (v0.15.0)
- [x] Textboxes via `textboxes` parameter (v0.14.0)
- [x] Checkboxes via `checkboxes` parameter (v0.13.0)
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
