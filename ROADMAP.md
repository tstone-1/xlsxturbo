# Roadmap

Planned features for xlsxturbo, ordered by priority.

## High Priority

Features that address common use cases and add significant value.

- [x] **Column auto-width** - Automatically fit column widths to content (v0.3.0)
- [x] **Custom column widths** - Set specific widths per column
- [x] **Custom row heights** - Set specific heights per row
- [ ] **Header styling** - Option for bold/colored headers
- [x] **Freeze panes** - Freeze header row for easier scrolling (v0.3.0)

## Medium Priority

Power user features for more control over output.

- [x] **Multi-core support** - Parallel CSV parsing with rayon (~7% speedup for large files)
- [ ] **Cell formatting options** - Custom number/date formats per column
- [ ] **Cell styling** - Background color, font color, bold, borders
- [ ] **Merged cells** - Merge cell ranges for headers/documentation sheets
- [ ] **Conditional formatting** - Color scales, data bars, icon sets
- [x] **Named tables** - Create Excel tables with auto-filters and 61 built-in styles (v0.3.0)
- [ ] **Column type hints** - Override auto-detection for specific columns

## Lower Priority

Niche features for specific use cases.

- [ ] **Formulas** - Add calculated columns with Excel formulas
- [ ] **Data validation** - Dropdowns, input constraints
- [ ] **Streaming write** - Handle datasets larger than available RAM
- [ ] **Append mode** - Add sheets to existing workbook

## Completed

- [x] CSV to XLSX conversion with type detection (v0.1.0)
- [x] pandas DataFrame support (v0.2.0)
- [x] polars DataFrame support (v0.2.0)
- [x] Multi-sheet workbooks with `dfs_to_xlsx()` (v0.2.0)
- [x] Parallel CSV parsing with `parallel=True` (v0.2.0)
- [x] Date/datetime formatting (v0.1.0)
- [x] CLI tool (v0.1.0)
- [x] Custom column widths with `column_widths` parameter
- [x] Custom row heights with `row_heights` parameter
