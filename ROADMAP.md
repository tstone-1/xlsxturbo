# Roadmap

Planned features for xlsxturbo, ordered by priority.

## High Priority

Features that address common use cases and add significant value.

- [ ] **Column auto-width** - Automatically fit column widths to content
- [ ] **Header styling** - Option for bold/colored headers
- [ ] **Freeze panes** - Freeze header row for easier scrolling
- [ ] **Multiple sheets** - Export multiple DataFrames to one workbook

## Medium Priority

Power user features for more control over output.

- [ ] **Multi-core support** - Parallelize parsing/type detection with rayon (1.5-2x speedup for large files)
- [ ] **Cell formatting options** - Custom number/date formats per column
- [ ] **Conditional formatting** - Color scales, data bars, icon sets
- [ ] **Named tables** - Create Excel tables with auto-filters
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
- [x] Date/datetime formatting (v0.1.0)
- [x] CLI tool (v0.1.0)
