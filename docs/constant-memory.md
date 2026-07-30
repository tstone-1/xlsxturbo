# Constant memory mode

`constant_memory=True` writes each row to disk as it is produced instead of holding
the sheet in RAM, which bounds memory use for very large exports. The trade-off is that
options needing a second pass over the finished sheet cannot be applied -- those are
skipped with a warning rather than silently dropped, and the
[capability matrix](capability-matrix.md) lists exactly which survive.

## Constant Memory Mode (Large Files)

For very large files (millions of rows), use `constant_memory=True` to minimize RAM usage:

```python
import xlsxturbo
import polars as pl

# Generate a large DataFrame
large_df = pl.DataFrame({
    'id': range(1_000_000),
    'value': [i * 1.5 for i in range(1_000_000)]
})

# Use constant_memory mode for large files
xlsxturbo.df_to_xlsx(large_df, "big_file.xlsx", constant_memory=True)

# Also works with dfs_to_xlsx
xlsxturbo.dfs_to_xlsx([
    (large_df, "Data")
], "multi_sheet.xlsx", constant_memory=True)
```

**Note:** Constant memory mode emits a `RuntimeWarning` and disables some features that require random access:
- `table_style` (Excel tables)
- `freeze_panes`
- `row_heights`
- `autofit`
- `conditional_formats`
- `formula_columns`
- `merged_ranges`
- `hyperlinks`
- `comments`
- `validations`
- `rich_text`
- `images`
- `checkboxes`
- `textboxes`
- `charts`
- `sparklines`
- `cells`

Plain `column_widths`, `header_format`, and `column_formats` remain supported.
