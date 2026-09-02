# Multi-sheet workbooks

`dfs_to_xlsx` writes several DataFrames into one workbook, one sheet each. Options
given at the top level apply to every sheet; options given per sheet override them.

## Multi-Sheet Workbooks

```python
import xlsxturbo
import pandas as pd

# Write multiple DataFrames to separate sheets
df1 = pd.DataFrame({'product': ['A', 'B'], 'sales': [100, 200]})
df2 = pd.DataFrame({'region': ['East', 'West'], 'total': [500, 600]})

xlsxturbo.dfs_to_xlsx([
    (df1, "Products"),
    (df2, "Regions")
], "report.xlsx")

# With styling applied to all sheets
xlsxturbo.dfs_to_xlsx([
    (df1, "Products"),
    (df2, "Regions")
], "styled_report.xlsx", table_style="Medium2", autofit=True, freeze_panes=True)

# With column widths applied to all sheets
xlsxturbo.dfs_to_xlsx([
    (df1, "Products"),
    (df2, "Regions")
], "report.xlsx", column_widths={0: 20, 1: 15})
```

## Per-Sheet Options

Override global settings for individual sheets using a 3-tuple with options dict:

```python
import xlsxturbo
import pandas as pd

df_data = pd.DataFrame({'Product': ['A', 'B'], 'Price': [10, 20]})
df_instructions = pd.DataFrame({'Step': [1, 2], 'Action': ['Open file', 'Review data']})

# Different settings per sheet:
# - "Data" sheet: has header, table style, autofit
# - "Instructions" sheet: no header (raw data), no table style
xlsxturbo.dfs_to_xlsx([
    (df_data, "Data", {"header": True, "table_style": "Medium2"}),
    (df_instructions, "Instructions", {"header": False, "table_style": None})
], "report.xlsx", autofit=True)

# Old 2-tuple API still works - uses global defaults
xlsxturbo.dfs_to_xlsx([
    (df_data, "Sheet1"),  # Uses global header=True, table_style=None
    (df_instructions, "Sheet2", {"header": False})  # Override just header
], "mixed.xlsx", header=True, autofit=True)
```

A `list` of the same shape is accepted wherever a tuple is shown; the tuple is the
recommended form and the one the type stubs describe.

Available per-sheet options:
- `header` (bool): Include column names as header row
- `autofit` (bool): Automatically adjust column widths
- `table_style` (str|None): Excel table style or None to disable
- `freeze_panes` (bool): Freeze header row
- `column_widths` (dict): Custom column widths
- `row_heights` (dict): Custom row heights
- `table_name` (str): Custom Excel table name
- `header_format` (dict): Header cell styling
- `column_formats` (dict): Column formatting with pattern matching
- `conditional_formats` (dict): Conditional formatting (color scales, data bars, icons)
- `formula_columns` (dict): Calculated columns with Excel formulas (column name -> formula template)
- `merged_ranges` (list): List of (range, text) or (range, text, format) tuples to merge cells
- `hyperlinks` (list): List of (cell, url) or (cell, url, display_text) tuples to add clickable links
- `comments` (dict): Cell comments/notes (cell_ref -> text or {text, author})
- `validations` (dict): Data validation rules (column name/pattern -> validation config)
- `rich_text` (dict): Rich text with multiple formats (cell_ref -> list of segments)
- `images` (dict): Embedded images (cell_ref -> path or {path, scale_width, scale_height, alt_text})
- `checkboxes` (dict): Interactive cell checkboxes (cell_ref -> bool or {checked, format})
- `textboxes` (dict): Floating text shapes (cell_ref -> text or textbox options)
- `charts` (dict): Native Excel charts (cell_ref -> chart options)
- `sparklines` (dict): Mini in-cell charts (location ref -> sparkline options; range key = grouped)
- `cells` (dict): Arbitrary cell writes (cell_ref -> value or {value, num_format})
