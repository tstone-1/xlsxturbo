# Multi-sheet workbooks

`dfs_to_xlsx` writes several DataFrames into one workbook, one sheet each. Options
given at the top level apply to every sheet; options given per sheet override them.
Two entries can also share a sheet: see
[Several frames on one sheet](#several-frames-on-one-sheet).

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
- `freeze_panes` (bool or (row, col)): Freeze the header row, or freeze at a chosen cell
- `autofilter` (bool): Filter buttons on the header row, without a table
- `tab_color` (str|None): Sheet tab colour, or None to remove the workbook-wide colour
- `startrow` (int): Zero-based row of the header
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

## Sheet tab colours

Set `tab_color` per sheet to colour the tabs by group. A sheet without the option keeps
the default tab.

```python
xlsxturbo.dfs_to_xlsx([
    (inputs,  "Inputs",  {"tab_color": "#4472C4"}),
    (results, "Results", {"tab_color": "#70AD47"}),
    (legend,  "Legend"),
], "report.xlsx")
```

As a keyword argument, `tab_color` colours every sheet. A per-sheet value overrides it,
and a per-sheet `None` removes it for that sheet.

## Several frames on one sheet

Entries that name the same sheet each write their frame to it, at their own `startrow`.
Use it for small tables that belong together, such as key figures with a breakdown below.

```python
summary = pd.DataFrame({"Item": ["rows checked", "rows flagged"], "Value": ["120", "7"]})
reasons = pd.DataFrame({"Reason": ["missing", "outdated"], "Count": [4, 3]})

xlsxturbo.dfs_to_xlsx([
    (summary, "Summary"),
    (reasons, "Summary", {"startrow": len(summary) + 3}),
], "report.xlsx")
```

Every frame starts in column A. The call returns one `(rows, columns)` per entry, as it
does with one frame per sheet. Sheet names are compared without regard to case, as Excel
compares them, and the sheet keeps the spelling of its first entry.

**Frames may not overlap.** Two frames that share a row raise `FrameOverlapError`, a
`WorkbookValidationError`, before anything is written, naming both entries and their rows. Two entries for one sheet
without a `startrow` both start in row 1, so they overlap.

Which entry an option belongs to depends on what the option describes:

| The option describes | Options | Rule |
|---|---|---|
| The frame | `header`, `table_style`, `table_name`, `autofilter`, `startrow`, `header_format`, `column_formats`, `conditional_formats`, `formula_columns`, `validations` | Each entry has its own. A keyword argument is the default for every frame |
| The whole sheet | `autofit`, `freeze_panes`, `tab_color`, `column_widths`, `row_heights` | The first entry for the sheet sets them. A later entry that sets one raises `ConfigurationError` |
| A cell reference | `cells`, `merged_ranges`, `hyperlinks`, `comments`, `rich_text`, `images`, `checkboxes`, `textboxes`, `charts`, `sparklines` | Each entry applies what its own options dict gives. A keyword argument is applied once per sheet |

- `autofit` fits each column to the longest content of all frames on the sheet.
  `freeze_panes=True` freezes below the header of the first entry.
- A sheet holds one plain filter, so `autofilter=True` on two frames of a sheet raises
  `ConfigurationError`. Give each frame a `table_style` if both need filter buttons: a
  table has its own.
- The cell-addressed options are applied after every frame is written, so a `cells`
  entry can overwrite a value of any frame on the sheet.
- With `constant_memory=True` rows are written once and in order. List the frames of a
  sheet from top to bottom; any other order raises `ConfigurationError`. In normal mode
  the order in the list is free.
