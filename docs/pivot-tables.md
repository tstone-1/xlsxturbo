# Pivot tables

`pivot_tables` adds native Excel pivot tables. Each one summarises the frame it is given
with: its source is the frame's written range, header row and formula columns included,
so every field is one of the frame's column names.

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'Region': ['East', 'West', 'East', 'West'],
    'Item':   ['Apple', 'Pear', 'Pear', 'Apple'],
    'Volume': [9000, 5000, 3000, 7000],
    'Month':  ['July', 'July', 'August', 'August'],
})

xlsxturbo.df_to_xlsx(df, "sales.xlsx", sheet_name="Data",
    pivot_tables={
        'Summary!A3': {
            'rows': ['Region'],
            'columns': ['Month'],
            'values': [{'field': 'Volume', 'function': 'sum', 'num_format': '#,##0'}],
        }
    }
)
```

## Where a pivot table goes

The key is the pivot table's first cell:

- `'F3'` places it on the frame's own sheet.
- `'Summary!A3'` places it on sheet `Summary`. When the workbook has a sheet of that name
  (compared without case, as Excel does), the pivot table goes there; otherwise an empty
  sheet is created for it after every other sheet. A name that needs quotes in Excel is
  written the same way here: `"'Q1 summary'!A3"`. A sheet name is taken as written; an
  unquoted one with spaces before or after it is refused, so a sheet whose name really
  starts or ends with a space is given in quotes: `"' Data '!A3"`.

A first cell inside the header or data of any frame in the workbook is refused, and so is a
second pivot table at the same cell.

Excel decides how large a pivot table is when it builds it, so its size is not known when
the file is written. Leave room to the right of and below the first cell: anything there,
such as another frame, a table or a second pivot table, makes Excel show `#SPILL!` in
place of the pivot table. A sheet of its own is the safe choice.

## Fields

| Key | Value |
|---|---|
| `rows` | Column names for the row fields, outermost first |
| `columns` | Column names for the column fields, outermost first |
| `filters` | Column names for the filter fields, shown above the table |
| `values` | Column names, which are summed, or dicts (below) |
| `style` | `'Light16'` (the default) and the rest of Excel's pivot table gallery, `'Light1'` to `'Light28'`, `'Medium1'` to `'Medium28'`, `'Dark1'` to `'Dark28'`; `None` for no style |
| `layout` | `'compact'` (the default), `'outline'` or `'tabular'` |
| `row_grand_totals` | `False` hides the grand total of each row (Excel's "Grand Total" column) |
| `column_grand_totals` | `False` hides the grand total of each column (Excel's "Grand Total" row) |

A value field dict takes `field` (a column name, required), `function` (`'sum'`, `'count'`,
`'average'`, `'max'`, `'min'`, `'product'`, `'count_numbers'`, `'std_dev'`, `'std_dev_p'`,
`'var'`, `'var_p'`; default `'sum'`), `name` (its caption; default Excel's, such as
`Sum of Volume`) and `num_format`. Two value fields may not be given the same `name`;
captions Excel makes itself are not checked. A `name` may not equal a column name of the
frame either, compared without case, as Excel does: `{'field': 'Volume', 'name': 'item'}`
beside a column `Item` raises `ConfigurationError`. A frame column that happens to be
named like a default caption, such as `Sum of Volume` beside a `Volume` value field
without a `name`, is not refused; give that value field a `name` to keep the two apart.

A field may be used once across `rows`, `columns` and `filters`; it may also be a value
field. Field names are matched exactly, case included. A name that is not a column of the
frame raises `ConfigurationError` listing the columns that exist. A column whose name is
not a string is given by its text: a column `2024` is the field `'2024'`.

## What the frame needs

- A header row: `header=False` is refused, because the header texts are the field names.
- At least one column and at least one data row.
- Column names that are unique without regard to case, and not empty.
- A header row that keeps its texts: `cells`, `hyperlinks`, `merged_ranges`, `rich_text`
  or `checkboxes` writing into a header cell of the frame raises `ConfigurationError`
  naming both options and the cell, before anything is written. Writes anywhere else,
  data cells included, are allowed.

## When the table appears

The file holds the pivot table's layout, not its data: Excel builds the table from the
source cells when the workbook is opened. Two things follow:

- An application that does not refresh pivot tables on opening shows an empty one. So
  does Excel for a pivot table on a protected sheet.
- Excel asks whether to save changes when the workbook is closed, although nothing was
  edited.

## Several frames and several sheets

In `dfs_to_xlsx`, `pivot_tables` belongs to the entry that gives it. A workbook-wide
`pivot_tables` applies to the first frame of each sheet, not to a second frame on the same
sheet, in the same way as `charts`. So a workbook-wide key works only for the frame's own
sheet (`'F3'`): a key that names a sheet (`'Summary!A3'`) places every sheet's pivot table
at that one cell, and the second sheet is refused for it. A sheet's options dict switches
the workbook-wide `pivot_tables` off with `{}`; `None` there keeps it, as for every other
option. Pivot tables from several frames may share one sheet:

```python
xlsxturbo.dfs_to_xlsx([
    (sales, "Sales", {"pivot_tables": {"Summary!A3": {"rows": ["Region"], "values": ["Volume"]}}}),
    (costs, "Costs", {"pivot_tables": {"Summary!H3": {"rows": ["Region"], "values": ["Cost"]}}}),
], "report.xlsx")
```

## Constant memory mode

Pivot tables work under `constant_memory=True`. A constant-memory sheet does not keep its
header row once it is written, so xlsxturbo hands the column names to the pivot table
itself.
