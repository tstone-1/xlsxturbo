# Migrating from pandas, XlsxWriter, openpyxl or polars

xlsxturbo replaces the *write* step of an export. Your DataFrame code stays as it is; the
`to_excel` or `write_excel` call becomes a `df_to_xlsx` call. This page shows the common
exports side by side and lists every place the output differs, so you can decide before
switching whether a given export migrates cleanly.

Every example on this page runs in the test suite, and each "after" workbook is compared
cell by cell with its "before" workbook.

## What does not migrate

Check these first. If an export depends on one, keep it on its current writer.

- **Reading or editing an existing workbook.** xlsxturbo only writes new files:
  `openpyxl.load_workbook(...)`, `pd.ExcelWriter(..., mode="a")` and template-filling have
  no equivalent. See [Compatibility](compatibility.md).
- **Placing a frame at an offset**, with `startrow` / `startcol`, or several frames on one
  sheet. Each DataFrame fills its own sheet from cell `A1`. Single values can go anywhere
  with [`cells`](cells.md).
- **Formats other than `.xlsx`**: no `.xls`, no macro-enabled `.xlsm`.

## The data these examples use

```python
import datetime as dt

import pandas as pd
import polars as pl
import xlsxturbo

sales = pd.DataFrame(
    {
        "region": ["North", "South", "East"],
        "units": [120, 85, 240],
        "revenue": [15000.5, 9200.0, 31000.25],
        "closed": [dt.date(2026, 9, 1), dt.date(2026, 9, 2), dt.date(2026, 9, 3)],
    }
)
```

## From `pandas.DataFrame.to_excel`

### A plain export

```python
# Before
sales.to_excel("before_plain.xlsx", sheet_name="Sales", index=False)
```

```python
# After
xlsxturbo.df_to_xlsx(sales, "after_plain.xlsx", sheet_name="Sales")
```

xlsxturbo never writes the index, so `index=False` has no counterpart. If you relied on
the default `index=True`, see [the index](#the-index) below.

### Several sheets in one workbook

```python
# Before
with pd.ExcelWriter("before_sheets.xlsx") as writer:
    sales.to_excel(writer, sheet_name="Sales", index=False)
    sales.describe().to_excel(writer, sheet_name="Summary")
```

```python
# After
xlsxturbo.dfs_to_xlsx(
    [
        (sales, "Sales"),
        (sales.describe().reset_index(names=""), "Summary"),
    ],
    "after_sheets.xlsx",
)
```

`reset_index(names="")` turns the index into an ordinary column with an empty header,
which is what `to_excel` writes for an unnamed index.

### The differences to check

| Data | `to_excel` | `df_to_xlsx` | To keep the old result |
|------|------------|--------------|-------------------------|
| The index | Written as the first column(s) unless `index=False` | Never written | `df.reset_index()` first |
| `MultiIndex` columns | Two header rows, merged cells | One header row; each name is the tuple as text, `('g', 'a')` | Flatten first: `df.columns = ["_".join(c) for c in df.columns]` |
| Timezone-aware datetimes | Raises `ValueError` | Written as local wall-clock time, offset dropped | Convert explicitly — see [Compatibility](compatibility.md) |
| Integers above 2^53 | Rounded to the nearest float | Written as text, every digit kept | Nothing: this is the safer behaviour |
| `Timedelta` | A number of days, e.g. `0.0625` | Text, e.g. `5400000000 microseconds` | `df["t"] / pd.Timedelta(days=1)` and a `[h]:mm` number format |
| Non-string column names (`0`, `1`) | Numbers in the header | Text in the header | Usually nothing |
| Floats, with the default openpyxl engine | Rounded to 16 significant digits: `0.1 + 0.2` is stored as `0.3` | The exact value, `0.30000000000000004` | Nothing: the difference is in the last bit, and xlsxturbo's is the value you had |

Missing values (`NaN`, `None`, `pd.NA`, `NaT`) become empty cells in both. Dates and
datetimes get the same `yyyy-mm-dd` and `yyyy-mm-dd hh:mm:ss` formats. Neither writer
styles the header row.

### The index

```python
# Before
sales.set_index("region").to_excel("before_index.xlsx")
```

```python
# After
xlsxturbo.df_to_xlsx(sales.set_index("region").reset_index(), "after_index.xlsx")
```

## From pandas with XlsxWriter formatting

The usual reason to reach for `engine="xlsxwriter"` is formatting that `to_excel` cannot
express. In xlsxturbo that formatting is keyword arguments to the same call.

```python
# Before
with pd.ExcelWriter("before_formatted.xlsx", engine="xlsxwriter") as writer:
    sales.to_excel(writer, sheet_name="Sales", index=False)
    book, sheet = writer.book, writer.sheets["Sales"]
    money = book.add_format({"num_format": "#,##0.00"})
    header = book.add_format({"bold": True, "bg_color": "#DDEBF7"})
    for col, name in enumerate(sales.columns):
        sheet.write(0, col, name, header)
    sheet.set_column("C:C", 14, money)
    sheet.freeze_panes(1, 0)
```

```python
# After
xlsxturbo.df_to_xlsx(
    sales,
    "after_formatted.xlsx",
    sheet_name="Sales",
    header_format={"bold": True, "bg_color": "#DDEBF7"},
    column_formats={"revenue": {"num_format": "#,##0.00"}},
    column_widths={2: 14},
    freeze_panes=True,
)
```

Formats are keyed by column name or wildcard pattern, widths by column index, rather than
by letter. The
full set of options is on [Formatting](formatting.md) and in the
[capability matrix](capability-matrix.md). Tables, conditional formats, charts and data
validation are keyword arguments as well; see their pages in the Guide.

## From `polars.DataFrame.write_excel`

The examples use `pl.from_pandas(sales)`, which needs `pyarrow` for the Python date
objects in this sample frame. Install it alongside polars when running these examples.

`write_excel` makes more decisions for you: it wraps the frame in an Excel table with no
style (an autofilter, no colours) and gives numeric columns thousands separators with red
negatives. `df_to_xlsx` writes plain cells unless asked, so state the ones you want.
`table_style="None"` is the unstyled table; any named style such as `"Medium9"` works too.

```python
# Before
pl.from_pandas(sales).write_excel("before_polars.xlsx", worksheet="Sales")
```

```python
# After
xlsxturbo.df_to_xlsx(
    pl.from_pandas(sales),
    "after_polars.xlsx",
    sheet_name="Sales",
    table_style="None",
    column_formats={
        "units": {"num_format": "#,##0;[Red]-#,##0"},
        "revenue": {"num_format": "#,##0.000;[Red]-#,##0.000"},
    },
)
```

Both accept a polars frame directly; there is no conversion to pandas. A polars
`Duration` is written as text such as `1:30:00`, where `write_excel` writes a number of
days.

## From openpyxl, row by row

Code that builds a workbook with `ws.append(...)` is usually exporting a table that
already exists as rows. Put the rows in a DataFrame and write that.

```python
# Before
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(list(sales.columns))
for row in sales.itertuples(index=False):
    ws.append(list(row))
wb.save("before_openpyxl.xlsx")
```

```python
# After
xlsxturbo.df_to_xlsx(sales, "after_openpyxl.xlsx", sheet_name="Sales")
```

If the rows come from a database cursor or a list of dicts, `pd.DataFrame(rows)` or
`pl.DataFrame(rows)` is the only extra step. Code that edits cells of a loaded workbook,
reads values back, or styles individual cells after the fact does not migrate; see
[What does not migrate](#what-does-not-migrate).

## Checking a migrated export

Compare the old and new workbooks by value before switching over. This is the check the
test suite runs on every example above:

```python
from openpyxl import load_workbook


def sheet_values(path):
    """Every sheet's cell values, keyed by sheet name."""
    book = load_workbook(path)
    return {ws.title: [list(row) for row in ws.iter_rows(values_only=True)] for ws in book}


assert sheet_values("before_plain.xlsx") == sheet_values("after_plain.xlsx")
```

A float that pandas wrote through openpyxl can differ from xlsxturbo's in the last bit (see
the table above), so compare computed floats with `math.isclose` rather than `==`.

The performance gain depends on the shape of the data; measure your own export with
[the benchmark scripts](performance.md#benchmarking) rather than relying on the headline
numbers.
