# DataFrame export

`df_to_xlsx` writes a single pandas or polars DataFrame to one worksheet. It is the
main entry point, and every formatting, table, chart, and validation option in the rest
of these pages is a keyword argument to it.

## DataFrame Export (pandas/polars)

```python
import xlsxturbo
import pandas as pd

# Create a DataFrame
df = pd.DataFrame({
    'name': ['Alice', 'Bob'],
    'age': [30, 25],
    'salary': [50000.50, 60000.75],
    'active': [True, False]
})

# Export to XLSX (preserves types: int, float, bool, date, datetime)
rows, cols = xlsxturbo.df_to_xlsx(df, "output.xlsx")
print(f"Wrote {rows} rows and {cols} columns")

# Works with polars too!
import polars as pl
df_polars = pl.DataFrame({'x': [1, 2, 3], 'y': [4.0, 5.0, 6.0]})
xlsxturbo.df_to_xlsx(df_polars, "polars_output.xlsx", sheet_name="Data")
```

## Writing to memory instead of a file

`output_path` also accepts a binary file-like object: `io.BytesIO`, a file opened with
`"wb"`, or anything else with a `write(bytes)` method. This is the shape for a web
download or an object-storage upload, where a temporary file is only plumbing.
`dfs_to_xlsx` and `csv_to_xlsx` accept the same targets.

```python
import io

buffer = io.BytesIO()
xlsxturbo.df_to_xlsx(df, buffer, autofit=True)
data = buffer.getvalue()  # the complete .xlsx file
```

In a FastAPI endpoint:

```python
from fastapi import Response

@app.get("/report.xlsx")
def report() -> Response:
    buffer = io.BytesIO()
    xlsxturbo.df_to_xlsx(load_report(), buffer, table_style="Medium9")
    return Response(
        content=buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="report.xlsx"'},
    )
```

What to expect:

- **The whole file is held in memory** before it is written, compressed. With
  `constant_memory=True` the worksheet data is still staged on disk while rows are
  written, but the finished archive is not.
- **Conversion and workbook-save failures write nothing** into the object.
  A failure while delivering the finished archive can leave partial bytes in the
  object; unlike a path export, an arbitrary writer cannot be rolled back.
- **The object stays open, at its new position.** Call `buffer.getvalue()`, or
  `buffer.seek(0)` before handing it to something that reads it.
- **A text stream is refused** with a `ConfigurationTypeError`. An `.xlsx` file is
  binary, so open files with `"wb"`, not `"w"`.
- A `write()` that returns fewer bytes than it was given is called again with the rest;
  one that returns `None` (Django's `HttpResponse`, for example) is taken to have
  accepted everything, except for `io.RawIOBase` streams. Their `None` means the
  stream would block, and raises `FileError` with `errno.EAGAIN`.

## Type Detection Examples

| CSV Value | Excel Type | Notes |
|-----------|------------|-------|
| `123` | Number | Integer |
| `3.14159` | Number | Float |
| `true` / `FALSE` | Boolean | Case insensitive |
| `2024-01-15` | Date | Formatted as date |
| `2024-01-15T10:30:00` | DateTime | ISO 8601 format |
| `NaN` / `inf` / `Infinity` | Empty | Any spelling of a non-finite number — see below |
| `1899-12-30` … `1900-02-28` | Text | Excel numbers these a day out — see below |
| `hello world` | Text | Default |

Supported date formats: `YYYY-MM-DD`, `YYYY/MM/DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM-DD-YYYY`, `MM/DD/YYYY`

### Non-finite numbers, and a trap in text columns

Excel has no NaN or infinity, so a CSV field holding one becomes an empty cell. That covers
**every spelling the number parser accepts**, in any case and with an optional sign: `nan`,
`NaN`, `+nan`, `inf`, `-Inf`, `infinity`, `INFINITY`.

The trap is that some of those are also ordinary words. A text column whose values include
`NAN` or `Inf.` loses them to empty cells, because nothing distinguishes the word from the
number at that point. If a column is meant to be text, say so with a `column_format` — or
check for the spellings above before exporting.

### Dates before 1900-03-01

Excel's serial numbering assumes a 1900-02-29 that never existed, so every date before
1900-03-01 is numbered one day later than the real calendar. Rather than write a date that
renders on the wrong day, xlsxturbo writes those as text, exactly as they appeared in the
input. Dates from 1900-03-01 onward — serial 61 up to Excel's last day, 9999-12-31 — are
unaffected.

### Very large integers

An integer whose magnitude exceeds 2^53 (9,007,199,254,740,992) cannot be stored as an Excel
number without changing value, because Excel stores every number as a double. Those are
written as text so no digits are lost. Anything at or below that magnitude is a normal
number.

DataFrame columns follow the same mapping, with one addition worth knowing: **durations
(`pandas.Timedelta` / `numpy.timedelta64`) are written as text**, e.g. `86400 seconds`.
Excel has no duration type — convert to a number in the unit you want first
(`df["elapsed"].dt.total_seconds()`) and apply a `num_format` such as `[h]:mm:ss` if you
want it displayed as a duration.
