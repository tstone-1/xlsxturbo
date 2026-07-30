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

## Type Detection Examples

| CSV Value | Excel Type | Notes |
|-----------|------------|-------|
| `123` | Number | Integer |
| `3.14159` | Number | Float |
| `true` / `FALSE` | Boolean | Case insensitive |
| `2024-01-15` | Date | Formatted as date |
| `2024-01-15T10:30:00` | DateTime | ISO 8601 format |
| `NaN` | Empty | Graceful handling |
| `hello world` | Text | Default |

Supported date formats: `YYYY-MM-DD`, `YYYY/MM/DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM-DD-YYYY`, `MM/DD/YYYY`

DataFrame columns follow the same mapping, with one addition worth knowing: **durations
(`pandas.Timedelta` / `numpy.timedelta64`) are written as text**, e.g. `86400 seconds`.
Excel has no duration type — convert to a number in the unit you want first
(`df["elapsed"].dt.total_seconds()`) and apply a `num_format` such as `[h]:mm:ss` if you
want it displayed as a duration.
