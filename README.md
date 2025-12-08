# xlsxturbo

High-performance Excel writer with automatic type detection. Written in Rust, usable from Python.

## Features

- **Direct DataFrame support** for pandas and polars
- **Excel tables** - filterable tables with 61 built-in styles (banded rows, autofilter)
- **Auto-fit columns** - automatically adjust column widths to fit content
- **Custom column widths** - set specific widths per column or cap all with _all
- **Header styling** - bold, colors, font size for header row
- **Named tables** - set custom table names
- **Custom row heights** - set specific heights per row
- **Freeze panes** - freeze header row for easier scrolling
- **Multi-sheet workbooks** - write multiple DataFrames to one file
- **Per-sheet options** - override settings per sheet in multi-sheet workbooks
- **Constant memory mode** - minimize RAM usage for very large files
- **Parallel CSV processing** - optional multi-core parsing for large files
- **Automatic type detection** from CSV strings and Python objects:
  - Integers and floats → Excel numbers
  - `true`/`false` → Excel booleans
  - Dates (`2024-01-15`, `15/01/2024`, etc.) → Excel dates with formatting
  - Datetimes (ISO 8601) → Excel datetimes
  - `NaN`/`Inf` → Empty cells (graceful handling)
  - Everything else → Text
- **~25x faster** than pandas + openpyxl
- **Memory efficient** - streams data with 1MB buffer
- Available as both **Python library** and **CLI tool**

## Installation

```bash
pip install xlsxturbo
```

Or build from source:

```bash
pip install maturin
maturin develop --release
```

## Python Usage

### DataFrame Export (pandas/polars)

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

### Excel Tables with Styling

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'Product': ['Widget A', 'Widget B', 'Widget C'],
    'Price': [19.99, 29.99, 39.99],
    'Quantity': [100, 75, 50],
})

# Create a styled Excel table with autofilter, banded rows, and auto-fit columns
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    table_style="Medium9",   # Excel's default table style
    autofit=True,            # Fit column widths to content
    freeze_panes=True        # Freeze header row for scrolling
)

# Available styles: Light1-Light21, Medium1-Medium28, Dark1-Dark11
xlsxturbo.df_to_xlsx(df, "dark_table.xlsx", table_style="Dark1", autofit=True)
```

### Custom Column Widths and Row Heights

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Department': ['Engineering', 'Marketing', 'Sales'],
    'Salary': [75000, 65000, 55000]
})

# Set specific column widths (column index -> width in characters)
xlsxturbo.df_to_xlsx(df, "report.xlsx", 
    column_widths={0: 20, 1: 25, 2: 15}
)

# Set specific row heights (row index -> height in points)
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    row_heights={0: 25}  # Make header row taller
)

# Combine with other options
xlsxturbo.df_to_xlsx(df, "styled.xlsx",
    table_style="Medium9",
    freeze_panes=True,
    column_widths={0: 20, 1: 30, 2: 15},
    row_heights={0: 22}
)
```

### Global Column Width Cap

Use `column_widths={'_all': value}` to cap all columns at a maximum width:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'Name': ['Alice', 'Bob'],
    'VeryLongDescription': ['A' * 100, 'B' * 100],
    'Score': [95, 87]
})

# Cap all columns at 30 characters
xlsxturbo.df_to_xlsx(df, "capped.xlsx", column_widths={'_all': 30})

# Mix specific widths with global cap (specific overrides '_all')
xlsxturbo.df_to_xlsx(df, "mixed.xlsx", column_widths={0: 15, '_all': 30})

# Autofit with cap: fit content, but never exceed 25 characters
xlsxturbo.df_to_xlsx(df, "fitted.xlsx", autofit=True, column_widths={'_all': 25})
```

### Named Excel Tables

Set custom names for Excel tables:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'Product': ['A', 'B'], 'Price': [10, 20]})

# Name the Excel table
xlsxturbo.df_to_xlsx(df, "report.xlsx", 
    table_style="Medium2", 
    table_name="ProductPrices"
)

# Invalid characters are auto-sanitized, digits get underscore prefix
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    table_style="Medium2",
    table_name="2024 Sales Data!"  # Becomes "_2024_Sales_Data_"
)
```

### Header Styling

Apply custom formatting to header cells:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'Name': ['Alice', 'Bob'], 'Score': [95, 87]})

# Bold headers
xlsxturbo.df_to_xlsx(df, "bold.xlsx", header_format={'bold': True})

# Full styling with colors
xlsxturbo.df_to_xlsx(df, "styled.xlsx", header_format={
    'bold': True,
    'bg_color': '#4F81BD',   # Blue background
    'font_color': 'white'    # White text
})

# Available options:
# - bold (bool): Bold text
# - italic (bool): Italic text
# - font_color (str): '#RRGGBB' or named color (white, black, red, blue, etc.)
# - bg_color (str): Background color
# - font_size (float): Font size in points
# - underline (bool): Underlined text
```


### Multi-Sheet Workbooks

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

### Per-Sheet Options

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

Available per-sheet options:
- `header` (bool): Include column names as header row
- `autofit` (bool): Automatically adjust column widths
- `table_style` (str|None): Excel table style or None to disable
- `freeze_panes` (bool): Freeze header row
- `column_widths` (dict): Custom column widths
- `row_heights` (dict): Custom row heights
- `table_name` (str): Custom Excel table name
- `header_format` (dict): Header cell styling

### Constant Memory Mode (Large Files)

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

**Note:** Constant memory mode disables some features that require random access:
- `table_style` (Excel tables)
- `freeze_panes`
- `row_heights`
- `autofit`

Column widths still work in constant memory mode.

### CSV Conversion

```python
import xlsxturbo

# Convert CSV to XLSX with automatic type detection
rows, cols = xlsxturbo.csv_to_xlsx("input.csv", "output.xlsx")
print(f"Converted {rows} rows and {cols} columns")

# Custom sheet name
xlsxturbo.csv_to_xlsx("data.csv", "report.xlsx", sheet_name="Sales Data")

# For large files (100K+ rows), use parallel processing
xlsxturbo.csv_to_xlsx("big_data.csv", "output.xlsx", parallel=True)
```

## CLI Usage

```bash
xlsxturbo input.csv output.xlsx [--sheet-name "Sheet1"] [-v]
```

### Options

- `-s, --sheet-name`: Name of the Excel sheet (default: "Sheet1")
- `-v, --verbose`: Show progress information

### Example

```bash
xlsxturbo sales.csv report.xlsx --sheet-name "Q4 Sales" -v
```

## Performance

Benchmarked on 525,684 rows x 98 columns:

| Method | Time | Speedup |
|--------|------|---------|
| **xlsxturbo** | 28.5s | **26.7x** |
| PyExcelerate | 107s | 7.1x |
| pandas + xlsxwriter | 374s | 2.0x |
| pandas + openpyxl | 762s | 1.0x |
| polars.write_excel | 1039s | 0.7x |

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

## Building from Source

Requires Rust toolchain and maturin:

```bash
# Install maturin
pip install maturin

# Development build
maturin develop

# Release build (optimized)
maturin develop --release

# Build wheel for distribution
maturin build --release
```

## Benchmarking

Run the included benchmark script:

```bash
# Default: 100K rows x 50 columns
python benchmark.py

# Custom size
python benchmark.py --rows 500000 --cols 100
```

## License

MIT
