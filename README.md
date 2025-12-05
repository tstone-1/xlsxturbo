# xlsxturbo

High-performance Excel writer with automatic type detection. Written in Rust, usable from Python.

## Features

- **Direct DataFrame support** for pandas and polars
- **Multi-sheet workbooks** - write multiple DataFrames to one file
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
```

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
