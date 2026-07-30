# xlsxturbo

High-performance Excel writer with automatic type detection. Written in Rust,
usable from Python.

xlsxturbo exports pandas and polars DataFrames and CSV files to `.xlsx`, using Rust
for the hot path while keeping the Python API small enough to drop straight into a
script, a report job, or a batch pipeline. It is roughly 7-9x faster than
pandas + openpyxl on the [reference benchmarks](performance.md), and it supports the
Excel features those exports usually need -- tables, conditional formatting, charts,
data validation, images -- through focused keyword arguments rather than a workbook
object model.

## Install

```bash
pip install xlsxturbo
```

Wheels are published for Python 3.10+ on Linux, Windows, and macOS. There are no
runtime dependencies beyond the interpreter itself.

## Your first export

```python
import pandas as pd
from xlsxturbo import df_to_xlsx

df = pd.DataFrame({
    "product": ["Widget", "Gadget", "Gizmo"],
    "price": [19.99, 34.50, 8.75],
    "in_stock": [True, False, True],
    "restock": pd.to_datetime(["2024-03-01", "2024-03-15", "2024-04-01"]),
})

df_to_xlsx(df, "products.xlsx", table_style="Medium2", autofit=True)
```

Types carry across without configuration: integers and floats become Excel numbers,
booleans become Excel booleans, and dates and datetimes become real Excel date values
with a display format attached. See [DataFrame export](dataframe-export.md).

Converting a CSV is a single call, with the types detected from the file's text:

```python
from xlsxturbo import csv_to_xlsx

csv_to_xlsx("sales.csv", "sales.xlsx")
```

See [CSV conversion](csv-conversion.md), which also covers the `xlsxturbo` command-line
tool.

## Where to go next

- **[Capability matrix](capability-matrix.md)** -- which options each function accepts,
  which can be overridden per sheet, and which survive `constant_memory` mode. Generated
  from the source, so it cannot drift.
- **[API reference](api-reference.md)** -- the three entry points and their arguments.
- **[Errors](errors.md)** -- what gets raised, when, and what the file on disk looks
  like afterwards.
- **[Stability and support](stability.md)** -- what 1.0 promises, how long a deprecation
  lasts, and which Pythons and platforms are supported.
- **[Compatibility](compatibility.md)** -- known limitations and the parts of Excel's
  data model that do not round-trip.

## Feature overview

- **Direct DataFrame support** for pandas and polars
- **Excel tables** - filterable tables with 61 built-in styles (banded rows, autofilter)
- **Conditional formatting** - color scales, data bars, icon sets for visual data analysis
- **Formula columns** - add calculated columns with Excel formulas
- **Merged cells** - merge cell ranges for headers and titles
- **Hyperlinks** - add clickable links to cells
- **Comments/Notes** - add cell annotations with optional author
- **Data validation** - dropdowns, number ranges, text length constraints
- **Rich text** - multiple formats within a single cell
- **Images** - embed PNG, JPEG, GIF, BMP in cells
- **Checkboxes** - interactive cell checkboxes (Excel for Microsoft 365, Sept 2024+)
- **Textboxes** - floating text shapes with configurable font, fill, and line colors
- **Native Excel charts** - editable bar, column, line, pie, scatter, and other chart types
- **Sparklines** - mini in-cell line, column, and win/loss charts for inline trends
- **Defined names** - workbook-level named ranges for formulas and references
- **Arbitrary cell writes** - write values to specific cells with optional formatting
- **Border styles** - per-side borders (left, right, top, bottom) with 13 style options
- **Text alignment** - horizontal and vertical alignment with text wrapping
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
- **~7-9x faster** than pandas + openpyxl on reference systems (see [benchmarks](performance.md))
- **Memory efficient** - streams data with 1MB buffer
- Available as a **Python library**, plus a **CLI tool** that has to be
  [built from source](csv-conversion.md#cli-usage) — it is not in the PyPI wheel
