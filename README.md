# xlsxturbo

High-performance Excel writer with automatic type detection. Written in Rust, usable from Python.

## Features

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
- **~6x faster** than pandas + openpyxl (see [benchmarks](#performance))
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

### Column Formatting

Apply formatting to data columns using pattern matching:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'product_id': [1, 2, 3],
    'product_name': ['Widget A', 'Widget B', 'Widget C'],
    'price_usd': [19.99, 29.99, 39.99],
    'price_eur': [17.99, 26.99, 35.99],
    'quantity': [100, 75, 50]
})

# Format columns by pattern
xlsxturbo.df_to_xlsx(df, "report.xlsx", column_formats={
    'price_*': {'num_format': '$#,##0.00', 'bg_color': '#E8F5E9'},  # All price columns
    'quantity': {'bold': True}  # Exact match
})

# Wildcard patterns:
# - 'prefix*' matches columns starting with 'prefix'
# - '*suffix' matches columns ending with 'suffix'
# - '*contains*' matches columns containing 'contains'
# - 'exact' matches column name exactly

# Available format options:
# - bg_color (str): Background color ('#RRGGBB' or named)
# - font_color (str): Text color
# - num_format (str): Excel number format ('0.00', '#,##0', '0.00%', etc.)
# - bold (bool): Bold text
# - italic (bool): Italic text
# - underline (bool): Underlined text
# - border (bool): Add thin border

# First matching pattern wins (order preserved)
xlsxturbo.df_to_xlsx(df, "report.xlsx", column_formats={
    'price_usd': {'bg_color': '#FFEB3B'},  # Specific: yellow for USD
    'price_*': {'bg_color': '#E3F2FD'}     # General: blue for other prices
})
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
- `column_formats` (dict): Column formatting with pattern matching
- `conditional_formats` (dict): Conditional formatting (color scales, data bars, icons)
- `formula_columns` (dict): Calculated columns with Excel formulas (column name -> formula template)
- `merged_ranges` (list): List of (range, text) or (range, text, format) tuples to merge cells
- `hyperlinks` (list): List of (cell, url) or (cell, url, display_text) tuples to add clickable links
- `comments` (dict): Cell comments/notes (cell_ref -> text or {text, author})
- `validations` (dict): Data validation rules (column name/pattern -> validation config)
- `rich_text` (dict): Rich text with multiple formats (cell_ref -> list of segments)
- `images` (dict): Embedded images (cell_ref -> path or {path, scale_width, scale_height, alt_text})

### Conditional Formatting

Apply visual formatting based on cell values:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'name': ['Alice', 'Bob', 'Charlie', 'Diana'],
    'score': [95, 72, 88, 45],
    'progress': [0.9, 0.5, 0.75, 0.3],
    'status': [3, 2, 3, 1]
})

xlsxturbo.df_to_xlsx(df, "report.xlsx",
    autofit=True,
    conditional_formats={
        # 2-color gradient: red (low) to green (high)
        'score': {
            'type': '2_color_scale',
            'min_color': '#FF6B6B',
            'max_color': '#51CF66'
        },
        # Data bars: in-cell bar chart
        'progress': {
            'type': 'data_bar',
            'bar_color': '#339AF0',
            'solid': True  # Solid fill instead of gradient
        },
        # Icon set: traffic lights
        'status': {
            'type': 'icon_set',
            'icon_type': '3_traffic_lights'
        }
    }
)
```

**Supported conditional format types:**

| Type | Options |
|------|---------|
| `2_color_scale` | `min_color`, `max_color` |
| `3_color_scale` | `min_color`, `mid_color`, `max_color` |
| `data_bar` | `bar_color`, `border_color`, `solid`, `direction` |
| `icon_set` | `icon_type`, `reverse`, `icons_only` |

**Available icon types:**
- 3 icons: `3_arrows`, `3_arrows_gray`, `3_flags`, `3_traffic_lights`, `3_traffic_lights_rimmed`, `3_signs`, `3_symbols`, `3_symbols_uncircled`
- 4 icons: `4_arrows`, `4_arrows_gray`, `4_traffic_lights`, `4_rating`
- 5 icons: `5_arrows`, `5_arrows_gray`, `5_quarters`, `5_rating`

Column patterns work with conditional formats:
```python
# Apply data bars to all columns starting with "price_"
conditional_formats={'price_*': {'type': 'data_bar', 'bar_color': '#9B59B6'}}
```

### Formula Columns

Add calculated columns to your Excel output. Formulas are written after data columns and use `{row}` as a placeholder for the row number:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'price': [100, 200, 150],
    'quantity': [5, 3, 8],
    'tax_rate': [0.1, 0.1, 0.2]
})

xlsxturbo.df_to_xlsx(df, "sales.xlsx",
    autofit=True,
    formula_columns={
        'Subtotal': '=A{row}*B{row}',      # price * quantity
        'Tax': '=D{row}*C{row}',            # subtotal * tax_rate
        'Total': '=D{row}+E{row}'           # subtotal + tax
    }
)
```

Formula columns appear after data columns (A=price, B=quantity, C=tax_rate, D=Subtotal, E=Tax, F=Total).

**Notes:**
- `{row}` is replaced with the Excel row number (1-based, starting at 2 for data rows when header=True)
- Formula columns inherit header formatting if specified
- Column order is preserved (first formula = first new column)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)

### Merged Cells

Merge cell ranges to create headers, titles, or grouped labels:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'product': ['Widget A', 'Widget B'],
    'sales': [1500, 2300],
    'revenue': [7500, 11500]
})

# Merge cells for a title above the data
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    header=True,
    merged_ranges=[
        # Simple merge with text (auto-centered)
        ('A1:C1', 'Q4 Sales Report'),
        # Merge with custom formatting
        ('A2:C2', 'Regional Data', {
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white'
        })
    ]
)
```

**Merged range format:**
- Tuple of `(range, text)` or `(range, text, format_dict)`
- Range uses Excel notation: `'A1:D1'`, `'B3:B10'`, etc.
- Format options same as `header_format`: bold, italic, font_color, bg_color, font_size, underline

**Notes:**
- Merged cells are applied after data is written, so plan row positions accordingly
- When using with `header=True`, data starts at row 2 (Excel row 2)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)

### Hyperlinks

Add clickable links to cells:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'company': ['Anthropic', 'Google', 'Microsoft'],
    'product': ['Claude', 'Gemini', 'Copilot'],
})

# Add hyperlinks to a new column (D) after the data columns (A, B, C with header)
xlsxturbo.df_to_xlsx(df, "companies.xlsx",
    autofit=True,
    hyperlinks=[
        # Header for the links column
        ('C1', 'https://example.com', 'Website'),
        # Links with company names as display text
        ('C2', 'https://anthropic.com', 'anthropic.com'),
        ('C3', 'https://google.com', 'google.com'),
        ('C4', 'https://microsoft.com', 'microsoft.com'),
    ]
)
```

**Hyperlink format:**
- Tuple of `(cell, url)` or `(cell, url, display_text)`
- Cell uses Excel notation: `'A1'`, `'B5'`, etc.
- Display text is optional; if omitted, the URL is shown

**Notes:**
- Hyperlinks write to the specified cell position (overwrites existing content)
- To add a "links column", target cells beyond your DataFrame columns (as shown above)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

### Comments/Notes

Add cell annotations (hover to view):

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'product': ['Widget A', 'Widget B'],
    'price': [19.99, 29.99]
})

xlsxturbo.df_to_xlsx(df, "report.xlsx",
    comments={
        # Simple text comment
        'A1': 'This column contains product names',
        # Comment with author
        'B1': {'text': 'Prices in USD', 'author': 'Finance Team'}
    }
)
```

**Comment format:**
- Simple: `{'A1': 'Note text'}`
- With author: `{'A1': {'text': 'Note text', 'author': 'Name'}}`

**Notes:**
- Comments appear as small red triangles in the cell corner
- Hover over the cell to see the comment
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

### Data Validation

Add dropdowns and input constraints:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'status': ['Open', 'Closed'],
    'score': [85, 92],
    'price': [19.99, 29.99],
    'code': ['ABC', 'XYZ']
})

xlsxturbo.df_to_xlsx(df, "validated.xlsx",
    validations={
        # Dropdown list
        'status': {
            'type': 'list',
            'values': ['Open', 'Closed', 'Pending', 'Review']
        },
        # Whole number range (0-100)
        'score': {
            'type': 'whole_number',
            'min': 0,
            'max': 100,
            'error_title': 'Invalid Score',
            'error_message': 'Score must be between 0 and 100'
        },
        # Decimal range
        'price': {
            'type': 'decimal',
            'min': 0.0,
            'max': 999.99
        },
        # Text length constraint
        'code': {
            'type': 'text_length',
            'min': 3,
            'max': 10
        }
    }
)
```

**Validation types:**

| Type | Aliases | Description | Options |
|------|---------|-------------|---------|
| `list` | - | Dropdown menu | `values` (list of strings, max 255 chars total) |
| `whole_number` | `whole`, `integer` | Integer range | `min`, `max` |
| `decimal` | `number` | Decimal range | `min`, `max` |
| `text_length` | `textlength`, `length` | Character count | `min`, `max` |

**Optional message options:**
- `input_title`, `input_message`: Prompt shown when cell is selected
- `error_title`, `error_message`: Message shown when invalid data is entered

**Notes:**
- Validations apply to the data rows of the specified column
- Column patterns work: `'score_*': {...}` matches all columns starting with `score_`
- If only `min` or only `max` is specified, the other defaults to the type's extreme value
- List validation values are limited to 255 total characters (Excel limitation)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

### Rich Text

Multiple formats within a single cell:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'A': [1, 2, 3]})

xlsxturbo.df_to_xlsx(df, "rich.xlsx",
    rich_text={
        'D1': [
            ('Important: ', {'bold': True, 'font_color': 'red'}),
            'Please review ',
            ('all', {'italic': True}),
            ' values'
        ],
        'D2': [
            ('Status: ', {'bold': True}),
            ('OK', {'font_color': 'green', 'bold': True})
        ]
    }
)
```

**Segment format:**
- Formatted: `('text', {'bold': True, 'font_color': 'blue'})`
- Plain: `'plain text'` (no formatting)

**Available format options:**
- `bold` (bool)
- `italic` (bool)
- `font_color` (str): '#RRGGBB' or named color
- `bg_color` (str): Background color
- `font_size` (float)
- `underline` (bool)

**Notes:**
- Rich text writes to the specified cell position (overwrites existing content)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

### Images

Embed images in cells:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'Product': ['Widget A', 'Widget B'], 'Price': [19.99, 29.99]})

xlsxturbo.df_to_xlsx(df, "catalog.xlsx",
    autofit=True,
    images={
        # Simple path
        'C2': 'images/widget_a.png',
        # With options
        'C3': {
            'path': 'images/widget_b.png',
            'scale_width': 0.5,
            'scale_height': 0.5,
            'alt_text': 'Widget B photo'
        }
    }
)
```

**Image format:**
- Simple: `{'C2': 'path/to/image.png'}`
- With options: `{'C2': {'path': '...', 'scale_width': 0.5, ...}}`

**Available options:**
- `path` (str, required): Path to image file
- `scale_width` (float): Width scale factor (1.0 = original)
- `scale_height` (float): Height scale factor (1.0 = original)
- `alt_text` (str): Alternative text for accessibility

**Supported formats:** PNG, JPEG, GIF, BMP

**Notes:**
- Images are positioned at the specified cell (overlays any existing content)
- Image file must exist; non-existent files will raise an error
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

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
- `conditional_formats`
- `merged_ranges`
- `hyperlinks`
- `comments`
- `validations`
- `rich_text`
- `images`
- `autofit`
- `formula_columns`

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

# Handle ambiguous dates (01-02-2024: is it Jan 2 or Feb 1?)
xlsxturbo.csv_to_xlsx("us_data.csv", "output.xlsx", date_order="us")   # January 2
xlsxturbo.csv_to_xlsx("eu_data.csv", "output.xlsx", date_order="eu")   # February 1

# date_order options:
# - "auto" (default): ISO first, then European (DMY), then US (MDY)
# - "mdy" or "us": US format (MM-DD-YYYY)
# - "dmy" or "eu": European format (DD-MM-YYYY)
```

## CLI Usage

```bash
xlsxturbo input.csv output.xlsx [OPTIONS]
```

### Options

- `-s, --sheet-name <NAME>`: Name of the Excel sheet (default: "Sheet1")
- `-d, --date-order <ORDER>`: Date parsing order for ambiguous dates (default: "auto")
  - `auto`: ISO first, then European, then US
  - `mdy` or `us`: US format (01-02-2024 = January 2)
  - `dmy` or `eu`: European format (01-02-2024 = February 1)
- `-v, --verbose`: Show progress information

### Examples

```bash
# Basic conversion
xlsxturbo sales.csv report.xlsx

# With US date format
xlsxturbo sales.csv report.xlsx --date-order us

# With European date format and verbose output
xlsxturbo sales.csv report.xlsx -d eu -v --sheet-name "Q4 Sales"
```

## Performance

*Reference benchmark on 100,000 rows x 50 columns with mixed data types. Your results will vary by system - run the benchmark yourself (see [Benchmarking](#benchmarking)).*

| Library | Time (s) | Rows/sec | vs xlsxturbo |
|---------|----------|----------|--------------|
| **xlsxturbo** | **6.65** | **15,033** | **1.0x** |
| polars | 25.07 | 3,988 | 3.8x |
| pandas + xlsxwriter | 35.60 | 2,809 | 5.4x |
| pandas + openpyxl | 38.85 | 2,574 | 5.8x |

*Test system: Windows 11, Python 3.14, AMD Ryzen 9 (32 threads)*

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

Run the included benchmark scripts:

```bash
# Compare xlsxturbo vs other libraries (100K rows default)
python benchmarks/benchmark.py

# Full benchmark: small, medium, large datasets
python benchmarks/benchmark.py --full

# Custom size
python benchmarks/benchmark.py --rows 500000 --cols 100

# Output formats for CI/documentation
python benchmarks/benchmark.py --markdown
python benchmarks/benchmark.py --json

# Test parallel vs single-threaded CSV conversion
python benchmarks/benchmark_parallel.py
```

## License

MIT
