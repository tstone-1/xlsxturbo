# Formatting

Cell appearance is controlled by three groups of options: the header row
(`header_format`), the data columns (`column_formats`, matched by name or pattern), and
the sheet geometry (widths, heights, merged ranges). Format dictionaries reject unknown
keys rather than ignoring them, so a typo is an error and not a silently missing
style.

## Header Styling

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
# - border (bool|str): True = thin all sides, or a style name
# - border_left/right/top/bottom (bool|str): Per-side border, same values
# - border_color (str): Color for all borders
# - align_horizontal (str): 'left', 'center', 'right', 'fill', 'justify',
#     'center_across', 'distributed'
# - align_vertical (str): 'top', 'center', 'bottom', 'justify', 'distributed'
# - wrap_text (bool): Enable text wrapping within cell
```

> **Note:** Unknown keys (e.g. `'color'` instead of `'font_color'`) and wrong value types raise an error listing the valid options. Applies to `header_format`, `column_formats`, `conditional_formats[...]['format']`, `images`, `validations`, `textboxes`, `charts`, `sparklines`, and `rich_text` segment formats.
>
> `rich_text` segments accept **font-level keys only** (`bold`, `italic`, `underline`, `font_color`, `bg_color`, `font_size`). A segment is an inline run inside one cell, so cell-level keys — borders, `align_horizontal`/`align_vertical`, `wrap_text` — would never render and are rejected rather than silently ignored. Format the cell itself via `column_formats` or `cells` instead.
>
> Column patterns in `column_formats`, `conditional_formats`, and `validations` must match at least one DataFrame column. A zero-match exact name or wildcard raises `ValueError` instead of silently omitting the requested behavior.

## Column Formatting

Apply formatting to data columns using pattern matching. Unknown keys raise errors (see [Header Styling](#header-styling)).

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
# - border (bool|str): True = thin all sides, or a style name for all sides
# - border_left (bool|str): True = thin, or a style name, left side only
# - border_right (bool|str): True = thin, or a style name, right side only
# - border_top (bool|str): True = thin, or a style name, top side only
# - border_bottom (bool|str): True = thin, or a style name, bottom side only
# - border_color (str): Color for all borders ('#RRGGBB' or named)
#
# Every border key takes either a bool or a style name. True means thin;
# False means no border, which is the same as omitting the key.
# Border styles: thin, medium, thick, dashed, dotted, double, hair,
#   medium_dashed, dash_dot, medium_dash_dot, dash_dot_dot,
#   medium_dash_dot_dot, slant_dash_dot
# - align_horizontal (str): 'left', 'center', 'right', 'fill', 'justify',
#     'center_across', 'distributed'
# - align_vertical (str): 'top', 'center', 'bottom', 'justify', 'distributed'
# - wrap_text (bool): Enable text wrapping within cell

# First matching pattern wins (order preserved)
xlsxturbo.df_to_xlsx(df, "report.xlsx", column_formats={
    'price_usd': {'bg_color': '#FFEB3B'},  # Specific: yellow for USD
    'price_*': {'bg_color': '#E3F2FD'}     # General: blue for other prices
})

# Per-side borders with style control
xlsxturbo.df_to_xlsx(df, "report.xlsx", column_formats={
    'price_usd': {'border_right': 'thick'},              # Thick right border only
    'quantity': {'border': 'thin'},                       # Thin border all sides
    'product_name': {'border_left': 'medium', 'border_right': 'medium'},  # Left+right
})
```

## Custom Column Widths and Row Heights

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

### Excel's own limits

Excel caps a column at 255 characters and a row at 409.5 points. A larger value
is accepted rather than rejected, and nothing is lost by asking for one: widths
above the cap are already clamped as the file is written (255, 300 and
1,000,000 all produce the same stored width), and Excel clamps an over-limit
row height when it loads the workbook. Measured on Excel for Mac 16: a workbook
written with `row_heights={0: 500}` or `{0: 1e6}` opens with no repair prompt
and shows a height of 409.5, and one written with `column_widths={0: 1e6}`
shows the same width as `{0: 255}`. So the values below the cap are the ones
worth writing, but an over-limit one costs only its own precision.

## Global Column Width Cap

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

When `autofit=True` is combined with `column_widths` that names specific columns and has no `'_all'` key, the named columns get their explicit widths and every other column is autofitted to its content. Add `'_all'` back in to cap the autofitted columns instead of leaving them uncapped.

## Merged Cells

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
- Format options are the same set `header_format` takes -- the font keys (bold, italic,
  underline, font_color, bg_color, font_size) **and** the cell keys (border and the four
  per-side border keys, border_color, align_horizontal, align_vertical, wrap_text). Only
  `num_format` is out of scope; that one is accepted by `column_formats` and the nested
  `format` keys.

**Notes:**
- Merged cells are applied after data is written, so plan row positions accordingly
- When using with `header=True`, data starts at row 2 (Excel row 2)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)

## Rich Text

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
