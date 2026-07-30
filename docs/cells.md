# Individual cells

These options address specific cells by reference (`"B7"`) rather than by column, and
they are applied after the DataFrame is written -- so a `cells` entry can deliberately
overwrite a value that came from the data.

## Arbitrary Cell Writes

Write values to specific cells, optionally overwriting DataFrame data:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'product': ['Widget A', 'Widget B'],
    'price': [19.99, 29.99]
})

# Write simple values to specific cells
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    cells={
        'D1': 'Notes',          # String
        'D2': 'Reviewed',       # String
        'D3': 42,               # Number
        'E1': True              # Boolean
    }
)

# Write with number formatting (e.g., force text format for long numbers)
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    cells={
        'C5': 'Total',
        'C6': {'value': '934728173849', 'num_format': '@'},  # Text format
        'C7': {'value': 0.15, 'num_format': '0.00%'}         # Percentage
    }
)

# Overwrite DataFrame cells (cells are written after data)
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    cells={
        'A2': 'OVERRIDE',  # Replaces 'Widget A' in the output
    }
)
```

**Cell value format:**
- Simple: `{'A1': 'text'}`, `{'B2': 42}`, `{'C3': True}`
- With formatting: `{'A1': {'value': '...', 'num_format': '@'}}`
- Additional format options: `align_horizontal`, `align_vertical`, `wrap_text`

**Notes:**
- Cells are written after all DataFrame data, so they can overwrite existing values
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

## Hyperlinks

Add clickable links to cells:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'company': ['Anthropic', 'Google', 'Microsoft'],
    'product': ['Claude', 'Gemini', 'Copilot'],
})

# Add hyperlinks to a new column (C) after the data columns (A, B with header)
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

## Comments/Notes

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

## Checkboxes

Add interactive checkboxes to cells (Excel for Microsoft 365, Sept 2024+). Renders as `TRUE` or `FALSE` that can be toggled in Excel:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'Task': ['Write docs', 'Run tests', 'Ship release']})

xlsxturbo.df_to_xlsx(df, "checklist.xlsx",
    checkboxes={
        # Bare bool - simplest form
        'B2': True,
        'B3': False,
        'B4': False,
        # Dict form with cell format
        'C2': {'checked': True, 'format': {'bg_color': '#C6EFCE', 'bold': True}},
    }
)
```

**Checkbox format:**
- Simple: `{'A1': True}` or `{'A1': False}`
- With format: `{'A1': {'checked': True, 'format': {...}}}`

**Available options (dict form):**
- `checked` (bool, required): Initial state
- `format` (dict): Optional cell format. Accepts the same keys as [Column Formatting](formatting.md#column-formatting) (bg_color, font_color, border, bold, etc.)

**Notes:**
- Checkboxes are written AFTER DataFrame data — use cell refs that don't collide with data rows
- Requires Excel for Microsoft 365 (Sept 2024 or later); older versions will display the underlying boolean value instead
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode
