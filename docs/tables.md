# Excel tables

An Excel table gives the written range a name, an autofilter, banded rows, and one of
61 built-in styles. It is the cheapest way to make an export look deliberate.

## Excel Tables with Styling

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

## Named Excel Tables

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
