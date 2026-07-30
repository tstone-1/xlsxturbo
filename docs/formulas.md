# Formulas and defined names

xlsxturbo writes formulas; it does not evaluate them. Excel computes the results when
the file is opened, which means a formula referring to a cell xlsxturbo never wrote is
not an error here -- it becomes a `#REF!` in Excel.

## Formula Columns

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
- Combined with `table_style`, formula columns sit **outside** the Excel table: the table
  covers the DataFrame columns only, so formula columns get no banded fill, no autofilter
  dropdown, and are not covered by `column_widths`/`autofit`. Add the calculated column to
  the DataFrame instead if you need it inside the table.

## Defined Names

Create workbook-level named ranges that can be referenced in formulas:

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'product': ['Widget A', 'Widget B', 'Widget C'],
    'price': [19.99, 29.99, 39.99],
    'quantity': [100, 75, 50]
})

# Define named ranges for use in formulas or external references
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    defined_names={
        "PriceRange": "=Sheet1!$B$2:$B$4",
        "AllData": "=Sheet1!$A$1:$C$4"
    }
)

# Works with multi-sheet workbooks too
df1 = pd.DataFrame({'x': [1, 2, 3]})
df2 = pd.DataFrame({'y': [4, 5, 6]})
xlsxturbo.dfs_to_xlsx([
    (df1, "Data"),
    (df2, "Summary")
], "multi.xlsx",
    defined_names={
        "DataRange": "=Data!$A$1:$A$4",
        "SummaryRange": "=Summary!$A$1:$A$4"
    }
)
```

**Notes:**
- Defined names are workbook-level (not per-sheet)
- References must use Excel notation with sheet name: `=Sheet1!$A$1:$D$100`
- Works with both `df_to_xlsx` and `dfs_to_xlsx`
