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

# A name that would collide with a cell address gets the same prefix
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    table_style="Medium2",
    table_name="Q1"  # Becomes "_Q1"
)
```

### Name sanitization rules

The name you pass is adjusted to something Excel accepts. In order:

| Input | Result | Rule |
|-------|--------|------|
| `"Verka\u0308ufe"` | `"Verkäufe"` | The name is normalised to NFC first |
| `"Sales Q1!"` | `"Sales_Q1_"` | Anything that is not a letter, digit or `_` becomes `_` |
| `"2024Sales"` | `"_2024Sales"` | A leading digit gains a `_` prefix |
| `"Q1"`, `"A1"`, `"XFD1048576"` | `"_Q1"`, `"_A1"`, `"_XFD1048576"` | Excel reserves names that address a cell |
| `"R"`, `"C"`, `"R1C1"` | `"_R"`, `"_C"`, `"_R1C1"` | Excel reserves the R1C1 forms and the selection shortcuts `R`/`C` |
| `"R2D2"`, `"C3PO"`, `"R1_total"` | `"_R2D2"`, `"_C3PO"`, `"_R1_total"` | Excel ignores whatever follows a complete R1C1 reference |
| `"TRUE"`, `"false"` | `"_TRUE"`, `"_false"` | Excel reserves its logical constants, in any case |
| a name over 255 characters | first 255 characters | Excel's length limit |

The cell-reference rule is bounded by the actual grid, which ends at
`XFD1048576`. `"AAAA1"` and `"A1048577"` address no cell, so Excel takes them as
ordinary names and they pass through unchanged. A zero-padded row counts by the
row it parses to: `"A01"` is treated as a reference (Excel offers to repair a
workbook with a table of that name, exactly as for `"Q1"`), while `"A0"` — row
zero — addresses nothing and passes through.

Excel stops reading an R1C1 form once it has the index and ignores the rest, so
`"R2D2"` is the reference `R2` with `D2` trailing and gets the prefix. Only the
leading index has to exist, which is why `"R1C16385"` is prefixed even though
that column is past the grid. A name that never reaches an index is not a
reference at all: `"RCx"` and `"Rate1"` pass through unchanged.

Normalisation happens before the character rule, and it matters because a
combining mark is not a letter. Without it, `"Verkäufe"` typed as `"Verka"` plus
`U+0308 COMBINING DIAERESIS` — which is what a lot of text on macOS looks like —
would arrive as `"Verka_ufe"`. Composing first repairs every mark that has a
precomposed form. It is NFC and not NFKC, so compatibility characters stay
distinct: `"Ａ1"` (fullwidth A) is not folded to `"A1"` and does not become
`"_A1"`.

Marks with **no** composed form are still replaced, and Excel would have
accepted them: `"ไม่"` becomes `"ไม_"` (Thai tone mark `U+0E48`) and `"हिन्दी"`
becomes `"हिन_दी"` (Hindi virama `U+094D`). If you need those names exactly,
pass a spelling that survives the character rule.

Workbook-level `defined_names` are **not** sanitized — a reference-shaped key
there raises `ConfigurationError` instead, because silently renaming a defined
name would leave the formulas that use it pointing at a name the workbook no
longer defines.

### A table name may not equal a defined name

Excel requires the two kinds of name to be unique against each other, not only
within their own kind, and repairs a workbook that carries both. The name is
compared after sanitization and ignoring case, and a sheet-scoped defined name
collides just as a global one does:

```python
xlsxturbo.df_to_xlsx(df, "report.xlsx",
    table_style="Medium2",
    table_name="Sales",
    defined_names={"Sheet1!Sales": "=Sheet1!$A$1:$A$4"},
)
# WorkbookValidationError: defined_names['Sheet1!Sales'] collides with the
# table name 'Sales' on sheet 'Sheet1'.
```

The check runs before anything is written, so the output file is left alone.
A sheet that creates no table — no `table_style`, or an empty DataFrame —
claims no name and cannot collide.
