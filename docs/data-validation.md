# Data validation

Validation turns a column into a dropdown or constrains what can be typed into it.
Excel enforces the rule on input only; it does not re-check values already in the
cells, including the ones xlsxturbo wrote.

## Data Validation

Add dropdowns and input constraints. Unknown keys raise errors (see [Header Styling](formatting.md#header-styling)).

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
- `whole_number` `min`/`max` are bounded to the i32 range (-2147483648 to 2147483647); a value outside that range raises `ValueError` naming the field and range
- List validation values are limited to 255 total characters (Excel limitation)
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode
