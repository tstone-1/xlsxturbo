# Conditional formatting

Conditional formats are evaluated by Excel when the file is opened, so they stay live
as the reader edits or filters the data.

## Conditional Formatting

Apply visual formatting based on cell values. Unknown keys in the nested `format` dict raise errors (see [Header Styling](formatting.md#header-styling)).

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
| `cell` | `criteria`, `value`, `min_value`, `max_value`, `format` |

**Available icon types:**
- 3 icons: `3_arrows`, `3_arrows_gray`, `3_flags`, `3_traffic_lights`, `3_traffic_lights_rimmed`, `3_signs`, `3_symbols`, `3_symbols_uncircled`
- 4 icons: `4_arrows`, `4_arrows_gray`, `4_traffic_lights`, `4_rating`
- 5 icons: `5_arrows`, `5_arrows_gray`, `5_quarters`, `5_rating`

**Cell rules** — highlight cells based on value conditions:
```python
# Single rule
conditional_formats={
    'status': {
        'type': 'cell',
        'criteria': 'equal_to',
        'value': 'ERROR',
        'format': {'bg_color': '#FF0000', 'font_color': 'white', 'bold': True}
    }
}

# Multiple rules on one column (pass a list)
conditional_formats={
    'severity': [
        {'type': 'cell', 'criteria': 'equal_to', 'value': 'HIGH', 'format': {'bg_color': '#FF0000'}},
        {'type': 'cell', 'criteria': 'equal_to', 'value': 'MEDIUM', 'format': {'bg_color': '#FFA500'}},
        {'type': 'cell', 'criteria': 'equal_to', 'value': 'LOW', 'format': {'bg_color': '#FFFF00'}},
    ]
}

# Numeric comparison
conditional_formats={'score': {'type': 'cell', 'criteria': 'between', 'min_value': 0, 'max_value': 50, 'format': {'bg_color': '#FF0000'}}}
```

**Available criteria for `cell` type:**

| Criteria | Value keys | Description |
|----------|-----------|-------------|
| `equal_to`, `not_equal_to` | `value` | Exact match (string or number) |
| `greater_than`, `less_than` | `value` | Numeric comparison |
| `greater_than_or_equal_to`, `less_than_or_equal_to` | `value` | Numeric comparison |
| `between`, `not_between` | `min_value`, `max_value` | Range check |
| `containing`, `not_containing` | `value` | Text contains substring |
| `begins_with`, `ends_with` | `value` | Text prefix/suffix match |
| `blanks`, `no_blanks` | *(none)* | Empty/non-empty cells |

Column patterns work with conditional formats:
```python
# Apply data bars to all columns starting with "price_"
conditional_formats={'price_*': {'type': 'data_bar', 'bar_color': '#9B59B6'}}
```
