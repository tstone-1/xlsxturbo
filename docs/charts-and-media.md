# Charts and media

Charts are native Excel charts, not pictures: the reader can click one and edit its
series, and it updates if the underlying cells change. Images and textboxes are
anchored to a cell and float above the grid.

## Native Excel Charts

Add editable Excel charts anchored to cells. Use `data_range`/`values_range` for a single series, or `series` for multiple series.

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'month': ['Jan', 'Feb', 'Mar'],
    'sales': [120, 145, 160],
    'margin': [32, 41, 48],
})

xlsxturbo.df_to_xlsx(df, "charts.xlsx",
    charts={
        'E2': {
            'type': 'column',
            'series': [
                {'name': 'Sales', 'values_range': 'Sheet1!$B$2:$B$4'},
                {'name': 'Margin', 'values_range': 'Sheet1!$C$2:$C$4'},
            ],
            'categories_range': 'Sheet1!$A$2:$A$4',
            'title': 'Quarter Results',
            'x_axis_name': 'Month',
            'y_axis_name': 'Amount',
            'width': 720,
            'height': 480,
            'show_data_table': True,
            'legend_position': 'bottom',
        }
    }
)
```

**Chart format:**
- `{'E2': {'type': 'bar', 'data_range': 'Sheet1!$B$2:$B$10'}}`
- `{'E2': {'type': 'line', 'series': [{'values_range': 'Sheet1!$B$2:$B$10', 'name': 'Sales'}]}}`

**Available options:**
- `type` (str, required): `area`, `bar`, `column`, `doughnut`, `line`, `pie`, `radar`, `scatter`, `stock`, plus stacked variants
- `data_range`, `values_range`, `values` (str): Range for a single data series
- `categories_range`, `categories` (str): Category/X-axis range
- `series` (list): Multiple series, each with `values_range`/`data_range`, optional `categories_range`, and optional `name`
- `title`, `x_axis_name`, `y_axis_name` (str): Chart and axis titles
- `width`, `height`, `x_offset`, `y_offset` (int pixels): Size and position
- `style` (int): Excel chart style ID, 1-48
- `show_data_table` (bool): Show data table below the chart
- `show_legend` (bool): Show or hide legend
- `legend_position` (str): `right`, `left`, `top`, `bottom`, `top_right`

**Notes:**
- Charts are native Excel chart objects, not static images
- Value/category ranges must include a sheet name (e.g. `'Sheet1!$B$2:$B$10'`); a bare range like `'$B$2:$B$10'` raises `ValueError`
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode

## Sparklines

Add sparklines - mini charts that live inside a single cell - to show trends next to your data. The dict key is the sparkline *location*: a single cell (e.g. `'D2'`) places one sparkline, while a range (e.g. `'D2:D10'`) places a grouped sparkline, one per row of the data range. The `range` key (the data to plot) is required and must be sheet-qualified (e.g. `'Sheet1!A2:C10'`), like chart ranges.

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({
    'q1': [10, 30, 20],
    'q2': [15, 25, 35],
    'q3': [25, 20, 45],
    'trend': [None, None, None],  # column to hold the sparklines
})

xlsxturbo.df_to_xlsx(df, "sparklines.xlsx",
    sparklines={
        # One sparkline per row, plotting that row's q1:q3 values into column D
        'D2:D4': {
            'range': 'Sheet1!A2:C4',
            'type': 'line',
            'markers': True,
            'high_point': True,
            'low_point': True,
        }
    }
)
```

**Sparkline format:**
- `{'D2': {'range': 'Sheet1!A2:C2', 'type': 'column'}}` - a single sparkline
- `{'D2:D10': {'range': 'Sheet1!A2:C10', 'type': 'line'}}` - a grouped sparkline (one per row)

**Available options:**
- `range` (str, required): The data range to plot, sheet-qualified (e.g. `'Sheet1!A2:C10'`); 1D for a single cell, 2D for a group
- `type` (str): `line` (default), `column`, or `win_loss`
- `style` (int): Built-in sparkline style ID, 1-36
- `markers`, `high_point`, `low_point`, `first_point`, `last_point`, `negative_points` (bool): Point highlighting
- `show_axis` (bool): Show a horizontal axis line
- `color` (str): Sparkline series color (`'#RRGGBB'` or named)
- `high_point_color`, `low_point_color`, `first_point_color`, `last_point_color`, `negative_points_color`, `markers_color` (str): Per-feature colors
- `line_weight` (float): Line weight in points (line sparklines)
- `custom_max`, `custom_min` (float): Fixed vertical-axis bounds
- `group_max`, `group_min` (bool): Share a common max/min across a grouped sparkline
- `date_range` (str): Sheet-qualified range supplying X-axis date values
- `right_to_left`, `column_order`, `show_hidden_data` (bool): Plot direction and hidden-data handling

**Notes:**
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- `range` and `date_range` must be sheet-qualified Excel ranges (e.g. `'Sheet1!A2:C10'`), as with chart ranges
- Not available in constant memory mode

## Images

Embed images in cells. Unknown keys raise errors (see [Header Styling](formatting.md#header-styling)).

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

## Textboxes

Add floating text shapes (callouts, annotations) that sit on top of cells. Unknown keys raise errors (both at the top level and inside `font`).

```python
import xlsxturbo
import pandas as pd

df = pd.DataFrame({'Region': ['North', 'South'], 'Sales': [120, 95]})

xlsxturbo.df_to_xlsx(df, "report.xlsx",
    textboxes={
        # Bare string - simplest form, default size/style
        'D2': 'Simple note',
        # Dict form with full options
        'E5': {
            'text': 'Q4 target met for all regions',
            'width': 220,
            'height': 80,
            'x_offset': 10,
            'y_offset': 5,
            'font': {
                'name': 'Arial',
                'size': 12,
                'bold': True,
                'italic': False,
                'underline': False,
                'color': '#2C3E50',
            },
            'fill_color': '#ECF0F1',
            'line_color': '#34495E',
            'alt_text': 'Q4 summary callout',
        },
    }
)
```

**Textbox format:**
- Simple: `{'D2': 'Some text'}`
- With options: `{'D2': {'text': '...', 'width': 200, 'font': {'bold': True}, ...}}`

**Available options (dict form):**
- `text` (str, required): Textbox contents
- `width`, `height` (int pixels): Shape size. Defaults are 192 × 120 pixels
- `x_offset`, `y_offset` (int pixels): Shift within the anchor cell
- `font` (dict): Font options — `name`, `size` (points), `bold`, `italic`, `underline`, `color` (hex or named)
- `fill_color` (str): Background fill — hex `#RRGGBB` or named color
- `line_color` (str): Border line — hex `#RRGGBB` or named color
- `alt_text` (str): Alternative text for accessibility

**Notes:**
- Textboxes are floating shapes anchored to a cell, not cell-content — they overlay cells without overwriting them
- Works with both `df_to_xlsx` and `dfs_to_xlsx` (global or per-sheet)
- Not available in constant memory mode
