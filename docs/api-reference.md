# API reference

The public API is four names, re-exported from the compiled extension:

```python
from xlsxturbo import csv_to_xlsx, df_to_xlsx, dfs_to_xlsx, version
```

The package ships a `py.typed` marker and full type stubs, so the authoritative
signatures — including every keyword argument, its type, and its default — are the ones
your editor and type checker already see. This page describes the shape of each entry
point and where to find the detail; it deliberately does not restate the parameter lists,
which would be a second copy to keep in step with the first.

For which options each function accepts, use the
**[capability matrix](capability-matrix.md)**. It is generated from the Rust sources, so
it cannot drift from what the extension actually takes.

## `df_to_xlsx(df, output_path, **options)`

Writes one pandas or polars DataFrame to one worksheet.

- **`df`** — any object exposing a pandas- or polars-compatible interface. The type is
  detected at runtime; neither library is a dependency of xlsxturbo.
- **`output_path`** — `str` or `os.PathLike`.
- **returns** `tuple[int, int]` — the number of rows and columns written, the header
  row included when `header=True`.

Everything else is a keyword argument. See [DataFrame export](dataframe-export.md).

## `dfs_to_xlsx(sheets, output_path, **options)`

Writes several DataFrames into one workbook, one sheet each.

- **`sheets`** — a list of `(df, sheet_name)` or `(df, sheet_name, options)` tuples. The
  two-tuple form uses the top-level options; the three-tuple form overrides them for that
  sheet only. A `list` of the same shape is accepted wherever a tuple is shown; the tuple is
  the recommended form and the one the type stubs describe.
- **`output_path`** — `str` or `os.PathLike`.
- **returns** `list[tuple[int, int]]` — one `(rows, cols)` pair per sheet, in the order
  the sheets were given.

Top-level keyword arguments become the default for every sheet. Not every option is
overridable per sheet; the capability matrix has the exact set. See
[Multi-sheet workbooks](multi-sheet-workbooks.md).

## `csv_to_xlsx(input_path, output_path, sheet_name=..., parallel=..., date_order=...)`

Converts a CSV file to `.xlsx`, detecting each value's type from its text.

- **returns** `tuple[int, int]` — the number of rows and columns written.

This is a much smaller surface than the DataFrame functions: it takes no formatting,
table, chart, or validation options at all. See [CSV conversion](csv-conversion.md).

## `version()`

Returns the extension's version as a string. `xlsxturbo.__version__` holds the same
value. Both come from the compiled extension rather than from package metadata, which
makes them the right thing to report in a bug report — they describe the binary that
actually ran.

## Annotating options: `xlsxturbo.types`

Every option that takes a dict has a `TypedDict`, and every option that takes one of a fixed
set of strings has a `Literal` alias. They are real runtime objects in `xlsxturbo.types`, so
no `TYPE_CHECKING` guard is needed:

```python
from xlsxturbo.types import ChartOptions, HeaderFormat

header: HeaderFormat = {"bold": True, "bg_color": "#DDDDDD"}
chart: ChartOptions = {
    "type": "column",
    "categories": "Sheet1!$A$2:$A$10",
    "values": "Sheet1!$B$2:$B$10",
}

xlsxturbo.df_to_xlsx(df, "out.xlsx", header_format=header, charts={"D2": chart})
```

Chart and sparkline ranges must name their sheet: an unqualified `"B2:B10"` is refused
rather than quietly plotting the wrong thing (a values range without a `!` produces a
misleading error from the writer, and a categories range without one is ignored
altogether). See [Charts and media](charts-and-media.md).

The module imports nothing beyond the standard library, so importing it costs nothing and
works before the extension is built. `SheetOptions` is the shape of a `dfs_to_xlsx`
per-sheet dict, and `PathArg` is what the path parameters accept. `__all__` names the option
shapes and aliases and nothing else, so `import *` brings those in without the typing
helpers they are built from.

**Fields the library requires are marked required.** `ImageOptions` needs `path`,
`ChartOptions` needs `type`, `SparklineOptions` needs `range`, and so on — so a checker
rejects `images={"D1": {}}` rather than leaving it for the runtime. `ChartSeriesOptions` is
the exception: it requires one of `values_range` / `values` / `data_range`, which a
`TypedDict` cannot express, so all three stay optional to the checker and the runtime
enforces the choice.

Field annotations are unevaluated strings, so `typing.get_type_hints()` resolves them and
anything building a schema from these shapes -- pydantic, FastAPI, attrs -- works. That was
not true before 1.1.0: on Python 3.9 a `bool | str` annotation could be written but not
evaluated, so resolving the hints raised. Dropping 3.9 removed the split.

## Exceptions

Seven classes, all exported from `xlsxturbo`:

```
XlsxTurboError                      # base -- catches everything the library raises
├── OptionError                     # never raised itself; catches both of its children
│   ├── ConfigurationError          # also ValueError
│   │   └── WorkbookValidationError
│   └── ConfigurationTypeError      # also TypeError
├── InputDataError                  # also ValueError
└── FileError                       # also OSError and ValueError
```

`OptionError` exists so that `except OptionError` catches every problem with what you
passed -- a bad value and a wrong type -- and nothing else. It has no builtin base of its
own, which is what keeps the value/type split meaningful to an `except` clause.

Each of the others keeps the builtin exception its failures raised before 0.19.0, so
`except ValueError` and `except TypeError` behave as they always did. See
[Errors and warnings](errors.md) for which failures land where, and for the two places the
classification is deliberately coarser than the class names suggest.

## Option value types

Options fall into a few recurring shapes, and knowing which one you are looking at
usually answers the question faster than the stub does:

| Shape | Example options | Keyed by |
|-------|-----------------|----------|
| Scalar flag or name | `header`, `autofit`, `table_style`, `freeze_panes` | — |
| Column-keyed mapping | `column_widths`, `column_formats`, `formula_columns` | Column index, name, or wildcard pattern |
| Cell-keyed mapping | `cells`, `comments`, `hyperlinks`, `images`, `charts` | An A1-style reference such as `"B7"` |
| Row-keyed mapping | `row_heights` | Row index |
| Format dictionary | `header_format`, and the nested `format` key in many options | Format property names |

Two rules hold across all of them:

- **Unknown keys are rejected.** Every option dictionary validates its keys and raises on
  one it does not know, rather than ignoring it. See [Errors](errors.md).
- **Iteration order is preserved.** Cell- and column-keyed options are applied in the
  order you supplied them, and identical input produces byte-identical output.

## Ordering guarantees

Within a sheet, options are applied in a fixed order, and `cells` is applied last. That
is deliberate: an explicit `cells` entry can overwrite a value that came from the
DataFrame, which is what makes it usable for corrections and annotations over the data.
See [Individual cells](cells.md).
