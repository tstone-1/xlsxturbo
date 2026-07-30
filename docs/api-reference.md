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
- **returns** `None`.

Everything else is a keyword argument. See [DataFrame export](dataframe-export.md).

## `dfs_to_xlsx(sheets, output_path, **options)`

Writes several DataFrames into one workbook, one sheet each.

- **`sheets`** — a list of `(df, sheet_name)` or `(df, sheet_name, options)` tuples. The
  two-tuple form uses the top-level options; the three-tuple form overrides them for that
  sheet only.
- **`output_path`** — `str` or `os.PathLike`.
- **returns** `None`.

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

## Exceptions

Six classes, all exported from `xlsxturbo`:

```
XlsxTurboError                  # base -- catches everything the library raises
├── ConfigurationError          # also ValueError
│   └── WorkbookValidationError
├── ConfigurationTypeError      # also TypeError
├── InputDataError              # also ValueError
└── FileError                   # also OSError and ValueError
```

Each one keeps the builtin exception its failures raised before 0.19.0, so
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
