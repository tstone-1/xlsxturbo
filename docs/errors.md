# Errors and warnings

xlsxturbo validates eagerly and fails loudly. An option it does not recognise is an
error, not a value it quietly ignores — so a misspelled key surfaces at the call rather
than as a missing style in a spreadsheet somebody else opens next week.

## What gets raised

Every failure xlsxturbo itself raises is an `XlsxTurboError`. Catching that one class
catches everything the library validates, and nothing else ([one narrow
exception](#what-is-not-in-the-hierarchy)):

```python
import xlsxturbo

try:
    xlsxturbo.df_to_xlsx(df, "out.xlsx", header_format={"bold": True})
except xlsxturbo.XlsxTurboError as exc:
    log.error("export failed: %s", exc)
```

Five subclasses say what kind of failure it was:

| Exception | Raised for | Also a |
|-----------|------------|--------|
| `ConfigurationError` | An option or argument **value** is invalid: an unrecognised key, an unparseable colour or cell reference, an out-of-range number, an invalid chart range | `ValueError` |
| `ConfigurationTypeError` | An option or argument has the wrong **type**: a list where a dict is required, a non-string dictionary key, a `bytes` path | `TypeError` |
| `InputDataError` | The object passed as data is not a pandas or polars DataFrame, or its columns cannot be read | `ValueError` |
| `FileError` | A filesystem read or write failed: a missing output directory, a permissions problem, a full disk, an unreadable CSV input | `OSError`, `ValueError` |
| `WorkbookValidationError` | The configuration is well-formed but Excel forbids it — two sheets claiming the same table name | `ConfigurationError`, so also `ValueError` |

Messages carry the option and the specific key or cell reference that caused them, so
they identify the offending entry rather than just the option family:

```python
>>> df_to_xlsx(df, "out.xlsx", comments={"B2": {"txt": "note"}})
xlsxturbo.ConfigurationError: comments['B2']: unknown key 'txt'
```

### Existing `except ValueError` and `except TypeError` still work

The hierarchy arrived in 0.19.0 and changed nothing about which builtin any given failure
raises. Every class above inherits the builtin that the same failure raised in 0.18 and
earlier, so upgrading cannot break a handler you already have:

```python
# Written against 0.18. Still correct, and still catches exactly what it did.
try:
    xlsxturbo.df_to_xlsx(df, path, column_widths={"A": 12})
except ValueError as exc:
    ...
```

That constraint is why the mapping is not always the one you would pick from scratch.
`InputDataError` is a `ValueError` rather than a `TypeError`, even though "you passed a
list, not a DataFrame" is really a type problem — that failure has always raised
`ValueError`, and moving it would break working code. Catch `InputDataError` when you
want the distinction.

### I/O failures are an `OSError` — and still a `ValueError`

`FileError` inherits both, so either handler works:

```python
try:
    xlsxturbo.df_to_xlsx(df, "/nonexistent/dir/out.xlsx")
except xlsxturbo.FileError as exc:
    ...     # what to write today
except OSError:
    ...     # what you would write for any file operation
except ValueError:
    ...     # what 0.18 required
```

Two details worth knowing. The message still begins `Failed to save workbook to '<path>': `
for a save failure, so message-matching code written against 0.18 keeps working. And
because the exception is constructed from a single message, `errno`, `strerror` and
`filename` are always `None` — the path is in the message text, not in those fields.

### What is not in the hierarchy

Argument conversion performed by the Python/Rust binding before xlsxturbo sees the value
raises a plain `TypeError`. Passing a non-string where the signature requires `str` is the
usual way to hit this. Everything xlsxturbo itself validates is in the hierarchy.

Two dict-valued options fall on the binding's side of that line, which is worth knowing
because their neighbours do not: **`row_heights` and `defined_names`** are typed in the
signature, so a wrong inner type is rejected by the conversion and arrives as a plain
`TypeError`. `column_widths` looks identical from Python and is read by an extractor, so the
same mistake there is a `ConfigurationTypeError`. Catch `(XlsxTurboError, TypeError)` if you
need to treat all option-type mistakes alike.

A dtype problem discovered deep in the write pipeline surfaces as `ConfigurationError`
rather than `InputDataError`: the pipeline reports it as a message without a category, and
inventing one from the text would be guesswork. `InputDataError` covers the case that
matters — the object is not a supported DataFrame at all — which is checked before any
work starts.

## One combination is refused outright

A `data_bar` conditional format and a sparkline on the **same worksheet** raise
`ConfigurationError`:

```python
xlsxturbo.df_to_xlsx(
    df, "out.xlsx",
    conditional_formats={"Score": {"type": "data_bar"}},
    sparklines={"D2": {"range": "Sheet1!A2:C2"}},
)
# ConfigurationError: sheet 'Sheet1': conditional_formats['Score'] is a data bar
# and this sheet also has sparklines. ...
```

The underlying Excel writer produces a corrupt workbook for that pair — Excel opens it
and offers to repair it. The defect is upstream, in rust_xlsxwriter 0.97.0, and cannot be
fixed here, so xlsxturbo refuses rather than writing a file you cannot open. Failing at
the call is the lesser harm: a corrupt workbook is typically discovered by whoever you
sent it to.

Put the sparklines on a different sheet, or use `2_color_scale` / `3_color_scale`. The
restriction is narrow — every other conditional-format type works beside sparklines, and
each feature alone is unaffected.

## Nothing is half-written

Validation of a given option happens as that option is applied, and the workbook is only
serialised once every option has been applied. Two consequences:

- **A validation error leaves no output file at all.** If the destination already held a
  file, it is exactly as it was.
- **A failure during the save leaves the destination untouched too.** The workbook is
  written to a temporary file in the destination's own directory and renamed over the
  target only once it is complete. Re-exporting over yesterday's report can therefore
  never leave you with neither.

Because the staging file is created beside the destination, that directory must exist and
be writable — a path into a directory that does not exist is a save failure, not a
silently created tree. When an existing file is replaced, its permissions are preserved.

See [Compatibility and limitations](compatibility.md) for the full note on output safety.

## Warnings

`constant_memory=True` cannot apply options that need a second pass over a finished
sheet. Rather than dropping them silently, xlsxturbo emits a `RuntimeWarning` naming the
sheet and every option it skipped:

```
RuntimeWarning: sheet 'Data': constant_memory=True disables these features:
table_style, autofit, conditional_formats
```

New options default to skipped-and-warned, which is the safe direction: a feature that
turns out to be compatible gets promoted deliberately, rather than a new one being
silently ignored. The [capability matrix](capability-matrix.md) lists which options
survive constant-memory mode.

Since it is a real warning, it can be promoted to an error in a batch job or a test
suite that should not tolerate a silently reduced export:

```python
import warnings

warnings.simplefilter("error", RuntimeWarning)
```

## Reporting a problem

A crash, a corrupt output file, or a message that misidentifies the cause is a bug worth
reporting — see [CONTRIBUTING.md](https://github.com/tstone-1/xlsxturbo/blob/main/CONTRIBUTING.md).
A file Excel refuses to open is the most valuable kind of report; include the calling
code and the xlsxturbo version from `xlsxturbo.version()`.
