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

The subclasses say what kind of failure it was:

| Exception | Raised for | Also a |
|-----------|------------|--------|
| `OptionError` | Never raised itself — the parent of the two below, so `except OptionError` catches every problem with what you passed | — |
| `ConfigurationError` | An option or argument **value** is invalid: an unrecognised key, an unparseable colour or cell reference, an out-of-range number, an invalid chart range, an [image file that cannot be read](#save-time-validation-lands-in-fileerror) | `OptionError`, `ValueError` |
| `ConfigurationTypeError` | An option or argument has the wrong **type**: a list where a dict is required, a non-string dictionary key, a `bytes` path | `OptionError`, `TypeError` |
| `InputDataError` | The object passed as data is not a pandas or polars DataFrame, or its columns cannot be read | `ValueError` |
| `FileError` | A filesystem read or write failed: a missing output directory, a permissions problem, a full disk, an unreadable CSV input — and the [workbook rules Excel's writer checks only during the save](#save-time-validation-lands-in-fileerror) | `OSError`, `ValueError` |
| `WorkbookValidationError` | The configuration is well-formed but Excel forbids it. Raised by the two name pre-checks: two sheets claiming the same table name, and a table name that collides with a `defined_names` key | `ConfigurationError`, so also `OptionError` and `ValueError` |

### "Anything wrong with what I passed"

`ConfigurationError` and `ConfigurationTypeError` split a bad *value* from a
wrong *type*. That distinction is worth having and is usually not the one you
want when writing a handler — most callers want both and nothing else:

```python
try:
    xlsxturbo.df_to_xlsx(df, path, **user_supplied_options)
except xlsxturbo.OptionError as exc:
    return {"error": f"bad export options: {exc}"}    # the caller's mistake
except xlsxturbo.FileError as exc:
    return {"error": f"could not write the file: {exc}"}   # not their mistake
```

`OptionError` arrived in 0.21.0 by reparenting, not renaming, so every existing
`except` clause means exactly what it did before.

Two failures cross that split rather than following it — see [Save-time validation lands
in `FileError`](#save-time-validation-lands-in-fileerror) before writing a handler that
treats the two branches as "their mistake" and "ours".

Messages carry the option and the specific key or cell reference that caused them, so
they identify the offending entry rather than just the option family:

```python
>>> df_to_xlsx(df, "out.xlsx", comments={"B2": {"txt": "note"}})
xlsxturbo.ConfigurationError: comments['B2']: unknown option 'txt'. Valid: text, author
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

**`errno` is set**, so you can tell one filesystem failure from another without
matching on message text:

```python
import errno

try:
    xlsxturbo.df_to_xlsx(df, "/mnt/reports/out.xlsx")
except xlsxturbo.FileError as exc:
    if exc.errno == errno.ENOSPC:
        ...          # the disk is full: retry elsewhere
    elif exc.errno == errno.ENOENT:
        ...          # the directory is gone: create it, or fail the job
    raise
```

Two details worth knowing. The message still begins `Failed to save workbook to '<path>': `
for a save failure, so message-matching code written against 0.18 keeps working — and
that is exactly why **`strerror` and `filename` stay `None`**. `OSError` reformats its
own `str()` into `[Errno n] strerror: 'filename'` as soon as `filename` is set, which
would throw the message away; the path is already in it. `errno` is the one field that
can be populated without changing what the exception prints.

On platforms whose native error numbers are not POSIX — Windows — `errno` carries the
POSIX equivalent rather than the native code, so the comparison above means the same
thing everywhere. Where no meaningful number is available it stays `None`, so compare
against a specific constant rather than testing for truthiness.

### Save-time validation lands in `FileError`

Not every rule Excel imposes can be checked when the option is applied. The underlying
writer defers several of them to the moment the workbook is serialised, and whatever
fails there is reported by the save — so a mistake in what you passed can arrive as a
`FileError`:

```python
xlsxturbo.dfs_to_xlsx([(df, "Data"), (df, "Data")], "out.xlsx")
# FileError: Failed to save workbook to 'out.xlsx': Worksheet name 'data' has
# already been used in this workbook.
```

A duplicate sheet name and a chart range naming a sheet the workbook does not contain
are the two you are most likely to meet. Both keep the `Failed to save workbook to
'<path>': ` prefix, and the writer's own sentence after it names the rule that was broken.

The mirror case runs the other way: an image xlsxturbo cannot open is a
`ConfigurationError`, not a `FileError`, because the path arrived as an option and the
failure is reported by the layer that applies options.

```python
xlsxturbo.df_to_xlsx(df, "out.xlsx", images={"A1": "logo.png"})
# ConfigurationError: Failed to load image 'logo.png': No such file or directory (os error 2)
```

Both messages name the sheet, the option or the path, and `except
xlsxturbo.XlsxTurboError` catches either — which is what a top-level export handler
wants. The distinction matters for a handler that branches on whose fault the failure
was: a `FileError` is not on its own evidence that the disk was involved, and a
`ConfigurationError` is not on its own evidence that it was not.

### What is not in the hierarchy

Argument conversion performed by the Python/Rust binding before xlsxturbo sees the value
raises a plain `TypeError`. Passing a non-string where the signature requires `str` is the
usual way to hit this. Everything xlsxturbo itself validates is in the hierarchy.

One dict-valued option falls on the binding's side of that line, which is worth knowing
because its neighbours do not: **`defined_names`** is typed in the signature, so a wrong
inner type is rejected by the conversion and arrives as a plain `TypeError`.
`column_widths` and `row_heights` look identical from Python and are read by extractors,
so the same mistake in either is a `ConfigurationTypeError`. Catch
`(XlsxTurboError, TypeError)` if you need to treat all option-type mistakes alike.

`row_heights` was on the binding's side of the line until 1.3.x, and moving it fixed
more than the class name. The conversion accepted `{True: 40}` — `bool` subclasses `int`
in Python — and silently sized row 2; it accepted `{0: True}` as a height of one point;
and it answered a negative or out-of-range key with `OverflowError`, which is an
`ArithmeticError` and so is caught by neither `except XlsxTurboError` nor the
`except (XlsxTurboError, TypeError)` recommended above. Those are now
`ConfigurationTypeError` and `ConfigurationError`, each naming the entry:

```python
>>> df_to_xlsx(df, "out.xlsx", row_heights={-1: 20})
xlsxturbo.ConfigurationError: row_heights['-1']: must be a non-negative row index
```

A dtype problem discovered deep in the write pipeline surfaces as `ConfigurationError`
rather than `InputDataError`: the pipeline reports it as a message without a category, and
inventing one from the text would be guesswork. `InputDataError` covers the case that
matters — the object is not a supported DataFrame at all — which is checked before any
work starts.

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
