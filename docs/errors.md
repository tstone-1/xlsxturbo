# Errors and warnings

xlsxturbo validates eagerly and fails loudly. An option it does not recognise is an
error, not a value it quietly ignores — so a misspelled key surfaces at the call rather
than as a missing style in a spreadsheet somebody else opens next week.

## What gets raised

Every failure today is either a `TypeError` or a `ValueError`:

| Exception | Raised for |
|-----------|------------|
| `TypeError` | A value of the wrong Python type — a string where a dict is expected, a non-string dictionary key |
| `ValueError` | Everything else: unrecognised option keys, out-of-range numbers, unparseable colours, malformed cell references, invalid chart ranges, and I/O failures |

Messages carry the option and the specific key or cell reference that caused them, so
they identify the offending entry rather than just the option family:

```python
>>> df_to_xlsx(df, "out.xlsx", comments={"B2": {"txt": "note"}})
ValueError: comments['B2']: unknown key 'txt'
```

### I/O failures also arrive as `ValueError`

A missing output directory, a full disk, or a dropped network share raises `ValueError`
with a message beginning `Failed to save workbook to '<path>': `, not `OSError`. This is
a known wart rather than a deliberate design: the extension converts internal errors to
Python exceptions at a single boundary that currently classifies everything as
`ValueError`. Code that needs to distinguish an I/O problem from a bad option has to
match on the message today.

A typed exception hierarchy rooted at `XlsxTurboError` is planned. It will inherit from
`ValueError` and `TypeError` so that existing `except ValueError:` handlers keep working
unchanged.

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
