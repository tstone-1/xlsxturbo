# xlsxturbo

High-performance Excel writer with automatic type detection. Written in Rust, usable from Python.

[![CI](https://github.com/tstone-1/xlsxturbo/actions/workflows/ci.yml/badge.svg)](https://github.com/tstone-1/xlsxturbo/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/xlsxturbo.svg)](https://pypi.org/project/xlsxturbo/)
[![Python](https://img.shields.io/pypi/pyversions/xlsxturbo.svg)](https://pypi.org/project/xlsxturbo/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

xlsxturbo exports pandas and polars DataFrames and CSV files to `.xlsx`. It uses Rust for
the hot path and keeps the Python API small enough to drop straight into a script, a
report job, or a batch pipeline. Roughly **7-9x faster than pandas + openpyxl** on the
reference benchmarks, with the Excel features those exports usually need — tables,
conditional formatting, charts, data validation, images — available as focused keyword
arguments rather than a workbook object model.

The `.xlsx` files themselves are written by
**[rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter)** — see
[Built on rust_xlsxwriter](#built-on-rust_xlsxwriter) below.

**[Full documentation](https://tstone-1.github.io/xlsxturbo/)** ·
[Capability matrix](https://tstone-1.github.io/xlsxturbo/capability-matrix/) ·
[Changelog](CHANGELOG.md)

## Install

```bash
pip install xlsxturbo
```

Wheels are published for Python 3.10+ on Linux, Windows, and macOS. There are no runtime
dependencies beyond the interpreter.

## Export a DataFrame

```python
import pandas as pd
from xlsxturbo import df_to_xlsx

df = pd.DataFrame({
    "product": ["Widget", "Gadget", "Gizmo"],
    "price": [19.99, 34.50, 8.75],
    "in_stock": [True, False, True],
    "restock": pd.to_datetime(["2024-03-01", "2024-03-15", "2024-04-01"]),
})

df_to_xlsx(df, "products.xlsx", table_style="Medium2", autofit=True)
```

Types carry across without configuration: numbers stay numbers, booleans become Excel
booleans, and dates and datetimes become real Excel date values with a display format
attached. polars DataFrames work the same way — neither library is a dependency.

## Convert a CSV

```python
from xlsxturbo import csv_to_xlsx

csv_to_xlsx("sales.csv", "sales.xlsx")
```

Types are detected from the file's text. There is also a command-line tool for the same
job, though it is **not** included in the PyPI wheel — it has to be built from source. See
[CSV conversion](https://tstone-1.github.io/xlsxturbo/csv-conversion/).

## What it can do

- **DataFrame and CSV export** — pandas, polars, and CSV in, `.xlsx` out
- **Excel tables** with 61 built-in styles, autofilter, and banded rows
- **Formatting** — header and per-column styles, number formats, per-side borders,
  alignment, wrapping, auto-fit and explicit column widths, row heights, merged ranges,
  rich text
- **Conditional formatting** — colour scales, data bars, icon sets
- **Formulas** — calculated columns and workbook-level defined names
- **Native Excel charts** and in-cell sparklines, both editable in Excel
- **Data validation** — dropdowns, numeric ranges, text-length constraints
- **Cell-level extras** — arbitrary cell writes, hyperlinks, comments, checkboxes,
  images, textboxes
- **Multi-sheet workbooks** with per-sheet option overrides
- **Constant memory mode** for very large exports, and optional parallel CSV parsing
- **Atomic writes** — a failed export never truncates the file already at that path

The [capability matrix](https://tstone-1.github.io/xlsxturbo/capability-matrix/) is the
authoritative list: it is generated from the source and shows which options each function
accepts, which are overridable per sheet, and which survive constant-memory mode.

## Performance

On 100,000 rows x 50 columns of mixed types, xlsxturbo is about 4.6x faster than polars,
7x faster than pandas + xlsxwriter, and 9.3x faster than pandas + openpyxl. Absolute
timings are system-specific; the ratios are stable. Full tables, test systems, and
methodology are on the
[performance page](https://tstone-1.github.io/xlsxturbo/performance/), and both benchmark
suites live in [`benchmarks/`](benchmarks/) so you can measure your own hardware.

## Known limitations

- **Write-only.** xlsxturbo creates workbooks; it cannot open or modify an existing one.
- **Timezone-aware datetimes** are written as their local wall-clock value — Excel has no
  timezone concept, so the UTC offset is not preserved.
- **Integers above 2^53** are written as text to avoid silent precision loss.
- **Durations** (`Timedelta` / `timedelta64`) are written as text; Excel has no duration
  type.

The [compatibility page](https://tstone-1.github.io/xlsxturbo/compatibility/) has the
complete list with the workaround for each.

## Project status

Stable since 1.0.0. Everything reachable from `import xlsxturbo` without a leading
underscore is covered by [Semantic Versioning](https://semver.org/) and will not break
before 2.0.0; anything removed gets a `DeprecationWarning` naming its replacement and its
removal version, for at least one minor release and at least six months.

The [stability page](https://tstone-1.github.io/xlsxturbo/stability/) is the full statement
— the public surface named exhaustively, what does and does not count as breaking, and the
supported Python and platform matrices.

- Tested in CI on Python 3.10 and 3.12 across Linux, Windows, and macOS, plus Python 3.14
  on Linux. One `abi3` wheel per platform serves 3.10 through 3.14. Python 3.9 was dropped
  in 1.1.0; a 3.9 interpreter resolves to 1.0.0, which stays on PyPI.
- Advanced Excel features are exposed through focused parameters rather than a full
  workbook object model. That is a deliberate scope boundary, not a gap to be filled.

## Built on rust_xlsxwriter

Every byte of the `.xlsx` files xlsxturbo produces is written by
[**rust_xlsxwriter**](https://github.com/jmcnamara/rust_xlsxwriter), John McNamara's Rust
Excel writer (MIT licensed). It is the one substantial dependency, and it is not an
implementation detail you can ignore:

- **What xlsxturbo can do is bounded by what rust_xlsxwriter can do.** The features on the
  [capability matrix](https://tstone-1.github.io/xlsxturbo/capability-matrix/) are the ones
  it exposes; xlsxturbo's job is type detection, the DataFrame and CSV pipeline, option
  validation, and a Python API — not the XLSX format itself.
- **Bugs in the generated file are usually upstream.** When one is, it is reported to
  rust_xlsxwriter rather than papered over here — most recently
  [#185](https://github.com/jmcnamara/rust_xlsxwriter/issues/185), filed 2026-08-15 and
  fixed in 0.98.1 the next morning. A file Excel refuses to open is still worth
  [reporting to us](CONTRIBUTING.md); we will trace it and take it upstream.
- **The `xlsxwriter` in the benchmark table is a different project** — that is
  [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), the pure-Python library by the same
  author, and one of the things xlsxturbo is measured against.

If you write Excel files from Rust, use rust_xlsxwriter directly; xlsxturbo exists to put
it behind a Python DataFrame API. rust_xlsxwriter's notice, and those of every other
crate compiled into the wheel, are in
[THIRD-PARTY-LICENSES.md](THIRD-PARTY-LICENSES.md), which is generated from the
dependency tree rather than maintained by hand.

## Contributing

Setup takes about five minutes and is described in [CONTRIBUTING.md](CONTRIBUTING.md),
along with the exact lint, type, and test commands CI runs. Security reports go through
the process in [SECURITY.md](SECURITY.md).

## License

MIT — see [LICENSE](LICENSE).

The wheel contains compiled code from the Rust crates listed above, so their notices
travel with it: [THIRD-PARTY-LICENSES.md](THIRD-PARTY-LICENSES.md), also installed as
`xlsxturbo-<version>.dist-info/licenses/THIRD-PARTY-LICENSES.md`. It is generated from the
dependency tree by `python scripts/gen_third_party_licenses.py --write`; do not edit it by
hand.
