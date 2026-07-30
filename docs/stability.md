# Stability and support

From 1.0.0 onward, xlsxturbo follows [Semantic Versioning](https://semver.org/). This page
says exactly what that covers — which names are promised, what counts as a breaking change,
how long a deprecation lasts, which Pythons and platforms are supported, and what is
guaranteed about the `.xlsx` files themselves.

The short version: **anything you can reach from `import xlsxturbo` without a leading
underscore is covered, and it will not break until 2.0.0.**

## The public surface

Exactly these names, and nothing else:

| Name | Kind | Promised |
|------|------|----------|
| `df_to_xlsx` | function | Name, positional parameters, every keyword argument, and what each one does |
| `dfs_to_xlsx` | function | As above |
| `csv_to_xlsx` | function | As above |
| `version()` / `__version__` | function / string | Keeps returning the installed version |
| `ExportOptions` | dataclass | Field names, defaults, `as_kwargs()`, `as_sheet_options()` |
| `XlsxTurboError` and its subclasses | exceptions | Class names, the hierarchy, and which failure raises which — see [Errors and warnings](errors.md) |
| `xlsxturbo.types` | module | The option `TypedDict`s and `Literal`s, as type annotations |

Everything else is internal and may change in any release:

- The compiled `xlsxturbo.xlsxturbo` submodule. Import from the package, not from it.
- Any name beginning with `_`.
- Exception **message text**. Messages are written to be read by a person and are improved
  freely; match on the class, or on `errno` for `FileError`, never on the string.
- The internal layout of the generated XML — see [Generated files](#generated-files) below.

The `xlsxturbo` command-line binary is not part of the surface either, because it is not part
of the wheel: `cargo build --release` produces it, but a `pip install` does not.

## What counts as a breaking change

Breaking — 2.0.0 only, and never without a deprecation period:

- Removing or renaming any name in the table above
- Removing a keyword argument, or narrowing the values one accepts
- Changing which exception a given failure raises, including its builtin base class
- Changing the cell **value, type, or number format** produced for an input that already
  worked

Not breaking — these can land in a minor or a patch:

- A new keyword argument, option key, or exception subclass. Everything in this library is
  additive by design; the [7-touchpoint checklist](https://github.com/tstone-1/xlsxturbo/blob/main/AGENTS.md)
  exists so that stays true.
- Accepting an input that previously raised.
- A new `RuntimeWarning` — for example, a further option that `constant_memory=True` cannot
  apply.
- Rewording any error or warning message.
- Performance, in either direction, and memory use.
- Dropping a Python version that upstream has already end-of-lifed (see below).

That last one is the only genuine judgement call on the list, and it is stated because
leaving it unstated is how a project ends up either never dropping a version or dropping one
in a patch release.

## Deprecation policy

Nothing in the public surface is removed without warning first. Concretely:

1. The replacement lands first, so there is never a release where the old way is deprecated
   and the new way does not exist yet.
2. Using the old way emits a `DeprecationWarning` naming **what to use instead** and **the
   version it will be removed in**.
3. That warning ships for **at least one minor release and at least six months**, whichever
   is longer.
4. Removal happens only in a major release.

So the earliest anything deprecated in 1.1.0 can disappear is 2.0.0, and only if six months
have passed. To find deprecations early in your own test suite:

```python
import warnings

warnings.simplefilter("error", DeprecationWarning)
```

## Supported Python versions

One `abi3` wheel per platform covers every supported version, so support here is structural
rather than per-version: the wheel built against Python 3.10's stable ABI is the same file a
3.14 interpreter loads.

| Python | Supported | Run in CI |
|--------|-----------|-----------|
| 3.10 | yes | yes |
| 3.11 | yes | no |
| 3.12 | yes | yes |
| 3.13 | yes | no |
| 3.14 | yes | yes |

The versions marked "no" are not untested by oversight — they run the identical wheel through
the identical `abi3` interface, so CI covers the oldest, a middle, and the newest, and the
gaps carry no independent risk. What CI would catch on 3.11 that 3.10 and 3.12 do not is
essentially nothing.

**Python 3.9 was dropped in 1.1.0**, having reached upstream end of life in October 2025.
`pip` handles this without any action on your part: a 3.9 interpreter resolves to 1.0.0,
which stays on PyPI and keeps working. Move to 3.10 or newer to receive further releases.

A dropped Python version is a **minor** release, not a 2.0.0 event. An interpreter whose
upstream support has ended is not a platform this project can meaningfully promise anything
about, and holding the major version hostage to it would mean either never dropping one or
bundling unrelated breakage to justify the bump. The cost of *not* dropping it is not
theoretical: in the two days before 1.1.0, the 3.9 floor blocked pytest 9, numpy 2.1+ and
polars 1.37+ from the test matrix, and forced `python/xlsxturbo/types.py` to spell every
union `Union[str, int]` where `str | int` is the natural form.

## Supported platforms

Wheels are built for these targets on every release:

| Platform | Architecture | Wheel tag | Smoke-tested before publish |
|----------|--------------|-----------|------------------------------|
| Linux | x86_64 | `manylinux_2_28` | yes |
| Linux | aarch64 | `manylinux_2_28` | no |
| macOS | x86_64 | — | no |
| macOS | aarch64 (Apple silicon) | — | yes |
| Windows | x64 | — | yes |

An sdist is published alongside them, so a platform without a wheel — Windows on ARM, musl
Linux, FreeBSD — can still install by building from source, which needs a Rust toolchain.
Those builds are not tested here and carry no promise.

"Smoke-tested" means the built wheel is installed on a clean runner and the full test suite
runs against it **before** anything is published to PyPI. The two that are not are
cross-compiled for an architecture no hosted runner provides; they are built from the same
sources by the same toolchain in the same job matrix.

`manylinux_2_28` is a deliberate floor — glibc 2.28, so RHEL 8 and Debian 10 and newer. It is
chosen rather than left automatic; see the note in
[AGENTS.md](https://github.com/tstone-1/xlsxturbo/blob/main/AGENTS.md).

## Generated files

The output is a standard `.xlsx` — an OOXML SpreadsheetML package readable by Excel 2007 and
later, LibreOffice Calc, Numbers, Google Sheets, and library readers such as `openpyxl`,
`pandas.read_excel` and `polars.read_excel`.

**What is promised**: for the same input and options, within a major version, every cell has
the same value, the same type, and the same number format. That is the guarantee worth
depending on, and the one the test suite checks — it reads generated files back and asserts
on cells, not on bytes.

**What is not promised**: the byte content of the file. The internal XML — element order,
whitespace, how styles are shared, which optional parts are present — belongs to the writer
and changes when it is upgraded.

One consequence is worth stating plainly, because it is easy to discover the hard way:

!!! warning "Two exports of identical data are not identical files"

    `docProps/core.xml` records the moment the workbook was created, so hashing the output
    to detect "did anything change?" reports a change every time. Every other part of the
    archive **is** byte-identical across runs — measured, and pinned by a test — so compare
    at the level you actually care about: read both files back and compare cell values, or
    compare the archive members other than `docProps/core.xml`.

## Reporting a compatibility problem

A file Excel refuses to open, a cell whose value changed between versions without a changelog
entry, or a platform where the wheel does not load — those are bugs, and the most valuable
kind of report. See
[CONTRIBUTING.md](https://github.com/tstone-1/xlsxturbo/blob/main/CONTRIBUTING.md), and
include `xlsxturbo.version()` plus the calling code.
