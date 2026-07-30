# Contributing to xlsxturbo

Thanks for your interest. xlsxturbo is a Rust extension module for Python, so the
setup has two halves; the whole thing should take about five minutes.

## Setup

You need a [Rust toolchain](https://rustup.rs/) (stable), Python 3.9 or newer, and
[uv](https://docs.astral.sh/uv/).

```bash
git clone https://github.com/tstone-1/xlsxturbo
cd xlsxturbo

# This repo uses a project-local .venv on purpose.
uv venv
uv pip install -e ".[dev]"

# Build the Rust extension into the venv.
maturin develop --release

# Confirm it worked.
python -c "import xlsxturbo; print(xlsxturbo.version())"
```

`--release` matters. A debug build of the extension is 20-50x slower, which makes
the test suite crawl and makes any timing you look at meaningless.

## Checks

Run these before opening a pull request. They are the same checks CI runs, and CI
runs them with these exact flags -- `cargo clippy` without `--all-targets -- -D warnings`
is a weaker check than the gate.

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test
maturin develop --release
pytest tests/ -q
```

Python lint, type and security gates:

| Gate | Windows | macOS / Linux |
|------|---------|---------------|
| ruff | `.venv\Scripts\ruff.exe check python tests benchmarks` | `.venv/bin/ruff check python tests benchmarks` |
| bandit | `.venv\Scripts\bandit.exe -c pyproject.toml -r python` | `.venv/bin/bandit -c pyproject.toml -r python` |
| pyright | `.venv\Scripts\pyright.exe` | `.venv/bin/pyright` |

If a tool is missing, install the dev extras rather than reaching for a system copy:
`uv pip install -e ".[dev]"`. A stale extension in the venv is the usual cause of
confusing signature mismatches -- rebuild with `maturin develop --release` and check
that the interpreter in pytest's header is the project `.venv`.

If you change the `dev` dependencies, run `uv lock` -- the lockfile is tracked.

## Adding a feature: the touchpoint checklist

Options in this library thread through several layers, and missing one is usually a
compile error rather than a silent bug -- but not always. The full checklist, with the
reasoning behind each step, is in [AGENTS.md](AGENTS.md#adding-a-feature---the-7-touchpoint-checklist).
In short:

1. `src/types.rs` -- the `define_options!` list and the `SheetConfig` field. Use
   `IndexMap`, not `HashMap`, for anything keyed by cell reference: iteration order
   feeds straight into the generated XML, so a `HashMap` makes output non-reproducible.
2. `src/extract.rs` -- an `extract_<feature>()`, registered in `extract_sheet_info`,
   plus the option name in `SHEET_OPTION_NAMES` (a guard test enforces this).
3. `src/lib.rs` -- the `RawOptions` field, `extract_options()`, and the kwarg,
   `#[pyo3(signature)]` entry and docstring on **both** `df_to_xlsx` and `dfs_to_xlsx`.
4. `src/apply/<family>.rs` -- an `apply_<feature>()` with unknown-key validation and
   context-rich errors. Use `types::OptionMap` accessors rather than hand-rolling a new
   field-parsing wrapper family.
5. `src/convert.rs` -- the apply call in `apply_worksheet_features` (order matters:
   `cells` stays last), and a `constant_memory` classification decision. A guard test
   forces that decision.
6. `python/xlsxturbo/xlsxturbo.pyi` -- the option `TypedDict` and the kwarg on both
   signatures and on `SheetOptions`.
7. `tests/test_<feature area>.py` -- a `TestXxx` class that reads the produced `.xlsx`
   back via openpyxl or raw XML. Assertions about behaviour, not about internals.

Also update `README.md` and `CHANGELOG.md`.

## Tests

Tests read the generated workbook back rather than asserting on internal state. When
you add one, satisfy yourself that it can actually fail -- break the code on purpose,
one edit at a time, and check that your test is among what goes red. A test that
passes because it asserts something true by construction is indistinguishable from a
test that works.

## Pull requests

- Target `main`.
- Keep the change focused; unrelated cleanups are much easier to review separately.
- Say what you verified and on which platform. "Tests pass on Linux" is more useful
  than "tests pass".
- New behaviour needs a `CHANGELOG.md` entry under `## [Unreleased]`.

CI runs the Rust suite, the Python suite on Linux/Windows/macOS across several Python
versions, the lint gates, `cargo audit`, `pip-audit`, CodeQL and dependency review.

## Reporting bugs

Use the issue forms -- they ask for platform, Python version and a minimal input,
which are the three things that otherwise cost a round trip. For a security issue,
see [SECURITY.md](SECURITY.md); please don't open a public issue for it.

## Licence

By contributing you agree that your contribution is licensed under the MIT Licence,
as in [LICENSE](LICENSE).
