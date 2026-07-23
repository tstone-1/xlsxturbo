# xlsxturbo Agent Instructions

## Shared Memory Policy

- `AGENTS.md` is the canonical shared memory for Codex and other coding agents in this repository.
- Claude Code loads this file through `.claude/CLAUDE.md`.
- Durable project knowledge, workflows, commands, architecture notes, and recurring pitfalls belong here.
- Do not store durable project knowledge only in Claude auto memory.
- Keep entries concise, specific, and verifiable. Prefer updating existing sections over appending duplicate notes.

## Git Workflow

- Only commit and push when explicitly asked by the user.
- Do not include Claude-related or AI-generated footers in commit messages.
- Before commit or push, run `cargo update` to check for Rust dependency updates.
- Follow `BUILD.md` before release or push-ready work.

## Account Enforcement

- Before any commit, run: `git config user.email "48162401+tstone-1@users.noreply.github.com"` and `git config user.name "tstone-1"`.
- Before any push, run: `gh auth switch --user tstone-1`.
- In multi-push flows (branch push + tag push), re-run the auth switch and verify with `gh auth status` before EACH push: a local shell profile can re-assert a different active account between commands (observed 2026-04-18 — the tag push failed after the branch push succeeded).
- Do not use unrelated work or organization accounts in this repository.

## Build, Test, and Release

- Use `uv` for Python dependency and command execution.
- This repo uses a project-local `.venv` (an exception to any central-venv convention). Test deps (`pytest pandas polars openpyxl`) must be installed there; if they are missing, `uv run pytest` silently falls back to a system Python with a stale extension and reports bogus signature mismatches. Verify the interpreter in the pytest header is `.venv\Scripts\python.exe`; recover with `uv pip install pytest pandas polars openpyxl` and rebuild via `maturin develop --release`.
- Standard local checks: `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`, `cargo test`, `maturin develop --release`, then `pytest tests/`.
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.
- If multiple version-bump commits are awaiting release, tag each released version at its own commit; do not collapse distinct versions into one tag.
- Linux release wheels intentionally use `manylinux_2_28` with maturin's `--find-interpreter` and PyO3 `abi3-py39`. Do not switch back to automatic manylinux selection without verifying Python 3.9+ abi3 wheels.
- The release workflow must smoke-test the built Linux, Windows, and macOS wheels before publishing to PyPI.
- To confirm a release on PyPI, query the version-specific endpoint `https://pypi.org/pypi/xlsxturbo/<version>/json` (authoritative within seconds). The aggregate `https://pypi.org/pypi/xlsxturbo/json` `info.version` field lags several minutes behind (CDN cache) and can still show the previous version; trust the publish job's green status over it.
- For multi-phase implementation work, run a deep diff review after each completed, verified phase before building the next phase on top of it.

## Adding a Feature - the 7-Touchpoint Checklist

1. `src/types.rs` - add the field to the `define_options!` macro list (generates ExtractedOptions/EffectiveOpts/as_effective/merge_with) AND the matching field on the hand-written `SheetConfig` struct. A missing SheetConfig field is a compile error in the generated merge_with. A cell_ref/location-keyed feature map (images, charts, comments, ...) must be `IndexMap`, not `HashMap` — iteration order feeds straight into the generated XML, so a `HashMap` makes output non-reproducible across runs.
2. `src/extract.rs` - add `extract_<feature>()`; register it in `extract_sheet_info` via the `extract_dict_field!`/`extract_list_field!` macro and add the option name to `SHEET_OPTION_NAMES` (guard test enforces this). Two extraction patterns coexist by design: simple structures (column_widths, formula_columns, merged_range tuples) are eagerly typed into real Rust types here, at extract time; features whose parsing needs a `py`/rust_xlsxwriter type (a `Format`, a `Color`, a chart/sparkline builder) instead extract only a raw `HashMap<String, Py<PyAny>>` "blob" here and defer real validation to the matching `apply/*` function, since that parsing can't happen without the GIL-bound types apply time has. Don't "fix" a blob extractor by eagerly typing it — that's the wrong layer for that feature.
3. `src/lib.rs` - add the field to `RawOptions` + `extract_options()`, and the kwarg + `#[pyo3(signature)]` entry + docstring to BOTH `df_to_xlsx` and `dfs_to_xlsx`.
4. `src/apply/<family>.rs` (+ re-export in the `src/apply.rs` facade) - `apply_<feature>()` with unknown-key validation and context-rich errors (`format!("<feature>['{}']: ...", cell_ref)`). For a blob-extracted feature, build one `types::OptionMap::new(py, &blob, context)` per config and use its typed accessors (`.string()`, `.bool()`, `.f64()`, `.u32()`, `.dict()`, `.reject_unknown()`/`.reject_unknown_for()`) instead of hand-rolling a new `<feature>_string_field`-style wrapper family — that duplication (~400 lines across charts/sparklines/validations/media/conditional_formats/format-dict parsing) is exactly what `OptionMap` replaced.
5. `src/convert.rs` - apply call in `apply_worksheet_features` (order matters: `cells` stays last so user cells can overwrite data). Decide constant_memory classification: a new option defaults to skipped+warned; add to `CONSTANT_MEMORY_SAFE_OPTIONS` only if applied during the data write. The guard test `every_complex_option_is_classified_for_constant_memory` forces this decision.
6. `python/xlsxturbo/xlsxturbo.pyi` - TypedDict for the options, kwarg on both signatures and `SheetOptions`, docstrings. This compiled-extension stub is the type source of truth; `__init__.pyi` is a thin re-export - never hand-edit it for new options.
7. `tests/test_<feature area>.py` - a `TestXxx` class following the existing per-feature test files (behavior-coupled: read the produced xlsx back via openpyxl or XML).

## Python Lint, Type, and Security Gates

The Python tree (`python/`, `tests/`, `benchmarks/`) must stay clean under ruff, bandit, and pyright, with docstrings and type annotations on all functions. Config lives in `pyproject.toml`; the tools are in the `dev` optional-deps. These same three gates also run in CI (`python-lint` job in `.github/workflows/ci.yml`). Run from the repo root using the project-local `.venv`:

- `.venv/Scripts/ruff.exe check python tests benchmarks`
- `.venv/Scripts/bandit.exe -c pyproject.toml -r python`
- `.venv/Scripts/pyright.exe`
- `.venv/Scripts/python.exe -m pytest tests/ -q`

Scoping notes (intentional, do not "fix" by widening):
- pyright runs `typeCheckingMode = "standard"` project-wide, with the shipped library raised to strict via the top-level `strict = ["python/xlsxturbo"]` path list. Do not use `executionEnvironments` + `typeCheckingMode` for this — that key is silently ignored by pyright 1.1.x.
- bandit scans `python/` only; tests and benchmarks are excluded (asserts and non-crypto `random` data generation are expected there).
- ruff per-file-ignores: `S101` in tests; `S404/S603/S607/S311/T201` in benchmarks. Google docstring convention.
- When changing the `dev` deps, run `uv lock` (the lockfile is tracked).
- pandas-stubs rejects `pd.to_datetime([..., pd.NaT, ...])` (mixed `list[str | NaTType]`); use the string `"NaT"` instead — pandas parses it to NaT, keeping test data identical.

## Benchmarks

- The main comparison suite is `benchmarks/benchmark.py`; use `--markdown` to regenerate the README performance table and `--json` for machine-readable output.
- The parallel CSV conversion suite is `benchmarks/benchmark_parallel.py`.
- README performance numbers are system-specific and should identify the machine, OS, Python version, and run methodology.
- Keep comparisons reproducible and fair: seed generated data, use native-fast dtypes for every compared library, perform warmup runs, report medians and standard deviations, and keep both benchmark suites methodologically aligned.
- Prefer honest, reproducible results over flattering headline numbers, including when a fairer method reduces the reported speedup.
- Generate measured documentation claims (benchmark results, variance, counts, and similar values) from their source script when practical; avoid hand-maintained factoids that silently become stale.
