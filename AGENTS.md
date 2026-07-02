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
- Do not use unrelated work or organization accounts in this repository.

## Build, Test, and Release

- Use `uv` for Python dependency and command execution.
- Standard local checks: `cargo fmt --check`, `cargo clippy -- -D warnings`, `cargo test`, `maturin develop --release`, then `pytest tests/`.
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.

## Adding a Feature - the 7-Touchpoint Checklist

1. `src/types.rs` - add the field to the `define_options!` macro list (generates ExtractedOptions/EffectiveOpts/as_effective/merge_with) AND the matching field on the hand-written `SheetConfig` struct. A missing SheetConfig field is a compile error in the generated merge_with.
2. `src/extract.rs` - add `extract_<feature>()`; register it in `extract_sheet_info` via the `extract_dict_field!`/`extract_list_field!` macro and add the option name to `SHEET_OPTION_NAMES` (guard test enforces this).
3. `src/lib.rs` - add the field to `RawOptions` + `extract_options()`, and the kwarg + `#[pyo3(signature)]` entry + docstring to BOTH `df_to_xlsx` and `dfs_to_xlsx`.
4. `src/apply/<family>.rs` (+ re-export in the `src/apply.rs` facade) - `apply_<feature>()` with unknown-key validation (allowlist + error listing valid keys) and context-rich errors (`format!("<feature>['{}']: ...", cell_ref)`). Reuse `types::extract_opt` for field extraction.
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

## Benchmarks

- The main comparison suite is `benchmarks/benchmark.py`; use `--markdown` to regenerate the README performance table and `--json` for machine-readable output.
- The parallel CSV conversion suite is `benchmarks/benchmark_parallel.py`.
- README performance numbers are system-specific and should identify the machine, OS, Python version, and run methodology.
