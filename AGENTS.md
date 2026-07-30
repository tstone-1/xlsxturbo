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
- **A new third-party import in a test cannot be validated locally — add it to `requirements-test.txt`.** The local `.venv` holds the whole `dev` extras, but the CI test jobs and the release smoke test install only `requirements-test.txt`, so the local environment is a strict superset of CI and a test importing anything outside that file passes locally and fails only in CI. Declaring it in `[project.optional-dependencies] dev` does **not** fix it; those jobs never install `dev`. `tests/test_ci_config.py` fails if an import is undeclared or if a workflow re-inlines the list.

  That guard exists because the same bug landed three times in two days: `tests/test_docs_site.py` imported `yaml` (declared in `dev`, broke three CI jobs); fixing those three left a **fourth** copy of the list in `release.yml`'s smoke-test job, which failed the v0.19.0 release after every wheel had already built; and the guard's first run found `numpy` imported by `tests/test_core.py` and never declared anywhere — working only because pandas pulls it in. A comment saying "remember the other copies" was the fix after the first, and it did not survive a day.

- **The release smoke test runs `pytest tests/` against an installed wheel from outside the checkout, with only `tests/` copied.** So a test module that reads *repository* files — `mkdocs.yml`, the capability matrix, the generator script, the workflows — has nothing to read there and must skip via `tests.helpers.repo_checkout_available()`, not fail. This is invisible in ordinary CI and surfaces only when a tag is pushed: `test_docs_site.py` and `test_capability_matrix.py` were both added after v0.18.0, so v0.19.0 was the first release to run them and all 16 failed. Inside a checkout a missing file stays a hard failure — the guard distinguishes "no repository here" from "the repository is broken". Reproduce the job locally before tagging: `cp -r tests "$TMP/smoke" && cd "$TMP/smoke" && python -m pytest tests/ -q`.
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
6. **`python/xlsxturbo/types.py`** - the option `TypedDict`/`Literal` for the new feature. Since 0.19.0 this runtime module, NOT the stub, is where shapes are declared; `xlsxturbo.pyi` imports them with the redundant-alias form (`X as X`) and keeps only the four function signatures, the exception classes and `__version__`. So a new option means: add the shape to `types.py`, add `X as X` to the stub's import block **and** to its `__all__`, then add the kwarg to both function signatures and to `SheetOptions` (which is itself in `types.py`). `tests/test_types_module.py` fails if the stub's re-export list and the runtime module disagree. `__init__.pyi` is a thin re-export of the *runtime* surface - never hand-edit it for new options. Two constraints on `types.py`: module-level aliases must use `Union[...]`/`Optional[...]` (`requires-python` is `>=3.9`, where `str | PathLike[str]` raises at import), while field annotations inside a `TypedDict` may use `|` because the module has `from __future__ import annotations`. Verified against a real 3.9 interpreter; the trade is that `typing.get_type_hints()` on these classes fails on 3.9 and works from 3.10.
7. `tests/test_<feature area>.py` - a `TestXxx` class following the existing per-feature test files (behavior-coupled: read the produced xlsx back via openpyxl or XML).

Raising from Rust: **never `pyo3::exceptions::Py*Error::new_err` in `src/`.** Use the
`crate::errors::*` helpers — see the section below. `src/apply/` and `src/parse/` are
unaffected: they return `Result<_, String>` and the boundary classifies for them.

Then regenerate the capability matrix: `python scripts/gen_capability_matrix.py --write`. `docs/capability-matrix.md` is GENERATED from the Rust sources and must never be hand-edited; `tests/test_capability_matrix.py` fails if the committed page is stale. The generator parses `SHEET_OPTION_NAMES`, `define_options!`, `CONSTANT_MEMORY_SAFE_OPTIONS`, `warn_constant_memory_skips` and the three `#[pyo3(signature)]` blocks, so touching any of those changes the page. Each parser raises rather than returning an empty list, because a structural audit that matches nothing reads exactly like a clean result — and each parsed parameter must be a Python identifier, which is what caught a regex that spanned from the file's first pyo3 attribute through to the requested function and produced "parameters" like `) -> PyResult<(u32`.

## The Exception Hierarchy (0.19.0+)

`src/errors.rs` owns the public exception classes and is the only place in `src/` that may
construct a `PyErr` from scratch. Raise with `errors::configuration`,
`errors::configuration_type`, `errors::input_data`, `errors::file` or
`errors::workbook_validation`. The classes are built by calling the `type` metaclass rather
than with `create_exception!`, because that macro takes a single base and every class here
needs two or three.

Five facts that are expensive to rediscover:

- **The second base is a compatibility contract, not decoration.** Each class inherits the
  builtin its failures raised in 0.18. Pick a new class's builtin by **grepping what the site
  raises today**, not by what the failure morally is. Getting this from taste produced a
  breaking change twice during the original implementation — once for I/O (`OSError` alone
  would have broken `except ValueError`) and once for `InputDataError` (which is a
  `ValueError`, because frame detection has always reached the boundary through the pipeline).
- **The 93 `pytest.raises(ValueError|TypeError)` assertions across `tests/` are the
  behaviour record for pre-0.19.** They are the compatibility gate. A change to this area that
  needs one of them edited to stay green is a breaking change, whatever the changelog claims.
- **A class no site raises is dead API that can never be removed.**
  `tests/test_errors.py` asserts the exported set equals the set with a working trigger, which
  is what kept `UnsupportedFeatureError` from shipping (a `constant_memory` conflict is a
  `RuntimeWarning` and the call succeeds — nothing would have raised it).
- **The boundary is `src/lib.rs` *and* `src/extract.rs`,** 50 sites, not the dozen the plan
  assumed. Everything below `extract.rs` is uniformly `Result<_, String>`.
- **`ConvertError` in `src/convert.rs` is the one seam.** Two variants, `Config` and `File`,
  because `save_workbook` runs inside the pipeline and its failure would otherwise be
  indistinguishable from a bad option at the boundary. `From<String>` maps to `Config`, so
  every existing `?` still compiles — which also means **a new filesystem call defaults to
  `Config` unless you tag it `ConvertError::File`.** Frame detection needed the same treatment
  and got a boundary call (`require_supported_dataframe`) instead, since tagging it would have
  pushed `ConvertError` down into the write layer.

Full reasoning, including the shapes considered and rejected: `docs/roadmap-1.0.md`
decision D6. User-facing contract: `docs/errors.md`.

## Python Lint, Type, and Security Gates

The Python tree (`python/`, `tests/`, `benchmarks/`, `scripts/`) must stay clean under ruff, bandit, and pyright, with docstrings and type annotations on all functions. Config lives in `pyproject.toml`; the tools are in the `dev` optional-deps. These same three gates also run in CI (`python-lint` job in `.github/workflows/ci.yml`). Run from the repo root using the project-local `.venv`:

On Windows the venv's executables live in `.venv\Scripts\`, on macOS/Linux in `.venv/bin/` — this repo is worked on from both, so use the pair for the machine you are on:

| Gate | Windows | macOS / Linux |
|------|---------|---------------|
| ruff | `.venv\Scripts\ruff.exe check python tests benchmarks scripts` | `.venv/bin/ruff check python tests benchmarks scripts` |
| bandit | `.venv\Scripts\bandit.exe -c pyproject.toml -r python` | `.venv/bin/bandit -c pyproject.toml -r python` |
| pyright | `.venv\Scripts\pyright.exe` | `.venv/bin/pyright` |
| pytest | `.venv\Scripts\python.exe -m pytest tests/ -q` | `.venv/bin/python -m pytest tests/ -q` |

If a tool is missing from the venv, install the dev extras (`uv pip install -e ".[dev]"`) rather than reaching for a system copy; `uvx <tool>` also works for a one-off check and needs no venv.

Scoping notes (intentional, do not "fix" by widening):
- pyright runs `typeCheckingMode = "standard"` project-wide, with the shipped library raised to strict via the top-level `strict = ["python/xlsxturbo"]` path list. Do not use `executionEnvironments` + `typeCheckingMode` for this — that key is silently ignored by pyright 1.1.x.
- bandit scans `python/` only; tests and benchmarks are excluded (asserts and non-crypto `random` data generation are expected there).
- ruff per-file-ignores: `S101` in tests; `S404/S603/S607/S311/T201` in benchmarks. Google docstring convention.
- When changing the `dev` deps, run `uv lock` (the lockfile is tracked).
- pandas-stubs rejects `pd.to_datetime([..., pd.NaT, ...])` (mixed `list[str | NaTType]`); use the string `"NaT"` instead — pandas parses it to NaT, keeping test data identical.

## Benchmarks

- The main comparison suite is `benchmarks/benchmark.py`; use `--markdown` to regenerate the tables on `docs/performance.md` and `--json` for machine-readable output. (Those tables were on the README until the 0.19.0 documentation split; the README now carries only the headline ratios in prose.)
- The parallel CSV conversion suite is `benchmarks/benchmark_parallel.py`.
- Published performance numbers are system-specific and must identify the machine, OS, Python version, and run methodology.
- Keep comparisons reproducible and fair: seed generated data, use native-fast dtypes for every compared library, perform warmup runs, report medians and standard deviations, and keep both benchmark suites methodologically aligned.
- Prefer honest, reproducible results over flattering headline numbers, including when a fairer method reduces the reported speedup.
- Generate measured documentation claims (benchmark results, variance, counts, and similar values) from their source script when practical; avoid hand-maintained factoids that silently become stale.

## Documentation Sync

Migrated here from the personal cross-repo memory file on 2026-07-30: it was xlsxturbo-specific knowledge living outside the repo, and its first step had gone stale (it said usage examples belong in the README, which stopped being true with the 0.19.0 split).

When adding or modifying a feature:

1. **The relevant `docs/` page** — add or update the usage example. `docs/` is the MkDocs Material site published to GitHub Pages; the README is a landing page and should NOT grow new per-feature examples. Match the option to its page from the nav in `mkdocs.yml` (formatting, tables, charts-and-media, cells, ...). A brand-new page must be added to that nav or `tests/test_docs_site.py` fails it as unreachable.
2. **`docs/capability-matrix.md`** — regenerate with `python scripts/gen_capability_matrix.py --write`. Never hand-edit it. `tests/test_capability_matrix.py` fails when the committed page is stale.
3. **`CHANGELOG.md`** — document all changes under the appropriate version heading. Note the file is NOT uniform: 37 headings are `## [X.Y.Z]` and three older ones are `## X.Y.Z` without brackets. `.github/scripts/release-notes.sh` handles both by fixed-string matching; do not "simplify" it to a regex, which is how a release once got the wrong version's notes.
4. **Type stubs** (`python/xlsxturbo/xlsxturbo.pyi`) — add new parameters with types and docstrings. This compiled-extension stub is the type source of truth; `python/xlsxturbo/__init__.pyi` is a thin re-export and must not be hand-edited for new options.

Before commit or push, follow the `BUILD.md` checklist.

### docs/ is published wholesale — mind what is sitting in it

MkDocs publishes every file under `docs/`, and knows nothing about git. Two tracked-but-internal files (`roadmap-1.0.md`) and two untracked ones (`strategic-recommendations-plan.md`, `reviews/`) are therefore listed in `mkdocs.yml`'s `exclude_docs`, and `tests/test_docs_site.py` asserts that list stays in step with `.gitignore`.

Verified by removing one `exclude_docs` entry: the private planning memo appeared in `site/` and `mkdocs build --strict` still **exited 0**. Strict mode does not cover this. The reliable protection is that deployment happens only from `.github/workflows/docs.yml`, which builds from a clean checkout and so cannot see an untracked file at all — never run `mkdocs gh-deploy` by hand.

### The CLI is not in the wheel

`Cargo.toml` has a `[[bin]] xlsxturbo` target with `default = ["cli"]`, so `cargo build --release` produces a working CLI. maturin builds only the extension module, so the published wheel contains **no console script and no binary** — confirmed by inspecting the PyPI artifact for 0.18.0, not by reading the config. The README and the CLI docs claimed otherwise until 0.19.0. If shipping it is ever wanted, that is a packaging change, not a documentation one.
