# xlsxturbo Agent Instructions

## Shared Memory Policy

- `AGENTS.md` is the canonical shared memory for Codex and other coding agents in this repository.
- Claude Code loads this file itself; there is no `CLAUDE.md`, and adding one stops that unless it imports `@AGENTS.md`.
- Durable project knowledge, workflows, commands, architecture notes, and recurring pitfalls belong here.
- Do not store durable project knowledge only in Claude auto memory.
- Keep entries concise, specific, and verifiable. Prefer updating existing sections over appending duplicate notes.

## Git Workflow

- Only commit and push when explicitly asked by the user.
- Do not include Claude-related or AI-generated footers in commit messages.
- Before commit or push, run `cargo update` to check for Rust dependency updates, and confirm
  with `git diff --stat -- Cargo.lock` that the lockfile changed whenever cargo printed
  `Updating` lines (on 2026-09-30 `cargo update` printed 55 updates and exited 0 three times in a row without writing the lockfile) — **and
  regenerate `THIRD-PARTY-LICENSES.md` afterwards** (`python scripts/gen_third_party_licenses.py --write`).
  The notice lists a version beside every crate, so a lock refresh makes it stale; the 1.3.0
  release commit refreshed the lock and shipped a notice naming the previous version of five
  crates — `rust_xlsxwriter 0.98.2` in the release that moved it to 0.99.0 — and the guard
  compared names only until 2026-09-02. It compares versions now, so
  forgetting this fails `tests/test_third_party_licenses.py` rather than shipping.
- Follow `BUILD.md` before release or push-ready work.

## Account Enforcement

- Before any commit, run: `git config user.email "48162401+tstone-1@users.noreply.github.com"` and `git config user.name "tstone-1"`.
- Before any push, run: `gh auth switch --user tstone-1`.
- In multi-push flows (branch push + tag push), re-run the auth switch and verify with `gh auth status` before EACH push: a local shell profile can re-assert a different active account between commands (observed 2026-04-18 — the tag push failed after the branch push succeeded).
- Do not use unrelated work or organization accounts in this repository.

## Build, Test, and Release

- Use `uv` for Python dependency and command execution.
- *The test deps live in the project-local `.venv`; a pytest header naming another interpreter means they are missing* -- recover with `uv pip install -e ".[dev]" -r requirements-test.txt -r requirements-docs.txt`. Full text: [dev-docs/build-and-release.md](dev-docs/build-and-release.md#the-project-local-venv).
- *Never run a bare `uv run` here; always `uv run --no-sync`* -- a plain one syncs `.venv` to `uv.lock` and uninstalls pytest and the other dev tools. Full text: [dev-docs/build-and-release.md](dev-docs/build-and-release.md#never-run-a-bare-uv-run).
- *A new third-party import in a test goes in `requirements-test.txt`, not the `dev` extras* -- CI installs only that file, so the import passes locally and fails in CI. Full text: [dev-docs/build-and-release.md](dev-docs/build-and-release.md#a-new-third-party-import-in-a-test-goes-in-requirements-testtxt).

  *A declared version range is only supported at the end CI installs* -- pandas 2 and polars 1 are each covered only by a layered `requirements-test-<package><major>.txt` leg; `OLDER_MAJOR_LEGS` in `tests/test_ci_config.py` fails on a two-major range without one. Full text: [dev-docs/testing.md](dev-docs/testing.md).

- *The release smoke test runs `pytest tests/` outside the checkout* -- a test that reads repository files must skip via `repo_checkout_available()`; reproduce the job before tagging. Full text: [dev-docs/build-and-release.md](dev-docs/build-and-release.md#the-release-smoke-test-runs-outside-the-checkout).
- Standard local checks: `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`, `cargo test --release`, `maturin develop --release`, then `pytest tests/`, then `mkdocs build --strict` when anything under `docs/` moved (needs `uv pip install -r requirements-docs.txt` into the venv once; the `dev` extras deliberately do not carry it). `BUILD.md`'s Pre-Push Checklist is the CI-verbatim list; this line is the short form.
- *`tests/test_build_is_current.py` fails when the built extension is older than any Rust source, `Cargo.toml`, `Cargo.lock` or `pyproject.toml`* -- a stale `.so` measures the wrong code without erroring; run `maturin develop --release`. Full text: [dev-docs/testing.md](dev-docs/testing.md#a-stale-build-fails-the-suite).
- *Reading a workbook back is not opening it in Excel* -- openpyxl and the XML passed a one-bound `decimal` validation that Excel refused for ten months; check in Excel with `scripts/excel_open_check.ps1 -Control <good.xlsx> <files>`, which also prints the dialog a refused file shows. Full text: [dev-docs/testing.md](dev-docs/testing.md#reading-a-workbook-back-is-not-opening-it-in-excel).
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.
- If multiple version-bump commits are awaiting release, tag each released version at its own commit; do not collapse distinct versions into one tag.
- Linux release wheels intentionally use `manylinux_2_28` with maturin's `--find-interpreter` and PyO3 `abi3-py311`. Do not switch back to automatic manylinux selection without verifying Python 3.11+ abi3 wheels. (Raised from `abi3-py39` in 1.1.0 and `abi3-py310` in 1.13.0.)
- *The release workflow smoke-tests every published wheel on a runner of its own architecture* -- each leg asserts its wheel's platform tag; `tests/test_stability_policy.py` ties the build and smoke matrices to `docs/stability.md`. Full text: [dev-docs/testing.md](dev-docs/testing.md).
- *Confirm a release at `https://pypi.org/pypi/xlsxturbo/<version>/json`* -- the aggregate endpoint lags minutes behind. Full text: [dev-docs/build-and-release.md](dev-docs/build-and-release.md#confirming-a-release-on-pypi).
- For multi-phase implementation work, run a deep diff review after each completed, verified phase before building the next phase on top of it.

## Adding a Feature - the Touchpoint Checklist

Read [dev-docs/adding-a-feature.md](dev-docs/adding-a-feature.md) in full before adding an option; each line here is a pointer, not the instruction.

1. `src/types.rs` -- the field in `define_options!` and on `SheetConfig`; a location-keyed map is `IndexMap`, never `HashMap`.
2. `src/extract.rs` -- `extract_<feature>()` registered in `extract_sheet_info`, the name in `SHEET_OPTION_NAMES`; a blob extractor stays a blob.
3. `src/lib.rs` -- `RawOptions`, `extract_options()`, and kwarg, signature entry and docstring in BOTH `df_to_xlsx` and `dfs_to_xlsx`.
4. `src/apply/<family>.rs` -- `apply_<feature>()` with unknown-key validation, built on `types::OptionMap`.
5. `src/convert.rs` -- the call in `apply_frame_features`, `finish_sheet` or `apply_cell_options`; classify it for constant_memory and for several frames.
6. `python/xlsxturbo/types.py` -- the shape, `X as X` in the stub's imports and `__all__`, the kwarg in both signatures and `SheetOptions`.
7. `tests/test_<feature area>.py` -- a `TestXxx` class that reads the produced xlsx back.
8. `python/xlsxturbo/options.py` -- the `ExportOptions` field, annotated byte-identically to the stub, and a `SAMPLE_VALUES` entry.

- *A dict option is `Mapping[...]`, a list option `Sequence[...]`* -- reaching Rust through `types::mapping_as_dict` and `types::sequence_as_list`.
- *The capped autofit path never calls the writer's `autofit()`* -- allowances such as `FILTER_BUTTON_WIDTH` are added by hand in `write_frame`.
- *A sheet property runs on the constant-memory path too* -- tab colour, freeze panes and the autofilter are set after the data.
- *Several entries may name one sheet; each option is frame-bound, whole-sheet or cell-addressed* -- two guard tests force the choice.
- *Never `Py*Error::new_err` in `src/`* -- raise through `crate::errors::*`.
- *Then regenerate the capability matrix* -- `python scripts/gen_capability_matrix.py --write`; never hand-edit `docs/capability-matrix.md`.

## Upstream defects belong upstream — file them

- *When rust_xlsxwriter is what is wrong, file it at jmcnamara/rust_xlsxwriter alongside the local workaround* -- #185 and #186 were each fixed within a day of filing. Full text: [dev-docs/upstream-rust-xlsxwriter.md](dev-docs/upstream-rust-xlsxwriter.md#upstream-defects-belong-upstream--file-them).
- *A report lands with a reproducer that needs no Excel, controls in the same program, and only what was measured* -- review the draft against what you actually ran.
- *Pin an unreachable upstream defect in a test that fails when it is fixed* -- delete that test together with the workaround.
- *Closed: #213, `TempFileError` drops the `io::Error`* -- fixed in rust_xlsxwriter 0.101.0; `temp_file_failure` now maps the `io::Error`, so a refused temporary folder is `FilePermissionError`. No open upstream defect is pinned now.
- *Fixed in rust_xlsxwriter 0.103.0 (2026-11-18), not yet taken here: a sheet named `TRUE` or `FALSE` with `autofilter=True` gives a workbook Excel cannot open* -- upstream commit `0c034f85b`; no guard or test here yet. The 3-day publish-age rule holds 0.103.0 back until 2026-11-21 23:25 UTC.
- *Filed as #216 and #217, both fixed in 0.103.0, not yet taken here: three more ways rust_xlsxwriter 0.102.0 writes a workbook Excel cannot open* -- control characters in attributes, a tab or CR in a table column header, `_x000D_`-shaped names. Only the header case is guarded (`unwritable_table_header`); see the notes before removing it.
- *Filed as #221, fixed in 0.103.0: a failed `Workbook::save(path)` leaves a 0-byte file and replaces an existing workbook* -- xlsxturbo saves through `save_to_writer()` into a temporary file and renames it; keep that path.
- *Pivot tables shipped in 0.103.0 (`PivotTable`; PR #220 was closed, the feature landed separately): read the findings before exposing them* -- at `238bc5b` bad headers got Excel's repair prompt; `30147bb` handles them; a constant-memory source needs `set_field_names()`, wrong names there give an empty table, and placement is checked in the first column only.

## Name validation is ours now — upstream has drawn its line

Full text: [dev-docs/upstream-rust-xlsxwriter.md](dev-docs/upstream-rust-xlsxwriter.md#name-validation-is-ours-now--upstream-has-drawn-its-line). Read it before touching `reject_reference_shaped_name`, `sanitize_table_name`, `claimed_table_name` or a rust_xlsxwriter bump.

- *The crate's name checks are a foot-gun guard, not a specification* -- the two screens here are the layer xlsxturbo owns, not temporary workarounds.
- *Rename what nobody points at, refuse what somebody might* -- `reject_reference_shaped_name` refuses defined names; `sanitize_table_name` rewrites table names.
- *A table/defined-name collision is a pre-check, not a save failure* -- `claimed_table_name` (`src/convert.rs`) owns the gate both pre-checks use; any new save-time rule upstream is a pre-check candidate.
- *`sanitize_table_name` normalises to NFC, never NFKC, before the screen* -- Thai/Hindi marks are still rewritten on purpose; a denylist flip needs the mirror audit, not reading.
- *The allowlist must stay at least as wide as the crate's denylist* -- on every bump, account for each rule in `utility::check_name`; re-pin the empty-name message tests, do not reinstate the screen.

## Multi-frame name validation

- Automatic table IDs follow worksheet order, then frame order within each sheet,
  not global entry order. Keep name preflight in that same order; interleaved A/B/A
  entries must cover explicit names and defined-name collisions in regression tests.
- Conditional-format numeric thresholds must be finite in both single-value and range
  comparisons. Preserve the separate string-comparison path.
- Release SBOMs describe shipped Rust dependencies. The wheel has no Python runtime
  dependencies; inventorying the generator environment does not describe the wheel.

## Zip compression level

- *Zip compression level — measured before exposing it* -- exposed as `compression_level` in 1.12.0; the crate only warns on stderr for a level outside 1-9, so we validate it. Level 1 cut `df_to_xlsx` by 28%. Full text: [dev-docs/upstream-rust-xlsxwriter.md](dev-docs/upstream-rust-xlsxwriter.md#zip-compression-level--measured-before-exposing-it).

## The bundled license notice

- *`THIRD-PARTY-LICENSES.md` is generated; never edit it* -- `python scripts/gen_third_party_licenses.py --write`; `tests/test_third_party_licenses.py` checks it both ways.
- *`cargo install cargo-about` installs nothing and exits 0* -- use `--features cli`; the generator passes `--output-file` on purpose.
- *A Dependabot cargo PR for a shipped crate fails CI on the notice test, by construction* -- it cannot regenerate the notice; take the update on `main` as `BUILD.md` step 5 describes.
- *The notice test keeps only crates a dev-free `cargo tree` also lists* -- `cargo metadata` counts crates that only a dev-dependency's feature enables.
- *PEP 639 `license-files` puts the notice in the wheel only with `maturin>=1.9`* -- an older backend silently drops it. Full text: [dev-docs/license-notice.md](dev-docs/license-notice.md).

## The Exception Hierarchy (0.19.0+)

`src/errors.rs` is the only place in `src/` that constructs a `PyErr`; raise through the `errors::*` helpers. Full text: [dev-docs/exceptions.md](dev-docs/exceptions.md). User-facing contract: `docs/errors.md`.

- *A new class's builtin base is what the site raises today, found by grepping* -- the second base is a compatibility contract; `OptionError` must take no builtin base and is never raised.
- *The 93 `pytest.raises(ValueError|TypeError)` assertions are the pre-0.19 compatibility gate* -- editing one to stay green is a breaking change.
- *Every exported class needs a working trigger* -- a class no site raises is dead API that can never be removed.
- *`ConvertError` has no `From<String>`* -- every failure site names `Config` or `File`; do not add a fallback variant.
- *`FileError.errno` is set; `strerror` and `filename` stay `None`* -- `filename` makes `str()` discard the message.
- *`FilePermissionError` is chosen in `errors::file` by `errno` (`EACCES`, `EPERM`)* -- a subclass, so `except PermissionError` does not catch a missing folder; the POSIX tests skip under root.
- *`FrameOverlapError` is also a `FileError` for compatibility only* -- a repeated sheet name raised `FileError` before; drop that parent in 2.0.0.

## Python Lint, Type, and Security Gates

- *ruff, bandit, pyright and pytest run from the project-local `.venv`* -- `.venv\Scripts\<tool>.exe` on Windows, `.venv/bin/<tool>` elsewhere: `ruff check python tests benchmarks scripts`, `bandit -c pyproject.toml -r python`, `pyright`, `python -m pytest tests/ -q`.
- *Their scopes are deliberate; do not widen them* -- pyright strict on `python/xlsxturbo` only, `pythonPlatform = "Linux"`, bandit on `python/` only. Full text: [dev-docs/python-gates.md](dev-docs/python-gates.md).

## Coverage, and why the obvious command lies

- *`cargo llvm-cov` alone reports 26% and is wrong here* -- use `python scripts/coverage_report.py`, which merges the Rust and Python-driven profiles.
- *There is no coverage threshold, and adding one would be a regression* -- a Coverage job failure means the measurement broke. Full text: [dev-docs/testing.md](dev-docs/testing.md#coverage-and-why-the-obvious-command-lies).

## Dates and times are our own code

- *`src/parse/civil.rs` replaced the unmaintained `chrono` and keeps its exact rules* -- year range, text leniencies and arithmetic were proven identical over every date and 80 million texts before the swap; do not replace it with a date crate without the same comparison. Full text: [dev-docs/testing.md](dev-docs/testing.md#replacing-a-dependency-run-both-side-by-side-before-the-old-one-goes).

## Property tests

- *State a property as an equivalence, and mutate the code to prove the generator reaches the failing case* -- `".*"` generates short strings; `proptest-regressions/` is gitignored, promote a real failure to a named test. Full text: [dev-docs/testing.md](dev-docs/testing.md#property-tests).

## The stability promise (1.0.0+)

`docs/stability.md` is the public contract: which names are covered, what counts as a
breaking change, the deprecation terms, and the supported Python and platform matrices.
Read it before changing anything user-visible — from 1.0.0 a rename, a narrowed option
value, or a different exception for an existing failure is a 2.0.0 event, except
for the documented-contract bug-fix exception in that page. A patch may repair a
verified violation of the pre-existing contract with a reproducer, regression test
and explicit old/new behavior in its release notes; it may not redefine the contract
to justify a redesign. Approved for the CSV I/O classification repair in 1.4.1.

- *The page is checked, not maintained by hand* -- `tests/test_stability_policy.py` fails a new export or CI Python version until the page agrees. Full text: [dev-docs/stability-policy.md](dev-docs/stability-policy.md).

### A new CPython needs no code change — and the `abi3` promise has one hole

Full text: [dev-docs/python-versions.md](dev-docs/python-versions.md#a-new-cpython-needs-no-code-change--and-the-abi3-promise-has-one-hole).

- *Declaring a new CPython is the trove classifier plus a `docs/stability.md` row, no code* -- never add a `t` row to that table; a CI leg waits for pandas wheels, and for 3.15 it also needed polars 2.
- *The `abi3` wheel does not cover free-threaded builds* -- `3.14t` falls back to the sdist; do not promise free-threaded support without asking.
- *`df_to_xlsx` and `dfs_to_xlsx` detach for the save* -- `convert_csv` already runs inside `csv_to_xlsx`'s detach, do not wrap it again.
- *`TestSaveReleasesTheGil` is the only guard for the detach, per call site* -- widen its threshold if it flakes, never delete it; it skips under `XLSXTURBO_COVERAGE`.

### Output is deterministic except for one part, and the obvious measurement says otherwise

- *Only `docProps/core.xml` differs between two exports, and a quick hash says identical* -- its timestamp has one-second resolution; `TestGeneratedFileDeterminism` waits 1.1 s on purpose. Full text: [dev-docs/python-versions.md](dev-docs/python-versions.md#output-is-deterministic-except-for-one-part-and-the-obvious-measurement-says-otherwise).

## Benchmarks

- *Benchmarks are seeded, use native-fast dtypes, rotate their run order and name the machine and method* -- suites, flags and the memory column: [dev-docs/benchmarks.md](dev-docs/benchmarks.md).

## Documentation Sync

- *A feature change updates its `docs/` page (not the README), the generated capability matrix, `CHANGELOG.md` and the stub `xlsxturbo.pyi`* -- full text: [dev-docs/documentation-sync.md](dev-docs/documentation-sync.md).
- *Unreleased work goes under `## [Unreleased]`, never `## [X.Y.Z] - Unreleased`* -- `release-notes.sh` matches the version heading as a prefix, so the second form publishes a Release whose notes say Unreleased with exit 0; `TestChangelogHeadings` refuses it. This repo's rule wins over the cross-repo note. Full text: [dev-docs/changelog.md](dev-docs/changelog.md#unreleased-work-goes-under-unreleased).
- Before commit or push, follow the `BUILD.md` checklist.

- *docs/ is published wholesale* -- internal files there must be in `mkdocs.yml`'s `exclude_docs`; `--strict` does not catch a leak; never `mkdocs gh-deploy` by hand. Full text: [dev-docs/docs-site.md](dev-docs/docs-site.md).

### Two CLIs, one interface

- *The wheel's `xlsxturbo` command is `python/xlsxturbo/cli.py`, not the Rust binary* -- a new option goes into both; `tests/test_cli.py::TestParityWithRustBinary` fails when their options differ. Full text: [dev-docs/cli.md](dev-docs/cli.md).
