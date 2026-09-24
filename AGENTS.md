# xlsxturbo Agent Instructions

## Shared Memory Policy

- `AGENTS.md` is the canonical shared memory for Codex and other coding agents in this repository.
- Claude Code loads this file through the root `CLAUDE.md`.
- Durable project knowledge, workflows, commands, architecture notes, and recurring pitfalls belong here.
- Do not store durable project knowledge only in Claude auto memory.
- Keep entries concise, specific, and verifiable. Prefer updating existing sections over appending duplicate notes.

## Git Workflow

- Only commit and push when explicitly asked by the user.
- Do not include Claude-related or AI-generated footers in commit messages.
- Before commit or push, run `cargo update` to check for Rust dependency updates — **and
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
- This repo uses a project-local `.venv` (an exception to any central-venv convention). Test deps (`pytest pandas polars openpyxl`) must be installed there; if they are missing, `pytest` can resolve to a system Python with a stale extension and report bogus signature mismatches. Verify the interpreter in the pytest header is `.venv\Scripts\python.exe`; recover with `uv pip install -e ".[dev]" -r requirements-test.txt -r requirements-docs.txt` and rebuild via `maturin develop --release`.
- **Never run a bare `uv run` here; always `uv run --no-sync`** (or `--no-project` for a throwaway `--with` environment). A plain `uv run` first syncs `.venv` exactly to `uv.lock` with no extras, which uninstalls pytest, ruff, pyright, maturin, pandas, polars and mkdocs (54 packages by `uv sync --dry-run`) and replaces the `maturin develop` build. That is the usual way the test deps go missing: during the 1.5.1 release, BUILD.md's own rebuild step (then without `--no-sync`) removed pytest one line before the suite was due to run. `tests/test_ci_config.py::TestUvRunDoesNotResyncTheVenv` fails on a bare `uv run` in `BUILD.md`, `AGENTS.md` or `CONTRIBUTING.md`.
- **A new third-party import in a test cannot be validated locally — add it to `requirements-test.txt`.** The local `.venv` holds the whole `dev` extras, but the CI test jobs and the release smoke test install only `requirements-test.txt`, so the local environment is a strict superset of CI and a test importing anything outside that file passes locally and fails only in CI. Declaring it in `[project.optional-dependencies] dev` does **not** fix it; those jobs never install `dev`. `tests/test_ci_config.py` fails if an import is undeclared or if a workflow re-inlines the list.

  *A declared version range is only supported at the end CI installs* -- pandas 2 is covered only by the `requirements-test-pandas2.txt` leg, which layers the main file; a second two-major dependency needs its own leg. Full text: [dev-docs/testing.md](dev-docs/testing.md).

- **The release smoke test runs `pytest tests/` against an installed wheel from outside the checkout, with only `tests/` copied.** So a test module that reads *repository* files — `mkdocs.yml`, the capability matrix, the generator script, the workflows — has nothing to read there and must skip via `tests.helpers.repo_checkout_available()`, not fail. This is invisible in ordinary CI and surfaces only when a tag is pushed: `test_docs_site.py` and `test_capability_matrix.py` were both added after v0.18.0, so v0.19.0 was the first release to run them and all 16 failed. Inside a checkout a missing file stays a hard failure — the guard distinguishes "no repository here" from "the repository is broken". Reproduce the job locally before tagging: `cp -r tests "$TMP/smoke" && cd "$TMP/smoke" && python -m pytest tests/ -q`.
- Standard local checks: `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`, `cargo test --release`, `maturin develop --release`, then `pytest tests/`, then `mkdocs build --strict` when anything under `docs/` moved (needs `uv pip install -r requirements-docs.txt` into the venv once; the `dev` extras deliberately do not carry it). `BUILD.md`'s Pre-Push Checklist is the CI-verbatim list; this line is the short form.
- **`tests/test_build_is_current.py` fails when the built extension is older than any tracked Rust source, `Cargo.toml`, `Cargo.lock` or `pyproject.toml`.** A stale `.so` does not error, it measures the wrong code with full confidence: on 2026-09-02 a whole review ran against 1.3.0 code believing it was HEAD, and a later gate run blamed a timing test for what was the build. CI builds immediately before pytest, so the test only ever fires locally, where `maturin develop --release` is the answer. A `git checkout` that touches a `.rs` file trips it on purpose.
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.
- If multiple version-bump commits are awaiting release, tag each released version at its own commit; do not collapse distinct versions into one tag.
- Linux release wheels intentionally use `manylinux_2_28` with maturin's `--find-interpreter` and PyO3 `abi3-py310`. Do not switch back to automatic manylinux selection without verifying Python 3.10+ abi3 wheels. (`abi3-py39` until 1.1.0 dropped Python 3.9.)
- *The release workflow smoke-tests every published wheel on a runner of its own architecture* -- each leg asserts its wheel's platform tag; `tests/test_stability_policy.py` ties the build and smoke matrices to `docs/stability.md`. Full text: [dev-docs/testing.md](dev-docs/testing.md).
- To confirm a release on PyPI, query the version-specific endpoint `https://pypi.org/pypi/xlsxturbo/<version>/json` (authoritative within seconds). The aggregate `https://pypi.org/pypi/xlsxturbo/json` `info.version` field lags several minutes behind (CDN cache) and can still show the previous version; trust the publish job's green status over it.
- For multi-phase implementation work, run a deep diff review after each completed, verified phase before building the next phase on top of it.

## Adding a Feature - the Touchpoint Checklist

(Named without a count on purpose: it was "7-Touchpoint" for two releases after the eighth
item arrived, and two documents and a source comment carried the wrong number with it.)

1. `src/types.rs` - add the field to the `define_options!` macro list (generates ExtractedOptions/EffectiveOpts/as_effective/merge_with) AND the matching field on the hand-written `SheetConfig` struct. A missing SheetConfig field is a compile error in the generated merge_with. A cell_ref/location-keyed feature map (images, charts, comments, ...) must be `IndexMap`, not `HashMap` — iteration order feeds straight into the generated XML, so a `HashMap` makes output non-reproducible across runs.
2. `src/extract.rs` - add `extract_<feature>()`; register it in `extract_sheet_info` via the `extract_dict_field!`/`extract_list_field!` macro and add the option name to `SHEET_OPTION_NAMES` (guard test enforces this). Two extraction patterns coexist by design: simple structures (column_widths, formula_columns, merged_range tuples) are eagerly typed into real Rust types here, at extract time; features whose parsing needs a `py`/rust_xlsxwriter type (a `Format`, a `Color`, a chart/sparkline builder) instead extract only a raw `HashMap<String, Py<PyAny>>` "blob" here and defer real validation to the matching `apply/*` function, since that parsing can't happen without the GIL-bound types apply time has. Don't "fix" a blob extractor by eagerly typing it — that's the wrong layer for that feature.
3. `src/lib.rs` - add the field to `RawOptions` + `extract_options()`, and the kwarg + `#[pyo3(signature)]` entry + docstring to BOTH `df_to_xlsx` and `dfs_to_xlsx`.
4. `src/apply/<family>.rs` (+ re-export in the `src/apply.rs` facade) - `apply_<feature>()` with unknown-key validation and context-rich errors (`format!("<feature>['{}']: ...", cell_ref)`). For a blob-extracted feature, build one `types::OptionMap::new(py, &blob, context)` per config and use its typed accessors (`.string()`, `.bool()`, `.f64()`, `.u32()`, `.dict()`, `.reject_unknown()`/`.reject_unknown_for()`) instead of hand-rolling a new `<feature>_string_field`-style wrapper family — that duplication (~400 lines across charts/sparklines/validations/media/conditional_formats/format-dict parsing) is exactly what `OptionMap` replaced.
5. `src/convert.rs` - apply call in `apply_worksheet_features` (order matters: `cells` stays last so user cells can overwrite data). Decide constant_memory classification: a new option defaults to skipped+warned; add to `CONSTANT_MEMORY_SAFE_OPTIONS` only if applied during the data write. The guard test `every_complex_option_is_classified_for_constant_memory` forces this decision.
6. **`python/xlsxturbo/types.py`** - the option `TypedDict`/`Literal` for the new feature. Since 0.19.0 this runtime module, NOT the stub, is where shapes are declared; `xlsxturbo.pyi` imports them with the redundant-alias form (`X as X`) and keeps only the four function signatures, the exception classes and `__version__`. So a new option means: add the shape to `types.py`, add `X as X` to the stub's import block **and** to its `__all__`, then add the kwarg to both function signatures and to `SheetOptions` (which is itself in `types.py`). `tests/test_types_module.py` fails if the stub's re-export list and the runtime module disagree. `__init__.pyi` is a thin re-export of the *runtime* surface - never hand-edit it for new options. Since 1.1.0 raised the floor to Python 3.10, `types.py` may use `X | Y` everywhere, in module-level aliases as well as field annotations. Two guards enforcing the old `Union[...]` spelling, and the ruff `keep-runtime-typing` setting behind them, were deleted with 3.9 rather than left as folklore -- the constraint is now enforced by the language version. What remains is that `typing.get_type_hints()` must keep working on these shapes, which `tests/test_types_module.py` checks.
7. `tests/test_<feature area>.py` - a `TestXxx` class following the existing per-feature test files (behavior-coupled: read the produced xlsx back via openpyxl or XML).
8. **`python/xlsxturbo/options.py`** - the matching `ExportOptions` field, plus a sample value in `tests/test_options.py`'s `SAMPLE_VALUES`. Both are enforced: `TestCoverage` derives the option list from `inspect.signature(df_to_xlsx)` and fails on a field the signature lacks or a signature parameter no field mirrors, and `test_sample_values_cover_every_field` fails if the sample table falls behind. **The field's annotation must be byte-identical to the same parameter's annotation in `xlsxturbo.pyi`** - `tests/test_options_types_match_the_stub.py` compares them as source text in both directions. That guard was added after five fields had drifted unnoticed: four to `Any` inside a container, and `row_heights` to something *narrower* than the function accepts, which is the worse direction because a checker then rejects working code. `inspect.signature` cannot catch this - a compiled pyo3 function carries no annotations, which is why the stub is the reference. Nothing else is needed - `as_kwargs`/`as_sheet_options` iterate the dataclass fields, so a new field flows through both lowerings automatically, and `as_sheet_options`'s workbook-only exclusion set is verified against what a per-sheet dict actually rejects rather than hand-maintained.

   This eighth touchpoint is a real, permanent tax, accepted deliberately for discoverability (roadmap D7). It is one line of code plus one line of test data; it is bounded because `ExportOptions` is flat and mirrors the kwargs one-to-one, and it is enforced rather than remembered, which is the only reason it is affordable.

Raising from Rust: **never `pyo3::exceptions::Py*Error::new_err` in `src/`.** Use the
`crate::errors::*` helpers — see the section below. `src/apply/` and `src/parse/` are
unaffected: they return `Result<_, String>` and the boundary classifies for them.

Then regenerate the capability matrix: `python scripts/gen_capability_matrix.py --write`. `docs/capability-matrix.md` is GENERATED from the Rust sources and must never be hand-edited; `tests/test_capability_matrix.py` fails if the committed page is stale. The generator parses `SHEET_OPTION_NAMES`, `define_options!`, `CONSTANT_MEMORY_SAFE_OPTIONS`, `warn_constant_memory_skips` and the three `#[pyo3(signature)]` blocks, so touching any of those changes the page. Each parser raises rather than returning an empty list, because a structural audit that matches nothing reads exactly like a clean result — and each parsed parameter must be a Python identifier, which is what caught a regex that spanned from the file's first pyo3 attribute through to the requested function and produced "parameters" like `) -> PyResult<(u32`.

## Upstream defects belong upstream — file them

- *When rust_xlsxwriter is what is wrong, file it at jmcnamara/rust_xlsxwriter alongside the local workaround* -- #185 and #186 were each fixed within a day of filing. Full text: [dev-docs/upstream-rust-xlsxwriter.md](dev-docs/upstream-rust-xlsxwriter.md#upstream-defects-belong-upstream--file-them).
- *A report lands with a reproducer that needs no Excel, controls in the same program, and only what was measured* -- review the draft against what you actually ran.
- *Pin an unreachable upstream defect in a test that fails when it is fixed* -- delete that test together with the workaround.

## Name validation is ours now — upstream has drawn its line

Full text: [dev-docs/upstream-rust-xlsxwriter.md](dev-docs/upstream-rust-xlsxwriter.md#name-validation-is-ours-now--upstream-has-drawn-its-line). Read it before touching `reject_reference_shaped_name`, `sanitize_table_name`, `claimed_table_name` or a rust_xlsxwriter bump.

- *The crate's name checks are a foot-gun guard, not a specification* -- the two screens here are the layer xlsxturbo owns, not temporary workarounds.
- *Rename what nobody points at, refuse what somebody might* -- `reject_reference_shaped_name` refuses defined names; `sanitize_table_name` rewrites table names.
- *A table/defined-name collision is a pre-check, not a save failure* -- `claimed_table_name` (`src/convert.rs`) owns the gate both pre-checks use; any new save-time rule upstream is a pre-check candidate.
- *`sanitize_table_name` normalises to NFC, never NFKC, before the screen* -- Thai/Hindi marks are still rewritten on purpose; a denylist flip needs the mirror audit, not reading.
- *The allowlist must stay at least as wide as the crate's denylist* -- on every bump, account for each rule in `utility::check_name`; re-pin the empty-name message tests, do not reinstate the screen.

## The bundled license notice

- *`THIRD-PARTY-LICENSES.md` is generated; never edit it* -- `python scripts/gen_third_party_licenses.py --write`; `tests/test_third_party_licenses.py` checks it both ways.
- *`cargo install cargo-about` installs nothing and exits 0* -- use `--features cli`; the generator passes `--output-file` on purpose.
- *PEP 639 `license-files` puts the notice in the wheel only with `maturin>=1.9`* -- an older backend silently drops it. Full text: [dev-docs/license-notice.md](dev-docs/license-notice.md).

## The Exception Hierarchy (0.19.0+)

`src/errors.rs` is the only place in `src/` that constructs a `PyErr`; raise through the `errors::*` helpers. Full text: [dev-docs/exceptions.md](dev-docs/exceptions.md). User-facing contract: `docs/errors.md`.

- *A new class's builtin base is what the site raises today, found by grepping* -- the second base is a compatibility contract; `OptionError` must take no builtin base and is never raised.
- *The 93 `pytest.raises(ValueError|TypeError)` assertions are the pre-0.19 compatibility gate* -- editing one to stay green is a breaking change.
- *Every exported class needs a working trigger* -- a class no site raises is dead API that can never be removed.
- *`ConvertError` has no `From<String>`* -- every failure site names `Config` or `File`; do not add a fallback variant.
- *`FileError.errno` is set; `strerror` and `filename` stay `None`* -- `filename` makes `str()` discard the message.

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

## Coverage, and why the obvious command lies

- *`cargo llvm-cov` alone reports 26% and is wrong here* -- use `python scripts/coverage_report.py`, which merges the Rust and Python-driven profiles.
- *There is no coverage threshold, and adding one would be a regression* -- a Coverage job failure means the measurement broke. Full text: [dev-docs/testing.md](dev-docs/testing.md#coverage-and-why-the-obvious-command-lies).

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

It is **checked, not maintained by hand**. `tests/test_stability_policy.py` compares the
page against the four places that actually decide its contents — the trove classifiers and
`requires-python` in `pyproject.toml`, the interpreter matrix in `ci.yml`, the wheel targets
in `release.yml`, and `xlsxturbo.__all__` — in both directions. Consequences:

- **A new exported name fails the suite until the page names it.** That is deliberate: the
  page is the list of things that cannot be removed before 2.0.0, so adding to it should be a
  decision rather than a side effect. New *options* need nothing here; the touchpoint
  checklist already covers those.
- **Adding a Python version to the CI matrix fails until the classifier and the page agree.**
  A version tested and nowhere else declared reads as a widened promise when it is only a
  widened test.
- **This worked.** Dropping Python 3.9 in 1.1.0 failed both suites exactly as designed — the
  page's version table against the classifiers and the CI matrix, and the Dependabot hold in
  `tests/test_ci_config.py` that existed only because pytest 9 needs 3.10. Neither failure was
  a surprise to be worked around; each named the work that had to accompany the drop.

### A new CPython needs no code change — and the `abi3` promise has one hole

Full text: [dev-docs/python-versions.md](dev-docs/python-versions.md#a-new-cpython-needs-no-code-change--and-the-abi3-promise-has-one-hole).

- *Declaring a new CPython is the trove classifier plus a `docs/stability.md` row, no code* -- never add a `t` row to that table; a CI leg waits for pandas wheels.
- *The `abi3` wheel does not cover free-threaded builds* -- `3.14t` falls back to the sdist; do not promise free-threaded support without asking.
- *`df_to_xlsx` and `dfs_to_xlsx` detach for the save* -- `convert_csv` already runs inside `csv_to_xlsx`'s detach, do not wrap it again.
- *`TestSaveReleasesTheGil` is the only guard for the detach, per call site* -- widen its threshold if it flakes, never delete it; it skips under `XLSXTURBO_COVERAGE`.

### Output is deterministic except for one part, and the obvious measurement says otherwise

- *Only `docProps/core.xml` differs between two exports, and a quick hash says identical* -- its timestamp has one-second resolution; `TestGeneratedFileDeterminism` waits 1.1 s on purpose. Full text: [dev-docs/python-versions.md](dev-docs/python-versions.md#output-is-deterministic-except-for-one-part-and-the-obvious-measurement-says-otherwise).

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

   **Unreleased work goes under `## [Unreleased]`, never `## [X.Y.Z] - Unreleased`, and the difference decides whether a forgotten rename fails loudly or ships wrong.** `release-notes.sh` matches `## [<version>]` as a prefix and ignores whatever follows on the line, so a heading already carrying the version number matches its tag: measured against this repo's own file, `release-notes.sh v1.1.2` on `## [1.1.2] - Unreleased` exits **0** and prints the section, and the release job then publishes a GitHub Release whose notes came from a heading that says Unreleased. The same call on `## [Unreleased]` exits **1** with `no CHANGELOG section found for version '1.1.2'`, which fails the job before anything is published. Step 2 of `BUILD.md`'s release process is where the heading gets its version and date. `tests/test_ci_config.py::TestChangelogHeadings` fails on a bracketed-version heading whose date field is missing or reads "Unreleased", so the unsafe form cannot come back by habit.

   This contradicts the cross-repo note in the personal `agent-memory/AGENTS.md`, which states the `## [version] - Unreleased` form as a general rule. That form is only safe where the release tooling does not slice notes by matching the version heading; here it is not, and the repo's own file wins for repo facts. `screenpick` and `tpdf` use the other form and were not measured — the convention is a per-repo fact, decided by that repo's release script.
4. **Type stubs** (`python/xlsxturbo/xlsxturbo.pyi`) — add new parameters with types and docstrings. This compiled-extension stub is the type source of truth; `python/xlsxturbo/__init__.pyi` is a thin re-export and must not be hand-edited for new options.

Before commit or push, follow the `BUILD.md` checklist.

### docs/ is published wholesale — mind what is sitting in it

MkDocs publishes every file under `docs/`, and knows nothing about git. Two tracked-but-internal files (`roadmap-1.0.md`) and two untracked ones (`strategic-recommendations-plan.md`, `reviews/`) are therefore listed in `mkdocs.yml`'s `exclude_docs`, and `tests/test_docs_site.py` asserts that list stays in step with `.gitignore`.

Verified by removing one `exclude_docs` entry: the private planning memo appeared in `site/` and `mkdocs build --strict` still **exited 0**. Strict mode does not cover this. The reliable protection is that deployment happens only from `.github/workflows/docs.yml`, which builds from a clean checkout and so cannot see an untracked file at all — never run `mkdocs gh-deploy` by hand.

### The CLI is not in the wheel

`Cargo.toml` has a `[[bin]] xlsxturbo` target with `default = ["cli"]`, so `cargo build --release` produces a working CLI. maturin builds only the extension module, so the published wheel contains **no console script and no binary** — confirmed by inspecting the PyPI artifact for 0.18.0, not by reading the config. The README and the CLI docs claimed otherwise until 0.19.0. If shipping it is ever wanted, that is a packaging change, not a documentation one.
