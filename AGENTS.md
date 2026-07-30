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

  **A declared version range is only supported at the end CI installs.** pip resolves to the
  newest allowed, so `pandas>=2.3.3,<4` means pandas 3 on every leg and pandas 2 never. Since
  1.1.0 a fourth `python-test` leg installs `requirements-test-pandas2.txt`, which **layers**
  the main file (`-r requirements-test.txt`) and overrides only the pandas ceiling — never
  restates the list, which would be a fifth copy of it. `tests/test_ci_config.py` fails if
  that file stops layering, gains a package, stops being referenced by a workflow, or outlives
  its reason (the pandas range narrowing back to one major). Scoped to pandas because it is
  the only dependency whose range currently spans two majors; a second one needs its own leg
  and nothing will notice on its own.

  How the gap was found is the transferable part: the local `.venv` had pandas 3 while
  `requirements-test.txt` said `<3`, so **local runs and CI had not been testing the same
  library for weeks** and everything was green on both. The usual assumption here — that the
  local environment is a strict superset of CI — was simply false, in the direction nothing
  checks.

  That guard exists because the same bug landed three times in two days: `tests/test_docs_site.py` imported `yaml` (declared in `dev`, broke three CI jobs); fixing those three left a **fourth** copy of the list in `release.yml`'s smoke-test job, which failed the v0.19.0 release after every wheel had already built; and the guard's first run found `numpy` imported by `tests/test_core.py` and never declared anywhere — working only because pandas pulls it in. A comment saying "remember the other copies" was the fix after the first, and it did not survive a day.

- **The release smoke test runs `pytest tests/` against an installed wheel from outside the checkout, with only `tests/` copied.** So a test module that reads *repository* files — `mkdocs.yml`, the capability matrix, the generator script, the workflows — has nothing to read there and must skip via `tests.helpers.repo_checkout_available()`, not fail. This is invisible in ordinary CI and surfaces only when a tag is pushed: `test_docs_site.py` and `test_capability_matrix.py` were both added after v0.18.0, so v0.19.0 was the first release to run them and all 16 failed. Inside a checkout a missing file stays a hard failure — the guard distinguishes "no repository here" from "the repository is broken". Reproduce the job locally before tagging: `cp -r tests "$TMP/smoke" && cd "$TMP/smoke" && python -m pytest tests/ -q`.
- Standard local checks: `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`, `cargo test`, `maturin develop --release`, then `pytest tests/`.
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.
- If multiple version-bump commits are awaiting release, tag each released version at its own commit; do not collapse distinct versions into one tag.
- Linux release wheels intentionally use `manylinux_2_28` with maturin's `--find-interpreter` and PyO3 `abi3-py310`. Do not switch back to automatic manylinux selection without verifying Python 3.10+ abi3 wheels. (`abi3-py39` until 1.1.0 dropped Python 3.9.)
- **The release workflow must smoke-test every published wheel on a runner of its own architecture, not a representative subset.** All five legs install the wheel and run the full suite before the publish job. Two of them existed only from 1.1.0 onward: Linux `aarch64` and macOS `x86_64` are cross-compiled and had no hosted runner when the pipeline was written, so they shipped untested for every release before that. The runners are `ubuntu-24.04-arm` and `macos-15-intel` — note `macos-13`, the old Intel image, has been retired, so check the current label before assuming. Each leg asserts the platform tag of the wheel it downloaded; without that a mistyped `wheel-artifact` installs some other wheel twice and two green legs read as coverage. `tests/test_stability_policy.py` compares the build matrix and the smoke-test matrix in both directions and ties them to the table in `docs/stability.md`, so a new build target cannot be published untested and a leg naming a non-existent artifact fails locally rather than after every wheel has built.
- To confirm a release on PyPI, query the version-specific endpoint `https://pypi.org/pypi/xlsxturbo/<version>/json` (authoritative within seconds). The aggregate `https://pypi.org/pypi/xlsxturbo/json` `info.version` field lags several minutes behind (CDN cache) and can still show the previous version; trust the publish job's green status over it.
- For multi-phase implementation work, run a deep diff review after each completed, verified phase before building the next phase on top of it.

## Adding a Feature - the 7-Touchpoint Checklist

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

## The one upstream defect worked around in this codebase

`apply::reject_databar_with_sparklines` refuses a `data_bar` conditional format beside a
sparkline on one worksheet, because rust_xlsxwriter 0.97.0 emits unbalanced `<ext>` elements
for that pair and Excel reports the workbook as damaged. Found by the full-bundle test in
`tests/test_options.py`, which writes all 24 options at once — a combination nothing else
exercised.

Two things worth knowing before touching it:

- **The guard makes the bug unreachable from Python, so no Python test can notice it being
  fixed.** `tests/upstream_defect.rs` uses rust_xlsxwriter directly and asserts the defect is
  *still present*; when upstream fixes it that test fails, which is the signal to delete the
  guard, the Rust test, and the note in `docs/errors.md`. It has a control alongside it so a
  worse regression cannot be misread as the known defect.
- **Over-reach is the expensive failure here, not under-reach.** A guard refusing a
  combination that is actually fine costs a feature silently, since users read the error as
  their own mistake. `TestDataBarSparklineGuard::test_the_guard_is_narrow` pins the adjacent
  cases that must keep working, and it is what catches a widened condition.

## The Exception Hierarchy (0.19.0+)

`src/errors.rs` owns the public exception classes and is the only place in `src/` that may
construct a `PyErr` from scratch. Raise with `errors::configuration`,
`errors::configuration_type`, `errors::input_data`, `errors::file` or
`errors::workbook_validation`. The classes are built by calling the `type` metaclass rather
than with `create_exception!`, because that macro takes a single base and every class here
needs two or three.

Facts that are expensive to rediscover:

- **`OptionError` is never raised, and that is not an oversight.** It exists so
  `except OptionError` catches both configuration classes and nothing else. The guard that
  every exported class needs a working trigger still applies to it, in a different form:
  `ABSTRACT` in `tests/test_errors.py` maps it to exactly the triggered classes it must
  catch, checked in both directions. Declaring a class abstract is otherwise a way to walk
  straight past the rule that killed `UnsupportedFeatureError`.
- **`OptionError` must not take a builtin base.** A builtin there lands on *both* children,
  so a `ConfigurationTypeError` would silently also be a `ValueError` and the value/type
  split would stop meaning anything to `except`. `FORBIDDEN_BASES` pins it.
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
- **`ConvertError` in `src/convert.rs` is the one seam, and it has no `From<String>`.** Two
  variants, `Config` and `File`, because `save_workbook` runs inside the pipeline and its
  failure would otherwise be indistinguishable from a bad option at the boundary.

  Until 0.21.0 a blanket `From<String>` mapped untagged failures to `Config`, so `?` compiled
  everywhere and **a new filesystem call silently blamed the caller's options.** That
  conversion is gone: every site now names its variant, and a new failure site *does not
  compile* until it chooses. Prefer that over the `Internal` fallback variant the review
  proposed — a fallback that still exists is still a default, and the default was the bug.
  `TestConvertErrorHasNoDefaultCategory` watches for it coming back, with a control asserting
  both variants are still constructed.

  Frame detection needed the same treatment and got a boundary call
  (`require_supported_dataframe`) instead, since tagging it would have pushed `ConvertError`
  down into the write layer.
- **`FileError.errno` is populated; `strerror` and `filename` must stay `None`.** Setting
  `filename` makes `OSError.__str__` switch to `[Errno n] strerror: 'filename'` and **discard
  the message**, which is where this library puts the path and the context. `errno` alone
  leaves `str()` untouched. On Windows `raw_os_error()` is a Win32 code, not an errno —
  `ERROR_PATH_NOT_FOUND` is 3, which as POSIX means "no such process" — so `errors::posix_errno`
  passes the number through on Unix and classifies via `io::ErrorKind` elsewhere.

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

## Coverage, and why the obvious command lies

`python scripts/coverage_report.py` (add `--html` for a browsable report). It needs
`rustup component add llvm-tools-preview`; everything else is in the dev extras.

**Do not use `cargo llvm-cov` on its own to judge this codebase.** It reports **26%** and
shows every `src/apply/*.rs` file at 0%, which reads as an untested library and is false —
those paths are covered thoroughly from the Python suite, on the other side of the FFI
boundary. The script instruments both the Rust test binaries *and* the extension module,
runs both suites, and merges the profiles: 92.96% of lines in the Rust core, 100% of the
Python layer. `cargo-llvm-cov` is deliberately not a dependency, because its `report`
subcommand takes no extra `--object` and the extension module is exactly that; the script
drives `llvm-profdata`/`llvm-cov` directly instead.

**There is no threshold, in CI or out, and adding one would be a regression.** A coverage
target gets met by tests that execute lines without asserting anything. The CI job is
informational: it publishes the table to the job summary and uploads HTML. It can still fail,
and a failure means the measurement broke, never that coverage fell.

Two filtering caveats the numbers depend on: `tests/`, `src/parse/proptests.rs` and
`src/parse/boundaries.rs` are excluded as test code, but the `#[cfg(test)] mod tests` *inside*
`src/parse/mod.rs` cannot be — `llvm-cov` has no sub-file filter — so that row means "the
tests in this file all ran", not anything about the parsers.

## Property tests

`src/parse/proptests.rs`. Three rules that are the difference between a property and a
decoration, each learned by writing one that failed the test:

- **State it as an equivalence, not an implication.** "A prefix pattern matches a string
  starting with the prefix" is satisfied by an implementation matching *everything*. Each
  pattern property asserts equality with the `str` method it claims to implement.
- **Check the generator can reach the failing case, by mutating the code.** A property over
  all printable ASCII stayed green when the guard it defends was deleted: the discriminating
  inputs were one in seventy thousand of that space. Narrow the alphabet until a near-miss is
  common, and where the case is nameable, write a second property whose generator *is* the
  case.
- **`".*"` generates short strings.** A property asserting a 255-character cap never entered
  the truncation branch. Anything about a length boundary needs a generator that straddles it.

`proptest-regressions/` is gitignored, against the usual advice: mutation-testing the suite
makes proptest save a seed for every property that correctly went red, describing code that no
longer exists. Promote a genuine failing case to a named test instead — it states the input
where a reader can see it.

## The stability promise (1.0.0+)

`docs/stability.md` is the public contract: which names are covered, what counts as a
breaking change, the deprecation terms, and the supported Python and platform matrices.
Read it before changing anything user-visible — from 1.0.0 a rename, a narrowed option
value, or a different exception for an existing failure is a 2.0.0 event, not a minor.

It is **checked, not maintained by hand**. `tests/test_stability_policy.py` compares the
page against the four places that actually decide its contents — the trove classifiers and
`requires-python` in `pyproject.toml`, the interpreter matrix in `ci.yml`, the wheel targets
in `release.yml`, and `xlsxturbo.__all__` — in both directions. Consequences:

- **A new exported name fails the suite until the page names it.** That is deliberate: the
  page is the list of things that cannot be removed before 2.0.0, so adding to it should be a
  decision rather than a side effect. New *options* need nothing here; the 7-touchpoint
  checklist already covers those.
- **Adding a Python version to the CI matrix fails until the classifier and the page agree.**
  A version tested and nowhere else declared reads as a widened promise when it is only a
  widened test.
- **This worked.** Dropping Python 3.9 in 1.1.0 failed both suites exactly as designed — the
  page's version table against the classifiers and the CI matrix, and the Dependabot hold in
  `tests/test_ci_config.py` that existed only because pytest 9 needs 3.10. Neither failure was
  a surprise to be worked around; each named the work that had to accompany the drop.

### Output is deterministic except for one part, and the obvious measurement says otherwise

Two exports of the same frame differ, because `docProps/core.xml` records the creation time.
Every other member of the archive is byte-identical across runs.

The trap is that writing both files and hashing them reports **identical** — that timestamp
has one-second resolution, so a loop with no delay measures nothing. It was one step from
being published as "output is byte-reproducible".
`TestGeneratedFileDeterminism` waits 1.1 s deliberately, and mutating that wait to zero turns
both of its tests red, which is what proves neither passes by accident. If a reproducible-build
option is ever added, it belongs here as an opt-in — and it is additive, so it needs no major.

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
