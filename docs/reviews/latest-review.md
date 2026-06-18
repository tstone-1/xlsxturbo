# Code Review

**Date**: 2026-06-09
**Scope**: codebase (at 610b20e, v0.15.3 — fast-forwarded from a 2-commit-stale clone before review)
**Depth**: deep
**Mode**: mine (elite)
**Reviewer**: Senior Dev (automated, 3 parallel agents + reconciliation)
**Tech Stack**: Rust library (rust_xlsxwriter 0.95, rayon, chrono) with PyO3 0.28 / abi3-py39 Python extension; optional clap CLI behind `cli` feature; Python type stubs; pytest suite (233 tests) + Rust unit tests (55). `target/`, venvs excluded.

**Verification performed during review**: `cargo test` 55 passed; full pytest suite 233 passed; `cargo audit` 0 vulnerabilities (117 crates); all Blocker/Warning behaviors below reproduced empirically against the locally built module with synthetic data (pandas 3.0.3, numpy 2.4.4, polars 1.40.1).

## Verdict: CHANGES REQUIRED

The architecture is in actively *improving* condition — the 0.15.2/0.15.3 anti-drift refactors were the right work, done well. But while everyone was polishing macros, the numeric/date boundary conversions in `convert.rs` were quietly mangling data: a non-nanosecond datetime64 column — which pandas 3 produces *by default* for out-of-range dates — gets silently wrapped around to the wrong century. Year 3000 in, year 1830 out, no error, no warning. That's silent data corruption reachable from ordinary input, in a package whose classifier says Production/Stable. Fix it before the next release.

## Maintainability Scorecard

| Dimension            | Score (0-5) | Notes |
|----------------------|:-----------:|-------|
| Cohesion             | 4 | `parse/` and `apply/` submodules sharply single-purpose; docked for `convert.rs` (1075 lines, two pipelines + writer primitives) and `media.rs` (4 feature families). |
| Coupling             | 4 | Near-perfect downward dependency flow; docked for the `apply/cells.rs → convert` back-edge and the 15-same-typed-positional-args `extract_options` seam. |
| Abstraction quality  | 4 | `define_options!`, `extract_opt`, `WriteConfig` all demonstrably earn their keep, zero speculative abstraction; docked for raw `Py<PyAny>` config blobs deferring validation to write time. |
| Complexity           | 4 | Long functions are linear ladders with rationale comments and checked arithmetic; docked for the 26-param × 4-declaration API surface. |
| Test robustness      | 4 | 233 behavioral tests with XML-level assertions; docked for existence-only assertions (images, color scales, validations), untested CLI, and missing 0.15.2 regression tests. |
| Operability          | 4 | Context-rich errors that enumerate valid alternatives; derived, guard-tested RuntimeWarning; docked for silently-ignored unknown per-sheet option keys. |

## Blockers

- **[B1] Non-ns numpy datetime64 silently corrupts out-of-range dates (wraparound)** (`src/convert.rs:409-415`)
  ```rust
  if type_name == "datetime64" {
      let ns_since_epoch: i64 = value
          .call_method1("astype", ("datetime64[ns]",))
          .and_then(|v| v.call_method1("astype", ("int64",)))
  ```
  pandas ≥2.x supports non-ns datetime units and pandas 3.0 produces them automatically: `pd.to_datetime(["3000-01-01"])` yields dtype `datetime64[us]`. The `astype("datetime64[ns]")` cast overflows i64 and numpy wraps silently — **reproduced**: the written cell read back as `datetime(1830, 11, 23, 0, 50, 52)`. Any date outside ~1677–2262 in a `[s]`/`[ms]`/`[us]` column (from `read_parquet`, `read_sql`, or plain `pd.to_datetime` on pandas 3) is silently destroyed. No error, no warning, wrong data in the output file.
  **Suggested fix:** convert via `astype("datetime64[us]")` (covers ±290k years) and build the timestamp from microseconds; or round-trip-check the ns cast and raise `ValueError` on mismatch.
  **Confidence:** WILL — reproduced.

## Warnings

- **[W1] Python ints beyond i64 silently lose precision, breaking the 2^53 string-fallback contract** (`src/convert.rs:509-517`)
  `write_int` falls back to string for |v| > 2^53 — but only for ints that fit i64. A Python int just above i64::MAX fails `extract::<i64>()`, then succeeds `extract::<f64>()` and is written as a rounded float. **Reproduced**: `2**63 + 1025` came back as `9223372036854778000`, while `2**53 + 1` correctly became the string `"9007199254740993"`. Inconsistent and silent.
  **Suggested fix:** before the f64 fallback, detect `PyInt` explicitly and route oversized ints to the string fallback.

- **[W2] `val.abs()` on i64::MIN: debug-build panic; release builds bypass the string fallback** (`src/convert.rs:93`, `src/convert.rs:131`)
  `i64::MIN.abs()` overflows: debug builds (`maturin develop`) hit an uncatchable pyo3 panic via `np.array([-2**63], dtype="int64")` or CSV cell `-9223372036854775808`; release builds wrap to a negative value, skip the `> 2^53` guard, and write the float `-9223372036854776000` instead of the intended string. **Reproduced** (release misclassification).
  **Suggested fix:** replace `val.abs() > MAX_SAFE_INT` with `val.unsigned_abs() > MAX_SAFE_INT as u64` at both sites.

- **[W3] Pre-1900 dates: DataFrame path writes negative Excel serials; CSV path guards against exactly this** (`src/convert.rs:425-427`, `:482-485`, `:502-506` vs `src/parse/values.rs:38-41`)
  The CSV path string-falls-back when `excel_dt <= 0.0` (tested in `tests/test_core.py:836-853`); the three DataFrame datetime/date branches don't. **Reproduced**: `datetime(1850,1,1,12,0)` → serial −18262.5, which Excel renders as `#####`. Same input, two behaviors.
  **Suggested fix:** apply the same `<= 0.0` string-fallback in the DataFrame branches.

- **[W4] Per-sheet options dict silently ignores unknown keys, contradicting the project's own strict-validation contract** (`src/extract.rs:69-180`)
  Found independently by two agents, **verified empirically**: `dfs_to_xlsx([(df, "S1", {"tabel_style": "Medium2"})], path)` succeeds silently. Everywhere else typos raise errors with a valid-keys list (`src/parse/formats.rs:226-241`, `src/apply/media.rs:52-59`, `:240-248`, `:479-528`, `src/apply/validations.rs:22-38`) — CHANGELOG 0.12.5 and 0.15.2 advertise this as a product guarantee. This is also the only option list with no drift guard: a new `define_options!` field missing its `extract_sheet_info` line compiles, passes the guard test, and silently no-ops only in the per-sheet path.
  **Suggested fix:** after the `PyDict` cast, reject unknown keys against a list derived from `EffectiveOpts::COMPLEX_OPTION_NAMES` + scalar names, so future drift fails a test.

- **[W5] Default integer column names fail with a confusing error** (`src/types.rs:220-231`)
  `df_to_xlsx(pd.DataFrame(np.eye(2)))` — pandas' default RangeIndex columns — raises `ValueError: TypeError: 'int' object is not an instance of 'str'`, which never mentions column names. **Reproduced.** `df.to_excel` handles this shape. Clean failure, terrible message, extremely common input.
  **Suggested fix:** stringify non-string column labels, or raise an error that names the problem.

- **[W6] `apply/cells.rs` depends back up on `convert`, the one cycle left in the layering** (`src/apply/cells.rs:3`)
  `use crate::convert::{write_py_value_with_format, DATETIME_NUM_FORMAT, DATE_NUM_FORMAT};` while `convert.rs:3-8` imports the `apply` facade — a `convert ↔ apply` cycle. CHANGELOG 0.15.3 celebrates "removed the last inward dependency arrow"; this arrow remains because the cell-writer primitives live inside the pipeline orchestrator instead of a leaf module. The next `apply/` feature needing a value writer will import `convert` too.
  **Suggested fix:** move `write_py_value_with_format`, `write_str/num/bool/int/float`, `write_cell`, and the NUM_FORMAT consts into a leaf `write` module that both `convert` and `apply/cells` import downward.

- **[W7] Single-sheet and multi-sheet pipelines are parallel implementations with verbatim-copied finalization logic** (`src/lib.rs:585-604` vs `src/convert.rs:996-1012`; `src/lib.rs:513, 556-566, 607-609` vs `src/convert.rs:973-981, 1014-1016`)
  The `defined_names` validation block is a character-for-character copy (including the panic-guard comment); workbook creation, `set_name` wrapping, and save handling are duplicated; the 26-entry kwargs surface is declared four times in `lib.rs`. A fix to the defined-names guard must currently be made twice — classic drift bait.
  **Suggested fix:** implement `df_to_xlsx` as the one-sheet case of the `dfs_to_xlsx` machinery, or extract a shared `write_workbook(...)`.

- **[W8] The 0.15.2 panic fixes shipped without regression tests** (`src/parse/mod.rs` — `test_sanitize_table_name_truncation` uses ASCII only; no test for empty `defined_names`)
  The multibyte `table_name` truncation fix is tested with `"a".repeat(300)` — the exact input class that caused the panic (`"é".repeat(300)`) is untested. The empty-defined-name → `ValueError` fix has no test at all. Both can regress silently.
  **Suggested fix:** see Test Assessment, proposed tests #3 and #4.

## Nitpicks

(Top 5 of ~14 raised by sub-agents; the rest were dropped as below the line.)

- **[N1] `pathlib.Path` rejected for path parameters** (`src/lib.rs:173-178`, `:310-313`) — `TypeError: 'PosixPath' object is not an instance of 'str'`. It's 2026; accept `os.PathLike`.
- **[N2] Dead `__main__` runner in `tests/test_core.py:861-918`** — references ~30 test classes that moved to other files in the suite split; `python tests/test_core.py` would `NameError` on line one of the block. Delete it; pytest is the documented runner.
- **[N3] `apply/media.rs` is a junk drawer in the making** (582 lines: images + checkboxes + textboxes + charts; charts alone ~280 lines at `media.rs:304-582`). Sparklines are on the roadmap and will land… where, exactly? Split `charts.rs` out now while it's cheap.
- **[N4] Docs still say constant_memory "silently disables" features** (`python/xlsxturbo/__init__.pyi:315-319, 407-411`, `README.md:1032`) — stale since the 0.15.3 RuntimeWarning; and the warning text itself (`src/convert.rs:757`) says "they will be silently skipped" *inside a warning*, which is a sentence at war with itself. README also undersells `CONSTANT_MEMORY_SAFE_OPTIONS` (omits `header_format`, `column_formats`).
- **[N5] `extract_options` takes 15 positional parameters of the identical type** (`src/lib.rs:68-84`) — every one is `Option<&Bound<PyAny>>`, so transposing two arguments at either call site compiles cleanly and surfaces only as wrongly-labeled runtime errors. Currently maintained by eyeball discipline at two call sites.

## Architecture Assessment

- **Boundaries & responsibilities**: Five clear layers with crisp jobs: `types` (data shapes + lowest-level helpers), `parse/` (string/dict → rust_xlsxwriter values, pure, unit-tested), `extract.rs` (Python → typed Rust, structure validation), `apply/` (typed Rust → worksheet calls, field-semantics validation), `convert.rs` (orchestration), `lib.rs` (PyO3 surface). The CLI consumes only the re-exported `pub` API behind the `cli` feature — it cannot reach Python-coupled internals.
- **Dependency direction**: Almost perfectly downward — `parse → types`; all `apply/*` import only `parse` + `types` **except** `apply/cells.rs:3 → convert` (W6); `extract → parse + types`; `convert → apply + parse + types`; `lib → convert + extract + types`. `extract` and `apply` correctly never import each other.
- **Change-risk hotspots**: (1) Adding a feature option touches ~9 places across 4 files; the macro + `merge_with` + guard test make most omissions compile errors — the only *silent* failure point is `extract_sheet_info` (W4). (2) The df/dfs finalization duplication (W7). (3) `convert.rs` at 1075 lines hosting two unrelated pipelines.
- **Abstraction audit**: Notably disciplined — no traits with one impl, no dyn dispatch, no premature generality. `define_options!` earns its keep decisively (it replaced a proven invisible-bug source) and stops at the right boundary: `SheetConfig` stays hand-written so a missing field is a compile error. The one thin spot: `validations`/`charts`/`conditional_formats` cross the extract→apply boundary as raw `HashMap<String, Py<PyAny>>` blobs (`types.rs:98, 125, 128`), so field validation happens at write time. Deliberate two-phase design; acceptable, but it's where errors surface latest.
- **3AM test**: Mostly passes — extending an existing feature is a 3-file grep with compile-error guardrails. What fails: adding a whole new feature option (the 9-touch-point checklist exists only in commit history) and discovering per-sheet support requires a separate, unenforced `extract_sheet_info` edit.

Overall: moving in a good direction, and measurably — the CHANGELOG shows a deliberate one-debt-paydown-per-release cadence. Structural debt is being paid down faster than it accrues.

## Overengineering Risk

Not overengineered; the macro budget is spent where repetition was a proven bug source. `define_options!` (`types.rs:286-344`) is 60 commented lines generating mechanical fan-out. `add_cell_cf!`/`add_viz_cf!`/`make_rule!` are forced by rust_xlsxwriter's trait-less CF types; each <12 lines. No config for things that never change; `PARALLEL_CHUNK_ROWS` and `MAX_SAFE_INT` are correctly constants with rationale comments. The ~6 families of per-feature `*_field` shims over `extract_opt` are mild residual duplication, not overengineering — they bake feature context into error messages, which is this project's operability signature. Leave alone.

## Underengineering Risk

- Long functions are linear ladders, not tangles: `write_py_value_with_format` (~170 lines, `convert.rs:357-525`) is a documented type-dispatch sequence; `apply_worksheet_features` (~180 lines) is a flat `if let Some → apply_x` sequence. Noted, not flagged.
- The genuine untyped-blob instance is the `Py<PyAny>` config maps crossing extract→apply (see abstraction audit).
- Real copy-paste debt: the df/dfs finalization duplication (W7) and the polars/pandas inner write loops (`convert.rs:606-644` vs `:645-684` — same body, different row acquisition; width-tracking changes must be made twice, three times counting the header path).
- `parse_format_dict(py, dict, include_column_options: bool)` is a textbook boolean mode-switch, but with two named wrappers and one extra key it's below the enum threshold.
- Deepest nesting: the rich-text extractor (`extract.rs:417-456`, 5-6 levels). Followable; extract a helper if touched again.

## Test Assessment

**What's good**: tests largely assert behavior at the right altitude — `test_media.py::TestCharts` inspects `xl/charts/chart1.xml` for actual data ranges and series counts; `test_conditional_formats.py` asserts rule type/operator/formula including the numeric-vs-string distinction; `test_error_paths.py` (19 tests) matches message content; the Rust `convert.rs` guard test forces a deliberate safe-vs-skipped decision per new option.

**Weak assertions (cited)**:
- `test_core.py::TestCsvConversion::test_csv_special_values` — opens the workbook, asserts nothing about cell contents. The one NaN/Inf/empty-CSV test verifies nothing.
- `test_media.py::TestImages::test_image_simple_path` / `test_image_with_options` — assert only `os.path.exists`. An image that never embeds would pass.
- `test_conditional_formats.py` color-scale/data-bar/icon-set tests — assert only `len(ws.conditional_formatting) > 0`.
- `test_validations.py` — all four tests assert existence only; `sqref` (the applied range) is never checked, so an off-by-one in `data_start_row` is invisible.
- `test_core.py::test_csv_parallel` uses 100 rows against `PARALLEL_CHUNK_ROWS = 10_000` — the multi-chunk path with the row-offset arithmetic is never exercised by any test.

**Proposed test cases** (for behavior-changing findings and top gaps):

| # | Gap | Test |
|---|-----|------|
| 1 | B1 datetime corruption | `test_non_ns_datetime64_roundtrip`: df with `datetime64[us]` column containing `3000-01-01`; assert the written cell reads back as year 3000 (or a clear `ValueError` — but not year 1830). |
| 2 | W1/W2 int boundaries | `test_int_beyond_i64_string_fallback`: `2**63 + 1025` → string cell; `test_i64_min`: `np.int64` min via numpy array and CSV → string cell, no panic. |
| 3 | W8 / 0.15.2 defined-names fix | `test_defined_name_empty_local_raises`: `defined_names={"Sheet1!": "..."}` → `pytest.raises(ValueError)`; both `df_to_xlsx` and `dfs_to_xlsx` (the logic is duplicated — W7). |
| 4 | W8 / 0.15.2 multibyte fix | Rust: `sanitize_table_name(&"é".repeat(300))` returns ≤255 chars without panic. |
| 5 | Parallel chunk seam | `test_csv_parallel_multi_chunk`: 25,000-row CSV, `parallel=True`; assert `A25001` equals the last CSV value. |
| 6 | W4 per-sheet keys (after fix) | `test_per_sheet_unknown_option_raises`: `{"freez_panes": True}` → `ValueError` listing valid keys. |
| 7 | Image embedding | assert `xl/media/image1.png` in `zipfile.ZipFile(path).namelist()`. |
| 8 | Validation range | `test_validation_sqref_and_header_false`: 2-row df → `dv.sqref == "A2:A3"`; with `header=False` → `"A1:A2"`. |
| 9 | CLI | Rust `tests/cli.rs` with `assert_cmd`: valid run → exit 0, stdout `OK 2 1`; invalid `--date-order` → exit 1. CLI currently has zero tests anywhere. |

## LLM Code Smell Scan

This codebase has been actively *de-smelled* — the 0.15.2/0.15.3 refactors removed exactly the duplication-with-drift an LLM accretes, and there is no module-level state anywhere. Remaining:
- **Duplication-with-drift (fix)**: the `defined_names` block, `lib.rs:585-604` vs `convert.rs:996-1012` — the only true near-copy left in `src/` (W7).
- **Thin wrappers (keep)**: per-feature `*_field` helpers — borderline, but they carry feature context into errors.
- **CI near-copies (consolidate)**: `.github/workflows/ci.yml` has three ~50-line `python-test*` jobs differing only in `runs-on` and matrix. Collapse into one job with an `os` matrix + `include`.
- **Test boilerplate (consolidate, low priority)**: the `get_temp_path()/try/finally os.unlink` pattern is hand-repeated ~230 times; a `tmp_xlsx` yield-fixture would delete several hundred lines. Also: importing pandas/pytest *through* `tests/helpers.py` is odd indirection.
- **Alias proliferation as quiet API entropy**: series values accept `values_range` | `values` | `data_range` (`apply/media.rs:406-414`); CF criteria accept four spellings each. Every alias is permanent public API maintained in match arms forever. Stop adding them.

## Operability Assessment

Strong. Error messages consistently carry identifiers and context (`validations['Score']: 'min' must be an integer, got str`; CSV errors with row numbers; write errors with (row, col)) and unknown-value errors enumerate valid alternatives. The constant-memory RuntimeWarning demonstrably reaches Python users, is derived from the option set so it can't go stale, and lives next to the skip. Holes: W4 (the one silent spot in a strict surface) and the stale "silently" wording (N4). Everything maps to `ValueError` including I/O failures — docs promise it, so fine, but `OSError` for file problems is worth a thought before 1.0. No logging — correct for a library.

## Documentation Assessment

Stubs are in sync with the Rust acceptors — verified by spot-check, and this is genuinely well done: every kwarg in both PyO3 signatures appears in `__init__.pyi` with matching defaults; `ImageOptions`/`TextboxOptions`/`ChartOptions`/`ValidationType` keys and Literals match the Rust key lists arm-for-arm including aliases. README/CHANGELOG/ROADMAP/BUILD are accurate; version 0.15.3 is consistent across `Cargo.toml`, `pyproject.toml`, lock, CHANGELOG, and tag. Fix N4 (stale "silently disables") and document the empty-dict per-sheet semantics: `{"comments": {}}` does **not** suppress a global `comments` (verified) — either document "empty dicts fall back to globals" or support explicit-empty override the way `table_style` supports explicit `None` (`extract.rs:104-110`).

## Architectural Drift Analysis

- **Boundary erosion**: one concrete back-edge (`apply/cells.rs:3 → convert`), introduced with the newest write-path feature — a classic erosion signature. The team actively fights these arrows (0.15.3 moved `pydict_to_hashmap` for exactly this reason); this one was missed.
- **Local exceptions becoming systemic**: none found — workarounds are not accumulating.
- **Contradictory patterns**: strict unknown-key rejection everywhere vs. silent key-dropping in `extract_sheet_info` (W4); explicit-off semantics for `table_style` vs. no off-switch for complex per-sheet options.
- **Entropy indicators**: `convert.rs` (1075) and `media.rs` (582) are the two growing files; `media.rs`'s name no longer describes its contents. Alias proliferation (see LLM scan). Counter-signals are strong: features.rs split (0.12.4), facade splits (0.15.0), `extract_opt` dedup (0.15.2), `define_options!` (0.15.3) — entropy is *decreasing* release over release.

## Strategic Recommendations

1. **Fix the `convert.rs` numeric/date boundary cluster (B1, W1, W2, W3) as one unit, with the regression tests from the table above** — Scope: small — Why: all four live within ~150 lines of each other, all are silent-wrong-output bugs, and B1 is reachable from default pandas 3 behavior. This is release-blocking; everything else is not.
2. **Close the per-sheet validation hole and add its drift guard (W4)** — Scope: small (~30 lines) — Why: eliminates the only remaining silent-drift path *and* a real user-facing typo trap, and restores the strict-validation contract the CHANGELOG advertises.
3. **Unify the df/dfs pipelines, then split the writer primitives out of `convert.rs` (W7 → W6)** — Scope: medium — Why: kills the verbatim duplication, halves per-option wiring in `lib.rs`, and breaking the `apply ↔ convert` cycle restores strict unidirectionality (`types ← parse ← write ← {apply, extract} ← convert ← lib`). Split `charts.rs` out of `media.rs` while in there, before sparklines land (N3).

Priority: #1 first — it's the only item where users get wrong data today.

## What's Actually Good

Fine, credit where due. The `define_options!` macro plus the constant-memory guard test is genuinely smart engineering — it converts "forgot a step" from silent misbehavior into a compile error or failing test, and it *stops at the right boundary* instead of trying to own everything, which is the part most people get wrong. The error messages are better than most commercial libraries: they name the feature, the key, the type you gave, and the valid alternatives. The chart and conditional-format tests assert against the actual XML inside the zip instead of "didn't crash". The CHANGELOG's one-debt-paydown-per-release habit is the rarest thing in this review: a codebase whose entropy is going *down*. And `cargo audit` is clean, CI covers three OSes, and the release pipeline uses PyPI trusted publishing. Now go fix the datetime corruption before someone's year-3000 bond maturity schedule ships as 1830.
