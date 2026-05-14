# Code Review

**Date**: 2026-05-14
**Scope**: codebase
**Depth**: deep
**Mode**: mine (elite)
**Reviewer**: Senior Dev (automated)
**Tech Stack**: Rust library/CLI with PyO3 Python extension; Python package stubs and integration tests; pandas/polars/openpyxl test surface. Generated/cache dirs excluded.

## Verdict: CHANGES REQUIRED

Verdict definitions:
- **APPROVE**: No blockers. Warnings may exist but are contained.
- **CHANGES REQUIRED**: Blockers exist but fixes are local/contained. The approach is sound.
- **REJECT**: Systemic risk requiring re-architecture.

The big strategic cleanup work is done: `apply` is split by feature family, dependency versions are current, and the tests pass. The remaining nasty bit is validation option parsing still doing the old "bad input? cool, pretend it was omitted" trick. You know better than this.

## Maintainability Scorecard

| Dimension            | Score (0-5) | Notes |
|----------------------|:-----------:|-------|
| Cohesion             | 4 | Feature application now has clear submodule ownership. |
| Coupling             | 3 | Direction is mostly clean; `apply::cells` still calls conversion write helpers. |
| Abstraction quality  | 4 | `ImageConfig`/`CheckboxConfig`/`TextboxConfig` and `EffectiveOpts` earn their keep. |
| Complexity           | 3 | Public API option arity is still high, but the internal feature split helps. |
| Test robustness      | 4 | Broad Python feature coverage plus Rust parser tests; a few negative schema cases remain missing. |
| Operability          | 3 | Error context is mostly good; warning behavior is tested; Python dependency audit tooling is unavailable here. |

## Blockers

- **[B1] Validation range options silently ignore wrong types and typos** (`src/apply/validations.rs:96`, `src/apply/validations.rs:123`, `README.md:602`)
  `whole_number`, `decimal`, and `text_length` validations extract `min`/`max` with `.and_then(|v| v.bind(py).extract().ok()).unwrap_or(...)`. That means `{"min": "zero"}` silently becomes `i32::MIN`, and a typo like `{"minimum": 0}` is ignored entirely. I verified this locally with `validations={"score": {"type": "whole_number", "min": "zero", "max": 100, "minimum": 0}}`; it succeeded and wrote a workbook. The README says validation unknown keys raise errors, and the whole point of validation is to enforce constraints. Silently widening the allowed range is wrong output.
  **Suggested fix:** Validate allowed keys per validation type, reject unknown keys, and use strict typed extraction for `min`/`max`. Missing `min` or `max` can still default if that is the intended API, but present-and-wrong must raise a `TypeError`/`ValueError` naming the column pattern and key.
  **Test case:** `test_validation_min_wrong_type_raises` with `validations={"Score": {"type": "whole_number", "min": "zero", "max": 100}}`, expecting an error mentioning `validations['Score']` and `min`.

## Warnings

- **[W1] Nested format containers are still silently dropped in a few extractors** (`src/extract.rs:236`, `src/extract.rs:311`, `src/extract.rs:443`)
  Most top-level option containers now fail loudly, but some nested format slots still use `if let Ok(dict) = ...` and otherwise treat invalid input as "no format". A non-dict `column_formats["A"]`, merged-range format, or rich-text tuple format can therefore produce unformatted output instead of a clear API error. This is the same bug pattern that was just removed from per-sheet options, just one level deeper.
  **Suggested fix:** For present optional nested format values, accept `None` as omitted, but reject any non-dict with context like `column_formats['A'] must be a dict`.

- **[W2] Documentation overpromises validation strictness** (`README.md:219`, `README.md:602`, `CHANGELOG.md:23`)
  Docs say validation options reject unknown keys or validate types. Runtime only type-checks `type`, `values`, and the message fields; range keys are permissive and unknown keys are ignored. This makes the docs look stricter than the code, which is a bad bargain because users trust validation config to protect spreadsheets.
  **Suggested fix:** Prefer fixing the runtime to match the docs. If permissive validation config is intentional, downgrade the README/CHANGELOG language, but that would be the weaker API.

## Nitpicks

- **[N1] Stale dependency comment after the version bump** (`src/parse.rs:27`)
  The comment says table style parsing is synced with `rust_xlsxwriter 0.94`, but `Cargo.toml:25` now uses `0.95`. Tiny thing, but stale comments are how future-you gets suspicious of everything.

- **[N2] `tests/test_features.py` is now a 4,153-line navigation tax** (`tests/test_features.py:1`)
  Coverage is good, but one giant file makes targeted review slower than it needs to be. Split by feature family to mirror `src/apply/`: formatting, annotations, media, validations, conditional formatting, constant memory, error paths.

- **[N3] `src/parse.rs` remains a mixed utility drawer** (`src/parse.rs:1`)
  It now handles table styles, cell refs, colors, format dicts, wildcard matching, CSV scalar parsing, date conversion, and unit tests. Nothing is on fire, but this is the next file that will get annoying.

## Architecture Assessment

- **Boundaries & responsibilities**: `lib.rs` owns PyO3 public entry points; `extract.rs` converts Python input into Rust-side config; `convert.rs` writes CSV/DataFrame data and orchestrates sheets; `apply/` owns feature-specific worksheet mutation; `parse.rs` owns parsing and format construction. The `apply` split paid down the previous biggest architectural debt.
- **Dependency direction**: Main direction is healthy: `lib -> extract/convert/types`, `convert -> apply/parse/types`, `apply -> parse/types`. The one awkward arrow is `apply/cells.rs` using `convert::write_py_value_with_format`, but that reuse is practical and not worth contorting right now.
- **Change-risk hotspots**: Validation schema handling is the current correctness hotspot. `parse.rs`, `convert.rs`, and `tests/test_features.py` are the remaining size hotspots.
- **Abstraction audit**: `ImageConfig`, `CheckboxConfig`, and `TextboxConfig` are clear wins over tuple aliases. `WriteConfig` and `EffectiveOpts` continue to reduce duplicated single-sheet/multi-sheet logic.
- **3AM test**: The feature-family split means a tired engineer can find most apply code quickly. They will still grumble at the duplicated API contract across README, `.pyi`, extractors, apply validators, and tests.

Overall: architecture is moving the right way. The remaining issue is not grand design; it is schema strictness.

## Overengineering Risk

Low. Do not invent a full schema framework just to fix validation. A small set of typed helpers for required/optional validation fields and per-type allowed-key checks is enough.

## Underengineering Risk

The option schema still lives in too many places: README, type stubs, `extract.rs`, `types.rs`, `apply/*`, runtime warnings, and tests. That is the systemic source of drift. The practical next step is not a new architecture astronaut layer; it is shared helpers and targeted negative tests for every user-facing config family.

## Test Assessment

Verified:
- `cargo audit` passed for Rust dependencies.
- `cargo test --lib --no-default-features` passed: 52 tests.
- `cargo clippy --lib --no-default-features -- -D warnings` passed.
- `uv run --extra dev python -m pytest tests/test_features.py` passed: 217 tests, 4 known warnings from openpyxl/constant-memory warning assertions.

Dependency audit:
- Rust: checked with `cargo audit`.
- Python: not checked; `uv pip audit` is unavailable in this installed `uv` (`unrecognized subcommand 'audit'`).

Missing tests:
- Wrong-type validation `min`/`max` should raise.
- Unknown validation option key should raise.
- Non-dict nested formats in `column_formats`, `merged_ranges`, and `rich_text` should raise.

## LLM Code Smell Scan

No fake service/repository layer nonsense. The recent module split is a real boundary, not decorative abstraction. The smell that remains is duplication-with-drift in the public option matrix; fix it with stricter helpers and tests, not more layers.

## Operability Assessment

Errors generally include enough context: cell refs, column patterns, option keys, and workbook operation failures. Constant-memory feature loss emits `RuntimeWarning`, which is testable and filterable. There is no logging surface, which is fine for a library/CLI that returns errors directly.

## Documentation Assessment

README and stubs are mostly aligned after the latest dependency/API updates. The validation strictness claim is currently false, and that is the only docs issue that rises above nitpick level.

## Architectural Drift Analysis

- **Boundary erosion**: The old `apply.rs` gravity problem has been fixed. Current drift is concentrated in schema validation, where extraction and application both own pieces of the same contract.
- **Local exceptions becoming systemic**: `if let Ok(dict)` silent-drop patterns remain in nested format parsing. That pattern already caused real bugs; don't let it survive in smaller corners.
- **Emerging contradictory patterns**: Most options now fail loud on wrong types; validation ranges still fail permissively. Pick one contract. The codebase has clearly chosen strictness everywhere else.
- **Entropy indicators**: `parse.rs` and `tests/test_features.py` are the remaining large files. No generated/vendor/cache dirs were included in the review.

## Strategic Recommendations

1. **Make validation config strict** - Scope: small - Why: Fixes the current blocker and aligns runtime with docs.
2. **Split Python tests by feature family** - Scope: medium - Why: Mirrors the new Rust layout and reduces review/navigation friction.
3. **Split `parse.rs` into parsing/format/style modules** - Scope: medium - Why: It is the next catch-all file after `apply.rs` was cleaned up.

Priority: validation strictness first. It is a correctness issue with a contained fix.

## What's Actually Good

The `apply` split is clean. The dependency update did not break the suite. The tests are broad enough to catch real workbook regressions, and the core data flow is still direct instead of wrapped in performative architecture. Grudgingly: this is now a maintainable codebase with one dumb validation corner left to sand down.
