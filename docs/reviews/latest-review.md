# Code Review

**Date**: 2026-05-14
**Scope**: codebase
**Depth**: deep
**Mode**: mine (elite)
**Reviewer**: Senior Dev (automated)
**Tech Stack**: Rust library/CLI with PyO3 Python extension; Python package stubs/tests; pandas/polars/openpyxl test surface. Generated/cache dirs excluded.

## Verdict: CHANGES REQUIRED

Verdict definitions:
- **APPROVE**: No blockers. Warnings may exist but are contained.
- **CHANGES REQUIRED**: Blockers exist but fixes are local/contained. The approach is sound.
- **REJECT**: Systemic risk requiring re-architecture.

The core writer is not a mess. The feature surface is broad but mostly grounded in direct extraction -> write -> apply phases. The problem is that the multi-sheet per-sheet option parser quietly drops invalid option containers, which is exactly the kind of "worked fine, produced wrong workbook" bug that ruins your day.

## Maintainability Scorecard

| Dimension            | Score (0-5) | Notes |
|----------------------|:-----------:|-------|
| Cohesion             | 3 | Clear file roles, but `apply.rs` and `parse.rs` are getting chunky. |
| Coupling             | 3 | Reasonable internal direction: `lib`/`convert` orchestrate, `extract`/`parse`/`apply` support. |
| Abstraction quality  | 3 | `ExtractedOptions`/`EffectiveOpts` earn their keep; tuple aliases are weak contracts. |
| Complexity           | 2 | Public API arity and option-matrix complexity are now the tax collector. |
| Test robustness      | 4 | 212 Python feature tests and Rust parser unit tests; missing negative tests for per-sheet invalid options. |
| Operability          | 3 | Error context is usually good; warnings exist for constant-memory feature loss. |

## Blockers

- **[B1] Per-sheet option extraction silently ignores invalid containers** (`src/extract.rs:35`, `src/extract.rs:51`, `src/extract.rs:153`)
  The global `df_to_xlsx` path uses `require_dict`/`require_list` and raises on wrong types, but the per-sheet `dfs_to_xlsx` macros do `if let Ok(dict) = val.cast::<PyDict>()` / `if let Ok(list) = val.cast::<PyList>()` and otherwise do nothing. Same for `cells`. I verified this with:
  `dfs_to_xlsx([(df, "S1", {"validations": "not_a_dict"})], path)` -> no error, no validation.
  The same silent drop happens for `merged_ranges`, `cells`, and `header_format`. That is silent wrong output, not a preference debate.
  **Suggested fix:** Make the per-sheet extraction path use the same strict type checks as the global path. If a known option key is present and non-`None`, wrong container type should raise `TypeError` naming the key. Also validate that the third tuple element itself is a dict-like options object before reading keys from it.

## Warnings

- **[W1] `cells.wrap_text` accepts invalid types and silently disables wrapping** (`src/extract.rs:615`)
  `wrap_text` is parsed with `v.extract::<bool>().unwrap_or(false)`, so `{"wrap_text": "yes"}` succeeds and writes an unwrapped cell. The rest of this API generally rejects wrong types, and the test suite already expects `num_format` type errors. This one is just sloppy.
  **Suggested fix:** Replace the `unwrap_or(false)` with a typed extraction that returns `TypeError` with `cells['A1']: 'wrap_text' must be a bool`.
  **Test case:** `test_cells_wrap_text_wrong_type_raises` with `cells={"B1": {"value": "x", "wrap_text": "yes"}}`, expecting `TypeError`.

- **[W2] Constant-memory documentation is behind the runtime behavior** (`README.md:957`, `python/xlsxturbo/__init__.pyi:263`, `python/xlsxturbo/__init__.pyi:350`, `src/convert.rs:543`)
  Runtime warning logic disables `checkboxes`, `textboxes`, and `cells` in constant-memory mode, but README's constant-memory list omits checkboxes and textboxes. The type-stub docstrings also omit checkboxes/textboxes from the disabled list. Users relying on docs will discover this only at runtime.
  **Suggested fix:** Update README and `.pyi` docs to list every disabled feature from `write_sheet_data`: table style, freeze panes, autofit, row heights, formula columns, conditional formats, merged ranges, hyperlinks, comments, validations, rich text, images, checkboxes, textboxes, and cells.

## Nitpicks

- **[N1] README per-sheet option list omits newer options** (`README.md:338`)
  The per-sheet options list includes `images` and `cells`, but not `checkboxes` or `textboxes`, even though `extract_sheet_info` supports both at `src/extract.rs:149` and `src/extract.rs:150`.

- **[N2] Type aliases hide contracts that deserve names** (`src/types.rs:88`)
  `type MergedRange = (String, String, Option<HashMap<String, Py<PyAny>>>)` and friends make call sites positional and easy to misread. This is tolerable today, but future feature additions will keep making these tuples more cryptic.

- **[N3] The feature tests are huge enough to slow navigation** (`tests/test_features.py:1`)
  A 4,090-line test file is still readable by search, but it is past the point where focused files by feature would be kinder to future-you.

## Architecture Assessment

- **Boundaries & responsibilities**: `lib.rs` owns PyO3 public functions and API argument extraction, `extract.rs` converts Python structures to Rust-ish option structures, `convert.rs` writes DataFrame/CSV data and orchestrates features, `apply.rs` applies worksheet-level features, and `parse.rs` holds parsing/format utilities. That split mostly works.
- **Dependency direction**: Direction is mostly one-way: `lib -> extract/convert/types`, `convert -> apply/parse/types`, `apply -> parse/extract/types`. `apply.rs` importing `write_py_value_with_format` from `convert.rs` (`src/apply.rs:3`) is the one awkward arrow: apply code reaches back into conversion code for arbitrary cell writes.
- **Change-risk hotspots**: `src/apply.rs` at 1,381 lines, `src/parse.rs` at 1,179 lines, and `src/convert.rs` at 974 lines are the entropy magnets. The public API option matrix in `src/lib.rs:273` and `src/lib.rs:445` is also expensive to change safely.
- **Abstraction audit**: `ExtractedOptions`, `SheetConfig`, `WriteConfig`, and `EffectiveOpts` earn their keep by removing duplicated single/multi-sheet flows. The tuple aliases in `types.rs` do not earn much; they save typing but not cognitive load.
- **3AM test**: A tired engineer can follow CSV/DataFrame writing in 15 minutes. They will curse the option extraction/apply matrix because the same conceptual option appears in stubs, README, global extraction, per-sheet extraction, runtime warnings, and tests.

Overall: the architecture is serviceable. It is not over-layered LLM soup. It is under-modularized around feature families and under-strict in one per-sheet parsing path.

## Overengineering Risk

Low. The code mostly uses direct functions rather than decorative service layers. Do not respond to the blocker with a grand schema framework. A small typed helper for "present and must be dict/list" is enough.

## Underengineering Risk

The real underengineering risk is option-schema drift. The contract is duplicated across README, `.pyi`, `extract_options`, `extract_sheet_info`, `SheetConfig`, `ExtractedOptions`, `EffectiveOpts`, constant-memory warning logic, and tests. That is why per-sheet validation drift happened.

## Test Assessment

Verified:
- `cargo test --lib --no-default-features`: 52 passed.
- `cargo clippy --lib --no-default-features -- -D warnings`: passed.
- `uv run --extra dev python -m pytest tests/test_features.py`: 212 passed, 4 warnings.
- `cargo audit`: passed, no advisories reported.

Could not use plain `cargo test`: it fails at link time trying to build the PyO3 extension without Python symbols in this local macOS setup. The narrower lib test path is clean.

Missing tests:
- `test_dfs_to_xlsx_per_sheet_invalid_dict_option_raises`: use `{"validations": "not_a_dict"}` and expect `TypeError`.
- `test_dfs_to_xlsx_per_sheet_invalid_list_option_raises`: use `{"merged_ranges": {"A1:B1": "Title"}}` and expect `TypeError`.
- `test_dfs_to_xlsx_per_sheet_invalid_cells_option_raises`: use `{"cells": "not_a_dict"}` and expect `TypeError`.
- `test_cells_wrap_text_wrong_type_raises`: use invalid `wrap_text` type and expect `TypeError`.

## LLM Code Smell Scan

No pass-through repository/service/manager stack, thankfully. The LLM-ish smell is duplication-with-drift: every new option must be wired through several parallel structures. Fix the drift with small shared extraction helpers or generated-ish local tables, not with a giant config DSL.

## Operability Assessment

Error messages generally carry useful context: cell refs, option names, row/column write positions, and feature names. Constant-memory warnings are present and verified by tests. CLI output is intentionally simple. No server/runtime observability concerns apply.

## Documentation Assessment

README and stubs are mostly useful and example-heavy. They are not fully synced with newer checkbox/textbox support in constant-memory/per-sheet lists. Public Python type stubs are valuable and worth keeping strict.

## Architectural Drift Analysis

- **Boundary erosion**: `apply.rs` depends on `convert::write_py_value_with_format`, so arbitrary cell writes pull conversion behavior back into feature application.
- **Local exceptions becoming systemic**: Per-sheet option extraction has separate permissive macros instead of using the strict global extraction contract.
- **Emerging contradictory patterns**: Global options reject wrong container types; per-sheet options often ignore them. That is the main contradiction.
- **Entropy indicators**: Large modules, duplicated option lists, tuple aliases, and one 4K-line test file. Manageable, but the trend is obvious.

## Strategic Recommendations

1. **Unify strict option extraction for global and per-sheet paths** — Scope: medium — Why: fixes the blocker and prevents the next option from drifting.
2. **Split `apply.rs` by feature family** — Scope: medium — Why: conditional formats, shapes/images, validations, and cell writes are independent enough to isolate without ceremony.
3. **Promote tuple aliases to small structs where options are growing** — Scope: small/medium — Why: reduces positional mistakes and makes feature additions less cryptic.

Priority: do #1 first. It is user-visible correctness. The rest is maintenance debt, not a fire.

## What's Actually Good

The core conversion path is pretty sane. Integer precision guardrails exist (`src/convert.rs:24`), date pre-epoch handling exists (`src/parse.rs:626`), constant-memory incompatibilities warn (`src/convert.rs:504`), and the Python test suite actually verifies generated workbook content. Annoyingly competent in the places that matter.
