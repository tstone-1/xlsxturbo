# Strategic Recommendations Plan

Date: 2026-05-14
Source: `docs/reviews/2026-05-14_124318_codebase_deep_mine.md`

## 1. Make Validation Config Strict

Status: implemented in this session.

Scope:
- Reject unknown keys in validation configs.
- Reject wrong-type `min` and `max` values instead of silently defaulting.
- Keep missing or `None` `min`/`max` values as intentional defaults.
- Cover `list`, `whole_number`/`whole`/`integer`, `decimal`/`number`, and `text_length`/`textlength`/`length`.

Verification:
- Add regression tests for wrong-type `min` and typoed unknown keys.
- Run `cargo test --lib --no-default-features`.
- Run `cargo clippy --lib --no-default-features -- -D warnings`.
- Rebuild the extension and run `pytest tests/`.

## 2. Split Python Tests By Feature Family

Status: implemented in this session.

Goal:
Reduce the navigation and review cost of the current monolithic `tests/test_features.py` while preserving coverage.

Proposed structure:
- `tests/test_core.py`: basic DataFrame/CSV conversion, type handling, version, error basics.
- `tests/test_formatting.py`: header formats, column formats, borders, alignment, row/column dimensions.
- `tests/test_annotations.py`: merged ranges, hyperlinks, comments, defined names.
- `tests/test_validations.py`: validation rules, validation errors, strict schema behavior.
- `tests/test_conditional_formats.py`: color scales, data bars, icon sets, cell rules.
- `tests/test_media.py`: images, checkboxes, textboxes.
- `tests/test_cells.py`: arbitrary cell writes.
- `tests/test_constant_memory.py`: supported/disabled features and warnings.
- `tests/conftest.py`: shared imports, temp path helper, openpyxl guard, workbook loading helpers.

Execution plan:
1. Move shared helpers and imports into `tests/conftest.py`.
2. Move one feature family at a time, running that subset after each move.
3. Keep class and test names stable unless a rename improves failure readability.
4. Finish with the full `pytest` suite to confirm discovery and fixtures behave the same.

Acceptance criteria:
- Full test suite passes with the same behavioral coverage.
- No duplicate helper code across split files.
- Test files stay feature-focused and comfortably searchable.

## 3. Split `parse.rs` Into Focused Modules

Status: implemented in this session.

Goal:
Break the remaining parser/format utility drawer into modules with clearer ownership.

Proposed structure:
- `src/parse/mod.rs`: facade re-exporting the crate-local parser API.
- `src/parse/cell_refs.rs`: `parse_cell_ref`, `parse_cell_range`.
- `src/parse/colors.rs`: `parse_color`, `parse_color_enum`.
- `src/parse/formats.rs`: format option parsing, border/alignment helpers, column format building.
- `src/parse/tables.rs`: table style parsing and table-name sanitization.
- `src/parse/values.rs`: CSV scalar parsing, date/datetime conversion, date-order tests.
- `src/parse/patterns.rs`: wildcard column-name matching.

Execution plan:
1. Create the module shell and move one cohesive section at a time.
2. Keep public crate-local function names stable so callers do not churn.
3. Move existing Rust unit tests with their owning modules.
4. Run `cargo fmt`, `cargo test --lib --no-default-features`, and clippy after each chunk.

Acceptance criteria:
- No behavioral changes.
- `src/parse/mod.rs` is a small facade, not a new dumping ground.
- Existing parser tests continue to pass through the facade; moving those tests into their owning modules can be done later if desired, but the runtime code is now split by responsibility.

## Priority

Validation strictness is first because it is correctness. Test splitting is second because it reduces ongoing maintenance cost without touching runtime behavior. `parse.rs` splitting is third because it is structural cleanup with more merge risk.
