# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.17.2] - 2026-07-23

### Fixed
- Pure-bool-dtype DataFrame columns (a pandas column whose dtype is entirely `bool`, or a polars `Boolean` column) now write real Excel booleans instead of the numbers 1/0. The `np.bool_`/`np.bool` scalar these columns yield satisfies `__index__` and was previously falling through to the numeric fallback before the boolean check.
- `autofit=True` combined with a `column_widths` dict that names specific columns but has no `'_all'` key now still autofits the remaining columns, instead of silently leaving them at Excel's default width.
- The `dfs_to_xlsx` duplicate-table-name pre-check no longer false-positives when two empty DataFrames share a `table_name`/`table_style`, since neither actually creates a table (matching the existing `row_count > 0` gate on table creation).
- Save failures (`df_to_xlsx`, `dfs_to_xlsx`, and CSV conversion) now include the output path in the error message.
- `dfs_to_xlsx` write-phase errors and `constant_memory` disabled-feature warnings now name the failing/affected sheet.

### Changed
- A per-sheet `column_widths={}` no longer suppresses `autofit` for that sheet: an explicitly empty dict now disables only the widths option, matching the "empty dict/list disables this option" convention already used elsewhere. `column_widths` combined with `autofit=True` and no `'_all'` key now autofits every column not explicitly listed, rather than dropping autofit entirely once any `column_widths` key was present.
- Wrong-typed per-sheet scalar options (e.g. `{"header": "yes"}`) now raise a context-rich `TypeError` naming the option and the offending type, instead of a generic pyo3 conversion error.
- Unknown-key error phrasing is unified through a single shared helper across `apply/*` and `parse/formats.rs` (previously "unknown font option", "Unknown format option", and similar messages varied by feature).
- Format-dict errors (header format, column formats, rich text, merged-range/border formats) now carry the owning feature/cell-ref context, e.g. `column_formats['price_*']: ...` instead of a bare `format option '...'`.
- `dfs_to_xlsx` rejects an empty `sheets` list instead of silently writing a blank workbook.
- A textbox font flag explicitly set to `None` (e.g. `font={"bold": None}`) is now treated as absent, matching the None-means-absent convention used by every other optional field, instead of raising a type error.

### Internal
- Consolidated per-feature option-dict extraction (charts, sparklines, validations, images/checkboxes/textboxes, conditional formats, format dicts) behind a single `OptionMap` view in `types.rs`, removing roughly 400 lines of near-duplicate `<feature>_string_field`/`<feature>_bool_field` wrapper functions.
- The cell-ref-keyed feature maps (`comments`, `rich_text`, `images`, `checkboxes`, `textboxes`, `charts`, `sparklines`) now use `IndexMap` instead of `HashMap`, so their iteration order follows Python dict insertion order and generated workbook XML is reproducible byte-for-byte across runs with the same input.
- New `tests/test_option_coverage.py`: a completeness-guarded test that writes a minimal workbook per per-sheet option and asserts an observable effect, so an option that is accepted/extracted but never applied in `apply_worksheet_features` fails a test instead of shipping silently.
- CI clippy now runs with `--all-targets`; `actions/setup-python` bumped v6 -> v7 across CI and release workflows.
- Integration tests (`tests/test_integration.py`) now read back comments, validations, rich text, and images via openpyxl/zipfile instead of only asserting the output file exists.
- README documents that CSV/DataFrame string values are always written as literal text, never interpreted as formulas (formula injection note).

## [0.17.1] - 2026-07-13

### Fixed
- Multi-sheet exports reject duplicate effective Excel table names before writing the conflicting sheet, including collisions introduced by sanitization or case differences.
- Column formats, conditional formats, and validations reject patterns that match no columns instead of silently omitting requested behavior; column-format dictionaries are validated before target resolution.
- Sheet, merged-range, hyperlink, and rich-text tuples reject surplus elements instead of silently discarding them.
- The internal Rust library target has a distinct name, so Windows builds no longer produce colliding library/CLI PDB paths; the Cargo package, Python module, and CLI remain `xlsxturbo`.

### Changed
- Local uv development is pinned to Python 3.14.6, and BUILD.md consistently uses uv commands.
- CLI documentation includes parallel mode; historical Windows benchmark numbers are explicitly labeled as non-comparable because dispersion was not captured.

## [0.17.0] - 2026-07-02

### Added
- CLI: new `--parallel`/`-p` flag enabling multi-core CSV parsing, mirroring the Python `parallel=True` option.
- CI: new `python-lint` job runs the documented ruff, bandit, and pyright gates; the release workflow gained a `smoke-test` job that installs each built wheel on Linux/Windows/macOS and runs the full test suite before publishing.

### Changed
- An explicitly empty per-sheet dict/list (e.g. `{"comments": {}}`) now disables the corresponding global option for that sheet instead of silently inheriting it. Empty options no longer appear in the `constant_memory` disabled-features warning. Note: per-sheet `column_widths={}` selects the explicit-widths branch and therefore also suppresses `autofit` for that sheet (consistent with non-empty dicts).
- Chart `values`/`values_range`/`data_range` and `categories`/`categories_range` must be sheet-qualified (e.g. `"Sheet1!A2:A10"`). Bare ranges now raise a clear error instead of producing a misleading message (values) or silently rendering default 1..N axis labels (categories) - the same guard sparklines received in 0.16.1.
- Unknown keys in `conditional_formats` configs (per type), `comments` dicts, `checkboxes` dicts, and `cells` dicts are now rejected with an error listing the valid keys, matching the strict validation charts, sparklines, images, textboxes, and validations already had. Previously typos were silently ignored, yielding default-styled output.
- CSV string cells preserve leading/trailing whitespace instead of being silently trimmed. Type detection still ignores surrounding whitespace (`" 123 "` stays numeric) and whitespace-only cells remain empty; the CSV and DataFrame paths now agree on string content.
- `column_widths` keys are validated: negative, non-integer, or beyond-Excel-limit (> 16383) keys raise a clear error, and explicit keys beyond the data's column count are now applied instead of silently ignored.
- `csv_to_xlsx` releases the GIL for the duration of the conversion, so other Python threads stay responsive during large sequential or parallel conversions.
- `Cargo.lock` is now tracked for reproducible builds; the CI cargo cache keys (which hash the lockfile) are effective as a result.

### Fixed
- Dates from 1899-12-31 through 1900-02-28 are now written as text instead of date serials that rendered one day late in Excel (the 1900 leap-year bug; the first correctly representable date is 1900-03-01). Applies to the CSV, Python `date`/`datetime`, and numpy `datetime64` paths.
- Hex colors containing sign characters (e.g. `"#+12345"`) are rejected instead of parsing to an unintended color.
- Subclasses of `datetime.datetime`/`datetime.date` (e.g. pendulum or freezegun types) are written as real datetimes/dates instead of falling back to their string representation.
- Out-of-range `whole_number` validation bounds report the supported i32 range instead of a misleading "must be an integer, got int" type error.
- `date_order` error messages and the CLI help now list the accepted `european` alias; runtime docstrings list the `cell` conditional-format type (supported since 0.12.0); the path-argument error message clarifies that bytes paths are unsupported.
- Benchmarks: warmup now actually runs in `--json`/`--quiet` modes, matching the stated "median of N runs after warmup" methodology.
- Documentation accuracy: README CI-matrix wording and hyperlink example prose corrected; BUILD.md job lists match the workflows; dead CHANGELOG links for never-released tags removed.

### Internal
- Dependencies: rust_xlsxwriter 0.95 -> 0.96 (table-style variant list verified unchanged; all chart/sparkline/table XML assertions pass), zlib-rs 0.6.4 -> 0.6.5 via `cargo update`; CI `actions/cache` v5 -> v6. Supersedes Dependabot PRs #18 and #17.
- The 7-touchpoint feature-wiring checklist is now committed in `AGENTS.md`.
- New signature-parity test guards `df_to_xlsx`/`dfs_to_xlsx` kwarg drift; conditional-format cell-rule tests assert operators and formulas read back via openpyxl instead of file existence; textbox tests moved to their own class; shared `tmp_xlsx` fixture and parametrized constant-memory warning tests replace copy-pasted scaffolding.

## [0.16.2] - 2026-06-25

### Fixed
- `__init__.pyi` no longer advertises the option `TypedDict`/`Literal` helpers (`SparklineOptions`, `ChartOptions`, `ValidationType`, ...) as importable from the top-level package. They are stub-only types with no runtime object, so `from xlsxturbo import SparklineOptions` raised `ImportError` at runtime despite type-checking as valid. The stub now mirrors the real runtime surface; annotate option dicts by importing these from `xlsxturbo.xlsxturbo` under `TYPE_CHECKING`.
- Sparkline `style` values outside the `u8` range (e.g. `300`) or negative now report the documented "must be in the range 1-36" message instead of a generic integer error.
- `parse_cell_range` rejects reversed ranges (e.g. `"D10:A1"`) with a clear "first cell must precede the last cell" message instead of deferring to an opaque backend error (affects `merged_ranges` and grouped sparkline locations).
- Validation docstrings now note that type aliases (e.g. `integer`/`number`/`length`) are accepted, matching the README and type stub.

### Changed
- Unified the per-feature option-extraction error messages across charts, sparklines, images, textboxes, validations, conditional formats, and column formats via a single shared `extract_field` helper, so the same kind of error reads consistently regardless of which feature surfaced it.
- Centralized the integer-overflow-to-string policy in `write.rs` behind one predicate shared by every integer write path.

### Internal
- Added a guard test ensuring every `define_options!` feature option is also a recognized per-sheet option key, preventing a silent multi-sheet feature gap.
- Added a regression test pinning `formula_columns` behavior on an empty DataFrame (the formula column is skipped when there are no data rows).

## [0.16.1] - 2026-06-25

### Fixed
- Sparkline `range` and `date_range` now raise a clear error when not sheet-qualified (e.g. `"A2:C10"` instead of `"Sheet1!A2:C10"`). Previously a bare range failed deep in the writer with an opaque "Sparkline data range not set" message. Corrected the README/CHANGELOG/docstring examples, which used bare ranges.
- Sparkline `style` is validated to the documented 1-36 range instead of being silently ignored by Excel for out-of-range values.
- A grouped sparkline location must be a single row or column; a 2D block is now rejected rather than producing unexpected placement.

## [0.16.0] - 2026-06-25

### Added
- **Sparklines** via the new `sparklines` parameter on `df_to_xlsx` and `dfs_to_xlsx`. Sparklines are mini in-cell charts. A single-cell location key (e.g. `"D2"`) places one sparkline; a range key (e.g. `"D2:D10"`) places a grouped sparkline, one per row of the data range. The `range` key (data to plot, sheet-qualified like a chart range) is required. Supported options: `type` (`line`/`column`/`win_loss`), `style` (1-36), `markers`, `high_point`, `low_point`, `first_point`, `last_point`, `negative_points`, `show_axis`, `show_hidden_data`, `group_max`, `group_min`, `right_to_left`, `column_order`, `color` and the per-point/marker colors, `line_weight`, `custom_max`, `custom_min`, and `date_range`. Like charts, sparklines are skipped under `constant_memory=True`.
  - Example: `df_to_xlsx(df, "out.xlsx", sparklines={"D2:D10": {"range": "Sheet1!A2:C10", "type": "line", "markers": True}})`

### Changed
- Refreshed `uv.lock` to the latest compatible dependency versions (numpy 2.5.0, polars 1.42.0, pyarrow 24.0.0, maturin 1.14.1, plus dev tools).

## [0.15.5] - 2026-06-20

### Changed
- Updated `pyo3` to 0.29, clearing RUSTSEC-2026-0176 and RUSTSEC-2026-0177 (neither vulnerable API was reachable from this crate; the bump is dependency hygiene). `cargo audit` is clean.

### Fixed
- List-validation length checks and autofit width estimates now count characters instead of UTF-8 bytes, so multibyte values are no longer over-counted.

### Documentation
- Replaced the stale hard-coded performance multiplier in the module docstring with a pointer to the README's machine-labeled benchmark tables.
- Documented the `cells` alignment/wrap options (`align_horizontal`, `align_vertical`, `wrap_text`) in the `df_to_xlsx`/`dfs_to_xlsx` docstrings.
- Added contextual row/column/column-name information to previously bare cell-write and column-extraction error messages.
- Restored the changelog version link references (0.13.0 through current).

### Tested
- Added CLI integration tests (`tests/cli.rs`): exit codes, the `OK rows cols` stdout contract, and the invalid-`date_order` error path.
- Added a `version()` regression test asserting it matches the installed package metadata.
- Upgraded `rich_text`, `images`, `textboxes`, `validations`, conditional-format, and `freeze_panes` happy-path tests from existence/count smoke checks to content/semantic assertions.

## [0.15.4] - 2026-06-09

### Fixed
- Prevented non-nanosecond NumPy `datetime64` values from overflowing through an unsafe nanosecond cast and writing wrapped dates.
- Preserved oversized Python integers and `i64::MIN` as strings instead of rounded floats.
- Matched CSV behavior for pre-1900 DataFrame dates by writing them as strings instead of unsupported Excel serials.
- Rejected unknown per-sheet option keys in `dfs_to_xlsx` with a valid-key list.
- Accepted pandas DataFrames with non-string column labels by stringifying labels.
- Accepted `os.PathLike` values for path arguments.

### Documentation
- Updated constant-memory documentation to describe RuntimeWarning behavior and the supported safe options.

### Refactored
- Moved shared cell-writing primitives into a leaf `write` module and split chart application into `apply/charts.rs`.
- Shared defined-name validation and worksheet creation/write setup between single-sheet and multi-sheet paths.
- Replaced the same-typed `extract_options` positional parameter list with a named raw-options struct.

### Tested
- Added regression coverage for datetime/int boundary conversions, strict per-sheet keys, multibyte table-name truncation, empty defined names, and pathlib paths.

## [0.15.3] - 2026-06-04

### Documentation
- **Timezone-aware datetimes**: Documented that tz-aware datetimes are written as their local wall-clock value with the UTC offset dropped (Excel has no timezone concept), including a normalization workaround.

### Tested
- Added behaviour tests for the datetime paths: object-dtype `Timestamp` fractional seconds, timezone-aware wall-clock (characterization), and polars datetime columns.

### Refactored
- **Single-sourced the write-option structs** - A `define_options!` macro generates `ExtractedOptions`, `EffectiveOpts`, `as_effective`, and `merge_with` from one field list, removing ~70 lines of hand-maintained boilerplate where a transposed field name was an invisible bug.
- **`constant_memory` skip warning is now derived, not hand-listed** - The disabled-feature list comes from the generated option set minus an explicit safe-options list, and a guard test forces a deliberate safe-vs-skipped decision whenever a feature option is added.
- **Removed the last inward dependency arrow** - `pydict_to_hashmap` moved from `extract` to `types` so the `apply/` modules no longer depend back up on `extract`.

## [0.15.2] - 2026-06-04

### Fixed
- **`table_name` no longer panics on multibyte characters** - A long `table_name` containing non-ASCII letters (e.g. `"é"`) could split a UTF-8 codepoint at the 255-character cap and panic across the Python boundary. The name is now truncated on a character boundary and the call succeeds.
- **Empty `defined_names` keys raise `ValueError` instead of panicking** - A defined name that is empty (`""`) or has an empty local part (e.g. `"Sheet1!"`) now produces a clear `ValueError` instead of an uncatchable panic from the underlying writer.
- **Chart `series` items reject unknown keys** - A typo in a series-item option (e.g. `categorie_range` instead of `categories_range`) now raises a clear error listing the valid keys, matching the strict-validation behaviour of top-level chart options instead of silently dropping the value.

### Changed
- **Type stub `ChartType` lists all accepted aliases** - `col`, `donut`, and the `stacked_*` / `percent_stacked_*` spellings the parser already accepts are now part of the `ChartType` Literal so type-checkers accept them.
- Updated the package classifier from Beta to Production/Stable to match the documented and tested API surface.
- Added changelog and roadmap project URLs to package metadata.

### Refactored
- **Shared optional-field extraction** - The ~20 near-identical per-feature `*_field` extractor helpers across `apply/` and `parse/formats.rs` now delegate to a single `extract_opt` helper, removing duplicated get/None-check/extract/error logic while preserving every error message.
- **`constant_memory` skip-warning co-located with the skip** - The warning that lists features disabled by `constant_memory` now lives next to the code that actually skips them, so the two can no longer drift out of sync.

### Documentation
- Added README trust signals with CI, PyPI, Python version, and license badges.
- Added a project status section that summarizes tested platforms, versioning expectations, and API scope.
- Updated the roadmap so completed chart, checkbox, and textbox work no longer appears in planned sections.
- Clarified benchmark artifact output and the append-mode limitation in the README.

## [0.15.1] - 2026-05-25

### Documentation
- **Added README benchmarks for macOS** - Added a second 100,000 row x 50 column performance table from a MacBook run while preserving the existing Windows/AMD Ryzen reference table.
- **Updated datetime precision notes** - Documented that stored datetime serials preserve sub-second precision while the default display format shows whole seconds.

### Fixed
- **Preserved pandas `datetime64[ns]` columns** - Normal pandas datetime columns now write as Excel datetime cells, and `NaT` values remain empty, instead of falling back to strings from NumPy scalars.
- **Preserved fractional seconds in datetime serials** - CSV, Python, pandas, and polars datetimes now include sub-second precision in the stored Excel serial value.

### Dependencies
- **Completed benchmark dev dependencies** - Added `xlsxwriter` and `pyarrow` to the `dev` extra so the documented pandas+xlsxwriter and polars benchmark paths run after `uv sync --extra dev`.
- **Added maturin to dev dependencies** - `uv run maturin develop --release` now works after syncing the dev extra.

## [0.15.0] - 2026-05-16

### Added
- **Native Excel charts** - Embedded editable Excel charts via the new `charts` parameter. Supports common chart types (`bar`, `column`, `line`, `area`, `pie`, `doughnut`, `radar`, `scatter`, `stock` and stacked variants), single-series `data_range`/`values_range`, multi-series `series`, categories, title, axis names, size, offsets, style, data tables, and legend controls. Works in both `df_to_xlsx` and `dfs_to_xlsx` (including per-sheet options).

## [0.14.1] - 2026-05-14

### Fixed
- **`cargo test` works outside maturin builds** - `pyo3/extension-module` is now enabled by maturin instead of the default Cargo dependency path, fixing normal Rust test linking on macOS.
- **Validation configs now fail loudly on typos and wrong range types** - `validations` rejects unknown keys and present-but-invalid `min`/`max` values instead of silently defaulting to unbounded ranges.
- **Nested format containers now reject wrong types** - `column_formats`, `merged_ranges` formats, and `rich_text` tuple formats now raise clear errors when a format value is not a dict.
- **Per-sheet option extraction is strict** - `dfs_to_xlsx` per-sheet options now reject wrong container types instead of silently ignoring them.
- **`cells.wrap_text` validates types** - Wrong-type values now raise a clear `TypeError`.

### Refactored
- **Split feature application modules** - `src/apply.rs` is now a facade over focused `apply/` modules for annotations, cells, conditional formats, dimensions, formulas, media, rich text, and validations.
- **Split parser utilities** - `src/parse.rs` is now a facade over focused parser modules for cell refs, colors, formats, patterns, tables, and values.
- **Split Python integration tests by feature family** - The monolithic test file is now organized into focused test modules with shared helpers.

### Dependencies
- **Updated `rust_xlsxwriter`** - Bumped from `0.94` to `0.95`.
- **Refreshed development lockfile** - Updated Python dev dependency lock entries.

## [0.14.0] - 2026-04-21

### Added
- **Textboxes** - Floating text shapes via the new `textboxes` parameter. Simple form `{"B2": "text"}` for a bare string, dict form with `text` + `width`/`height` (pixels), `x_offset`/`y_offset` (pixels), `font` (sub-dict with `name`/`size`/`bold`/`italic`/`underline`/`color`), `fill_color`, `line_color`, and `alt_text`. Works in both `df_to_xlsx` and `dfs_to_xlsx` (including per-sheet options). Unknown top-level and font keys produce errors listing the valid options.
- **`parse_color_enum` helper** - Internal helper in `parse.rs` returning a `rust_xlsxwriter::Color` (wraps existing `parse_color`). Used by shapes; will be reused by sparklines and charts in upcoming releases.

## [0.13.0] - 2026-04-21

### Added
- **Checkboxes** - Interactive cell checkboxes via the new `checkboxes` parameter. Accepts `{"A1": True}` for a bare bool or `{"A3": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}` for a checkbox with an attached cell format. Works in both `df_to_xlsx` and `dfs_to_xlsx` (including per-sheet options).

## [0.12.5] - 2026-04-18

### Fixed
- **Format option typos and wrong types now raise errors** - Unknown keys in `header_format`, `column_formats`, and `conditional_formats[...]['format']` dicts (e.g. `"color"` instead of `"font_color"`) now produce a clear error listing the valid options; bool/string/number fields error on wrong types instead of silently being ignored. Previously typos and type mismatches were silent no-ops that produced unformatted output.
- **Image and validation options validate types** - `images[...]` options (`scale_width`, `scale_height`, `alt_text`) and validation `input_message`/`error_message`/`input_title`/`error_title` fields error on wrong types rather than silently dropping them. Unknown image options are rejected with a list of valid keys.

### Improved
- **CSV parallel mode peak memory reduced from O(file) to O(chunk)** - `csv_to_xlsx(parallel=True)` now streams the CSV in 10,000-row chunks (parse-in-parallel → write → drop → next chunk) instead of buffering the entire file twice in memory. Large CSVs no longer require several GB of RAM regardless of file size.
- **DataFrame write hot-path avoids per-cell Python type lookup for primitives** - Bool/int/float/string cells skip the `value.get_type().name()` PyO3 round-trip; only date/datetime/numpy-scalar/NA paths still need it. Measurable on wide numeric DataFrames.

### Refactored
- **Split `apply_single_conditional_format` into per-type helpers** - `apply_2_color_scale`, `apply_3_color_scale`, `apply_data_bar`, `apply_icon_set`, `apply_cell_conditional`. The cell-rule dispatch flattens from a 5-level-nested match to three sequential `match`es (blanks / text / range / single-value) via `add_cell_cf!` / `add_viz_cf!` macros. Adding a new criteria is now a 1-2 line change.
- **Rich text uses a dedicated narrow format parser** - `parse_rich_text_format` excludes `num_format` (meaningless for inline text runs) to match the `RichTextFormat` type stub contract.
- **Removed redundant `BufReader` layer** - `csv::ReaderBuilder::buffer_capacity(1MB)` replaces a `BufReader` sitting on top of an already-buffering reader.

### Dependencies
- **Minor bumps** - `csv` 1.3 → 1.4, `clap` 4.5 → 4.6, `rayon` 1.10 → 1.12, `indexmap` 2.7 → 2.14. Transitives: `hashbrown` 0.16 → 0.17 plus 15 other patch/minor updates via `cargo update`.

## [0.12.4] - 2026-04-03

### Fixed
- **Pre-1900 dates no longer produce invalid Excel serial numbers** - Dates before 1900-01-01 (both date and datetime) are now written as strings instead of negative serial numbers that render as `#####` in Excel
- **`constant_memory` warning now uses `RuntimeWarning`** - Previously emitted a generic `UserWarning`; now uses `RuntimeWarning` for proper filtering with `warnings.filterwarnings()`

### Refactored
- **Split `features.rs` into `extract.rs` + `apply.rs`** - Extraction (Python-to-Rust) and application (Rust-to-Excel) logic separated into focused modules (~500 and ~940 LOC respectively), improving maintainability
- **Added `SheetConfig::merge_with()` method** - Replaces 38 lines of repetitive per-sheet option merging in `dfs_to_xlsx` with a single method call; adding new options is now a one-place change
- **Moved unit tests to `parse.rs`** - Tests now live alongside the code they verify, following Rust conventions

### Tests
- **7 new Rust unit tests** - `naive_datetime_to_excel` (3), `parse_icon_type` (3), `naive_date_to_excel_pre_epoch` (1)
- **8 new Python integration tests** - CSV error paths, `constant_memory` warning emission, `defined_names` verification, `formula_columns` with `header=False` regression, pre-epoch date handling
- **Module-level openpyxl guard** - Tests now skip loudly via `pytest.mark.skipif` instead of silently passing without content verification

### Documentation
- **CHANGELOG** - Added missing version link entries for v0.10.5 through v0.12.3

## [0.12.3] - 2026-03-17

### Fixed
- **Linux x86_64 wheels now work on Python 3.9+** - Release workflow switched from `manylinux2014` (Python 3.8 only) to `manylinux_2_28` with `--find-interpreter`, producing proper `abi3` wheels instead of `cp38-cp38` wheels
- **Invalid datetime/date values now raise errors** - Previously, invalid dates from Python objects (e.g., month=13) silently fell through to string conversion; now returns a clear error message
- **Improved test assertion** - `test_parse_float` uses `assert!(matches!(...))` instead of bare `panic!`

### Improved
- **Reduced internal parameter counts** - Introduced `WriteConfig` struct to group scalar sheet configuration, reducing `write_sheet_data` from 11 to 5 parameters and `apply_worksheet_features` from 16 to 10
- **Contextual error messages throughout** - All `.map_err(|e| e.to_string())` calls replaced with descriptive `format!("Context: {}", e)` messages
- **CI cargo caching** - Added `actions/cache@v4` for Rust dependencies across all CI jobs
- **CI pip caching** - Added `cache: 'pip'` to all `setup-python` steps
- **CI platform coverage** - Windows and macOS now test Python 3.9 + 3.12 (was only 3.12)

### Documentation
- **README** - Added documentation for `defined_names` and `cells` parameters with usage examples
- **README** - Updated feature list with v0.11.0+ features (defined names, arbitrary cells, borders, alignment)
- **README/type stubs** - Added `formula_columns` and `cells` to `constant_memory` disabled features list
- **Type stubs** - Fixed `column_widths` and `row_heights` value types from `float` to `int | float`
- **Type stubs** - Documented validation `min`/`max` default behavior

### Tests
- **Per-sheet cells** - Added `TestCellsPerSheet` (4 tests) covering 3-tuple SheetOptions with cells
- **Cells formatting** - Added `TestCellsFormatting` (5 tests) covering alignment and wrap_text options

## [0.12.2] - 2026-03-16

### Fixed
- **`autofit=True` + `column_widths={'_all': N}` now caps instead of overriding** - columns are autofit to content then capped at N (`min(content_width, cap)`), matching the documented behavior (#13)

## [0.12.1] - 2026-03-16

### Fixed
- **Table creation crash when `header=False`** - Excel tables require a header row; table creation is now skipped when `include_header` is false, preventing `Table must have at least one row` errors (#12)

## [0.12.0] - 2026-03-16

### Added
- **Per-side border styles** - Fine-grained border control for `column_formats` and `header_format`
  - `border` now accepts string style names: `'border': 'thick'` (all 4 sides)
  - Per-side keys: `border_left`, `border_right`, `border_top`, `border_bottom`
  - Per-side keys accept `True` (= thin) or a style name string
  - `border_color` for setting border color (`'#RRGGBB'` or named color, applies to all sides)
  - 13 border styles: thin, medium, thick, dashed, dotted, double, hair, medium_dashed, dash_dot, medium_dash_dot, dash_dot_dot, medium_dash_dot_dot, slant_dash_dot
  - Works in both `column_formats` and `header_format`
  - Backward compatible: `'border': True` still works (thin, all sides)
  - Example: `column_formats={'col': {'border_right': 'thick'}}`
- **Text alignment** - `align_horizontal`, `align_vertical`, and `wrap_text` formatting options
  - Available in `header_format`, `column_formats`, `merged_ranges`, and `cells`
  - Horizontal: `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'center_across'`, `'distributed'`
  - Vertical: `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`
  - `wrap_text: True` enables text wrapping within cells
  - Example: `column_formats={'description': {'align_horizontal': 'left', 'wrap_text': True}}`
- **Rule-based conditional formatting** - `type: 'cell'` in `conditional_formats` for value-based highlighting
  - Comparison criteria: `equal_to`, `not_equal_to`, `greater_than`, `less_than`, `greater_than_or_equal_to`, `less_than_or_equal_to`
  - Range criteria: `between`, `not_between` (with `min_value`/`max_value`)
  - Text criteria: `containing`, `not_containing`, `begins_with`, `ends_with`
  - Special: `blanks`, `no_blanks`
  - `format` key accepts all column format options (bg_color, font_color, bold, border, etc.)
  - Multiple rules per column: pass a list of config dicts instead of a single dict
  - Example: `conditional_formats={'status': {'type': 'cell', 'criteria': 'equal_to', 'value': 'ERROR', 'format': {'bg_color': '#FF0000'}}}`

## [0.11.0] - 2026-03-15

### Added
- **Defined names** - `defined_names` parameter for workbook-level named ranges
  - Dict mapping name to Excel reference: `defined_names={"MyRange": "=Sheet1!$A$1:$D$100"}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()`
- **Arbitrary cell writes** - `cells` parameter for writing values to specific cells
  - Simple values: `cells={"B9": "Label", "B10": 42}`
  - With number format: `cells={"D6": {"value": "934728173849", "num_format": "@"}}`
  - Cells are written after DataFrame data, allowing overwrite of data cells
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides

## [0.10.6] - 2026-03-12

### Fixed
- **Polars DataFrame detection now checks module name** - `is_polars_dataframe` checks `__module__` instead of duck-typing attributes, preventing misidentification of non-DataFrame objects with `.schema` attribute (e.g., Pydantic models)

### Changed
- **CI uses pytest** - Integration tests now run via `pytest tests/ -v` instead of `python tests/test_features.py`, with proper test discovery and failure reporting
- **CI Python dependencies pinned** - `pandas>=2,<3`, `polars>=1,<2`, `openpyxl>=3,<4`, `pytest>=8,<9`, `maturin>=1.4,<2.0` to prevent unexpected breakage from upstream releases
- **`parse_table_style` uses macro** - Replaced 79-line match statement with `table_style_match!` macro; added version sync comment for `rust_xlsxwriter` 0.94
- **Dependencies** - Updated `rust_xlsxwriter` 0.93 -> 0.94, `actions/upload-artifact` v4 -> v7, `actions/download-artifact` v7 -> v8

### Refactored
- **Extracted `apply_worksheet_features` from `write_sheet_data`** - Feature application (table, formulas, conditional formats, freeze panes, widths, heights, merged ranges, hyperlinks, comments, validations, rich text, images) moved to a dedicated function with a single `constant_memory` early-return gate instead of 12 scattered checks
- **Removed redundant `constant_memory` parameter** from `apply_column_widths_with_autofit_cap` (caller already guards)

### Tests
- Added 22 new Rust unit tests: `parse_cell_ref` (basic, case-insensitive, max column, overflow, Excel max, row zero, empty, no row, no column), `parse_cell_range` (basic, invalid), `parse_color` (hex, named, invalid, whitespace), `sanitize_table_name` (valid, special chars, digit prefix, truncation, empty), `parse_table_style` (valid, invalid), `naive_date_to_excel` (epoch, known date), `DateOrder::parse`

## [0.10.5] - 2026-03-02

### Fixed
- **Formula columns overwrite data when `header=False`** - `apply_formula_columns` no longer hardcodes the formula header to row 0; headers are only written when `include_header=True`, preventing data loss when combining `header=False` with `formula_columns`
- **`parse_cell_ref` overflow on adversarial input** - Column letter fold now uses `checked_mul`/`checked_add` instead of wrapping arithmetic, returning a clear error on pathologically long column strings

### Changed
- **Minimum Python version raised to 3.9** - Type stubs use PEP 585 lowercase generics (`list[str]`, `dict[str, ...]`) which require Python 3.9+. Python 3.8 reached EOL in October 2024. Updated `requires-python`, PyO3 ABI tag (`abi3-py39`), and classifiers accordingly
- **`clap` is now an optional dependency** - CLI argument parser is gated behind a `cli` feature flag (enabled by default), reducing compile time for library-only builds (Python extension)
- **CI Python test matrix** - Integration tests now run on Python 3.9, 3.12, and 3.14 (previously only 3.12)
- **Completed Python docstrings** - `df_to_xlsx` and `dfs_to_xlsx` docstrings now document all parameters including `table_name`, `formula_columns`, `merged_ranges`, `hyperlinks`, `comments`, `validations`, `rich_text`, and `images`
- **`_all` width cap documentation** - Clarified that `_all` sets a uniform width rather than capping autofit results, since `rust_xlsxwriter` does not expose autofitted widths for reading

### Refactored
- **`write_py_value_with_format` reduced from 252 to ~90 lines** - Extracted `write_str`, `write_num`, `write_bool`, `write_int`, `write_float` helpers to eliminate 10x duplicated format/no-format dispatch
- **`extract_sheet_info` reduced from 170 to ~60 lines** - Introduced `extract_scalar!`, `extract_dict_field!`, `extract_list_field!` macros to replace 13 copy-pasted extraction blocks
- **`pydict_to_hashmap` helper** - Replaced 6 duplicated `HashMap<String, Py<PyAny>>` extraction blocks with a single reusable function
- **Explicit imports in `features.rs`** - Replaced `use crate::types::*` glob import with explicit type imports
- **Dependencies** - Updated indirect dependencies via `cargo update` (js-sys, wasm-bindgen, tempfile, zlib-rs)

## [0.10.4] - 2026-02-23

### Fixed
- **Boolean column formatting ignored** - Boolean values now correctly receive column formatting via `write_boolean_with_format` instead of being written without format
- **formula_columns not disabled in constant_memory mode** - Formula columns are now correctly skipped when `constant_memory=True`, matching the documented behavior
- **Cell reference column overflow** - `parse_cell_ref` now validates columns against Excel's maximum (XFD = 16384) using u32 intermediate arithmetic instead of silently wrapping u16
- **Unchecked arithmetic in formula/row operations** - Row and column index calculations now use `checked_add` to prevent silent overflow on extremely large datasets
- **Dead code** - Removed unreachable `extract::<bool>()` fallback in `write_py_value_with_format`

### Changed
- **Deduplicated write logic** - Extracted shared `write_sheet_data` function (~200 lines) used by both `convert_dataframe_to_xlsx` and `dfs_to_xlsx`, eliminating ~300 lines of duplicated code
- **Reference-based option merging** - New `EffectiveOpts` struct uses references instead of cloning, avoiding unnecessary allocations when merging per-sheet and global options in `dfs_to_xlsx`
- **`extract_sheet_info` refactored** - Now delegates to existing `extract_*` functions instead of reimplementing parsing inline
- **constant_memory warnings** - When `constant_memory=True` is used with incompatible options, a Python `warnings.warn()` is now emitted listing the disabled features
- **Dependencies** - Updated `pyo3` 0.28.1 → 0.28.2 (fixes RUSTSEC-2026-0013)
- **Metadata** - Added Python 3.13 and 3.14 classifiers to pyproject.toml

## [0.10.3] - 2026-02-16

### Fixed
- **Large integer precision loss** - Integers exceeding 2^53 are now written as strings instead of silently losing precision when cast to f64
- **Numpy int extraction order** - Numpy integer types (e.g. `numpy.int64`) now go through i64 extraction before f64 fallback, preventing precision loss for large values
- **Unchecked column index casts** - CSV sequential path now uses `u16::try_from` with clear error messages instead of unchecked `as u16` casts
- **CLI branding** - Replaced 4 remaining `fast_xlsx` references in `main.rs` with `xlsxturbo`
- **Undocumented `.unwrap()`** - Changed Excel epoch date `.unwrap()` to `.expect()` with explanation

### Changed
- **DataFrame type detection** - Extracted `is_polars_dataframe()` and `extract_columns()` helpers into `types.rs`, replacing 4 duplicated detection blocks across `convert.rs` and `lib.rs`
- **Documentation** - Added `constant_memory` disabled features list to Python docstrings; added "Known Limitations" section to README (datetime precision, large integers)
- **CI** - Bumped `actions/setup-python` from v5 to v6

### Tests
- Added `TestUnicodeAndSpecialData` class with 10 new tests: Unicode/CJK column names and data, emoji, mixed-type columns, None/NaT/pd.NA handling, all-None columns, large integer precision, CSV with BOM, CSV with CRLF, CSV with quoted delimiters, Polars Unicode
- Total: 93 Python integration tests, 12 Rust unit tests

## [0.10.2] - 2026-02-06

### Fixed
- **Wildcard pattern panic** - `matches_pattern("*")` no longer panics on lone `*` pattern
- **Silent datetime defaults** - Datetime/date attribute extraction now propagates errors instead of silently defaulting to 1900
- **Index overflow safety** - Row count and column count use checked casts (`u32::try_from`, `u16::try_from`) instead of `as` casts
- **PyPI URL in release workflow** - Fixed leftover `fast_xlsx` reference to `xlsxturbo`
- **Validation type aliases** - Added `whole`, `integer`, `number`, `textlength`, `length` aliases to type stubs
- **CHANGELOG accuracy** - `constant_memory` entry now lists all 12 disabled features

### Changed
- **Module split** - Split monolithic `lib.rs` into `convert.rs`, `parse.rs`, `features.rs`, `types.rs`
- **Deduplicated option extraction** - New `ExtractedOptions` struct reduces `convert_dataframe_to_xlsx` from 22 to 12 parameters
- **Merged format parsers** - `parse_header_format` and `parse_column_format` share a single `parse_format_dict` implementation
- **Dependencies** - Updated `pyo3` 0.27 -> 0.28, `rust_xlsxwriter` 0.92 -> 0.93
- **CI** - Added Python integration test job (83 tests with pandas, polars, openpyxl)
- **PEP 561** - Added `py.typed` marker file for type checker support

### Tests
- Added `TestConditionalFormatting` (5 tests), `TestConstantMemoryMode` (3 tests), `TestRowHeights` (3 tests)
- Upgraded ~10 shallow tests with openpyxl content verification (column widths, table names, header formats, validations, comments)
- Total: 83 Python integration tests, 12 Rust unit tests

## [0.10.1] - 2026-01-16

### Changed
- **Benchmark suite reorganization** - Moved benchmarks to `benchmarks/` directory
  - New `benchmarks/benchmark.py` - comprehensive comparison vs polars, pandas+openpyxl, pandas+xlsxwriter
  - Moved `benchmark_parallel.py` to `benchmarks/`
  - Removed obsolete `benchmark.py` (referenced old Rust binary)
- **README Performance section** - Updated with reproducible benchmark methodology
  - Changed performance claim from "~25x faster" to "~6x faster" (accurate for typical workloads)
  - Added disclaimer that results vary by system
  - Linked to Benchmarking section for running your own tests

## 0.10.0 - 2026-01-16

### Added
- **Comments/Notes** - Add cell annotations with optional author
  - Simple text: `comments={'A1': 'Note text'}`
  - With author: `comments={'A1': {'text': 'Note', 'author': 'John'}}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Data Validation** - Add dropdowns and constraints to columns
  - List (dropdown): `validations={'Status': {'type': 'list', 'values': ['Open', 'Closed']}}`
  - Whole number: `validations={'Score': {'type': 'whole_number', 'min': 0, 'max': 100}}`
  - Decimal: `validations={'Price': {'type': 'decimal', 'min': 0.0, 'max': 999.99}}`
  - Text length: `validations={'Code': {'type': 'text_length', 'min': 3, 'max': 10}}`
  - Supports input/error messages: `input_title`, `input_message`, `error_title`, `error_message`
  - Supports column patterns (like `column_formats`)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Rich Text** - Multiple formats within a single cell
  - Format segments: `rich_text={'A1': [('Bold', {'bold': True}), ' normal text']}`
  - Supports: `bold`, `italic`, `font_color`, `bg_color`, `font_size`, `underline`
  - Mix formatted and plain text segments
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
- **Images** - Embed PNG, JPEG, GIF, BMP images in cells
  - Simple path: `images={'B5': 'logo.png'}`
  - With options: `images={'B5': {'path': 'logo.png', 'scale_width': 0.5, 'scale_height': 0.5}}`
  - Options: `path`, `scale_width`, `scale_height`, `alt_text`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides

### Notes
- All new features are disabled in `constant_memory` mode (they require random access)
- Data validation list values are limited to 255 total characters (Excel limitation)

## [0.9.0] - 2026-01-15

### Added
- **Conditional formatting** - Visual formatting based on cell values
  - `2_color_scale`: Gradient from min_color to max_color
  - `3_color_scale`: Three-color gradient with min/mid/max colors
  - `data_bar`: In-cell bar chart with customizable color, direction, solid fill
  - `icon_set`: Traffic lights, arrows, flags (3/4/5 icons), with reverse and icons_only options
  - Supports column name patterns: `'price_*': {'type': 'data_bar', ...}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `conditional_formats={'score': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'}}`
- **Formula columns** - Add calculated columns with Excel formulas
  - Use `{row}` placeholder for row numbers (1-based)
  - Columns appear after data columns
  - Order preserved (first formula = first new column)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `formula_columns={'Total': '=A{row}+B{row}', 'Percentage': '=C{row}/D{row}*100'}`
- **Merged cells** - Merge cell ranges for headers, titles, and grouped labels
  - Uses Excel notation for ranges (e.g., 'A1:D1')
  - Optional formatting with HeaderFormat options (bold, colors, etc.)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `merged_ranges=[('A1:C1', 'Title'), ('A2:C2', 'Subtitle', {'bold': True})]`
- **Hyperlinks** - Add clickable links to cells
  - Uses Excel notation for cell reference (e.g., 'A1', 'B5')
  - Optional display text (defaults to URL if not provided)
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()` with per-sheet overrides
  - Example: `hyperlinks=[('A2', 'https://example.com'), ('B2', 'https://google.com', 'Google')]`

## [0.8.0] - 2026-01-15

### Added
- **Date order parameter** - `date_order` for `csv_to_xlsx()` to handle ambiguous dates
  - `"auto"` (default): ISO first, then European (DMY), then US (MDY)
  - `"mdy"` or `"us"`: US format where 01-02-2024 = January 2nd
  - `"dmy"` or `"eu"`: European format where 01-02-2024 = February 1st
  - Also available in CLI: `--date-order us`
- **BUILD.md** - Developer guide for building, testing, and releasing

### Fixed
- **Pattern matching order** - `column_formats` patterns now match in definition order (first match wins). Previously, HashMap iteration order was non-deterministic.
- **Empty DataFrame with table_style** - No longer crashes; tables are skipped when DataFrame has no data rows
- **Hex color validation** - Colors like `#FF` now raise descriptive error instead of silently misparsing
- **Invalid table_style validation** - Unknown styles now raise error instead of silently defaulting to Medium9
- **CLI division by zero** - Instant conversions now show "instant rows/sec" instead of "inf"

### Changed
- Uses `indexmap` crate to preserve pattern insertion order
- Updated `pyo3` 0.23 → 0.27, `rust_xlsxwriter` 0.79 → 0.92
- Added Dependabot for automated dependency updates

## [0.7.0] - 2025-12-28

### Added
- **Column formatting with wildcards** - `column_formats` parameter for styling columns by pattern
  - Wildcard patterns: `prefix*`, `*suffix`, `*contains*`, or exact match
  - Format options: `bg_color`, `font_color`, `num_format`, `bold`, `italic`, `underline`, `border`
  - Example: `column_formats={'price_*': {'bg_color': '#D6EAF8', 'num_format': '$#,##0.00', 'border': True}}`
  - Available in both `df_to_xlsx()` and `dfs_to_xlsx()`
  - Per-sheet column formats via options dict in `dfs_to_xlsx()`

## [0.6.0] - 2025-12-08

### Added
- **Global column width cap** - `column_widths={'_all': 50}` to cap all columns at a maximum width
  - Can be combined with specific column widths: `{0: 20, '_all': 50}` (specific overrides '_all')
  - Works with autofit as a cap: `autofit=True, column_widths={'_all': 30}` fits then caps
- **Table name parameter** - `table_name="MyTable"` to set custom Excel table names
  - Invalid characters are automatically sanitized (spaces/special chars become underscores)
  - Names starting with digits get underscore prefix (Excel requirement)
  - Per-sheet table names in `dfs_to_xlsx()` via options dict
- **Header styling** - `header_format={'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'}`
  - Supported options: `bold`, `italic`, `font_color`, `bg_color`, `font_size`, `underline`
  - Colors accept hex (`#RRGGBB`) or named colors (white, black, red, blue, etc.)
  - Per-sheet header formats in `dfs_to_xlsx()` via options dict
- Per-sheet options now support: `table_name`, `header_format`, `column_widths` with '_all'

### Changed
- `column_widths` parameter now accepts both integer keys (`{0: 20}`) and string keys (`{"_all": 50}`)

## 0.5.0 - 2025-12-08

### Added
- **Per-sheet options for `dfs_to_xlsx()`** - override global settings per sheet
  - Each sheet can now be a 3-tuple: `(df, sheet_name, options_dict)`
  - Options dict supports: `header`, `autofit`, `table_style`, `freeze_panes`, `column_widths`, `row_heights`
  - Old 2-tuple API `(df, sheet_name)` still works (backward compatible)
  - Example: `[(df1, "Data", {"table_style": "Medium2"}), (df2, "Instructions", {"header": False})]`
- `SheetOptions` TypedDict for type hints

### Changed
- `dfs_to_xlsx()` signature now accepts mixed tuple formats internally
- Updated type stubs with new `SheetOptions` class and updated `dfs_to_xlsx` signature

## [0.4.1] - 2025-12-07

### Fixed
- Updated type stubs to include v0.4.0 parameters (`column_widths`, `row_heights`, `constant_memory`)
- Cleaned up ROADMAP.md

## [0.4.0] - 2025-12-07

### Added
- `constant_memory` parameter - minimize RAM usage for very large files
  - Uses rust_xlsxwriter's streaming mode to flush rows to disk
  - Ideal for files with millions of rows
  - Note: Disables `table_style`, `freeze_panes`, `row_heights`, `autofit`, `conditional_formats`, `formula_columns`, `merged_ranges`, `hyperlinks`, `comments`, `validations`, `rich_text`, and `images`
  - Column widths still work in constant memory mode
  - Example: `xlsxturbo.df_to_xlsx(df, "big.xlsx", constant_memory=True)`
- `column_widths` parameter - set custom column widths by index
  - Dict mapping column index (0-based) to width in characters
  - Example: `column_widths={0: 25, 1: 15, 3: 30}`
- `row_heights` parameter - set custom row heights by index
  - Dict mapping row index (0-based) to height in points
  - Example: `row_heights={0: 22, 5: 30}`
- All new parameters available in `df_to_xlsx()` and `dfs_to_xlsx()`

## [0.3.0] - 2025-12-05

### Added
- `autofit` parameter - automatically adjust column widths to fit content
- `table_style` parameter - apply Excel table formatting with 61 built-in styles
  - Light styles: Light1-Light21
  - Medium styles: Medium1-Medium28
  - Dark styles: Dark1-Dark11
  - Tables include autofilter dropdowns and banded rows
- `freeze_panes` parameter - freeze header row for easier scrolling
- All new parameters available in both `df_to_xlsx()` and `dfs_to_xlsx()`

### Changed
- Updated type stubs with new parameters and documentation

## 0.2.0 - 2025-12-05

### Added
- `df_to_xlsx()` function for direct DataFrame export (pandas and polars)
- `dfs_to_xlsx()` function for writing multiple DataFrames to separate sheets
- `parallel=True` option for `csv_to_xlsx()` using multi-core processing
- Type preservation for DataFrame columns:
  - Python int/float → Excel numbers
  - Python bool → Excel booleans
  - datetime.date → Excel dates with formatting
  - datetime.datetime / pandas.Timestamp → Excel datetimes with formatting
  - None/NaN/NaT → Empty cells
- Type stubs for better IDE support
- rayon dependency for parallel processing

### Changed
- Updated documentation to include DataFrame and parallel processing examples

## [0.1.0] - 2025-12-04

### Added
- Initial release
- Python bindings via PyO3
- `csv_to_xlsx()` function for converting CSV files to Excel format
- Automatic type detection from CSV strings:
  - Integers and floats → Excel numbers
  - Booleans (`true`/`false`, case-insensitive) → Excel booleans
  - Dates (YYYY-MM-DD, DD/MM/YYYY, etc.) → Excel dates with formatting
  - Datetimes (ISO 8601) → Excel datetimes with formatting
  - NaN/Inf → Empty cells
  - Empty strings → Empty cells
- CLI tool for command-line usage
- Support for custom sheet names
- Verbose mode for progress reporting

[0.17.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.17.0
[0.16.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.16.2
[0.16.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.16.1
[0.16.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.16.0
[0.15.5]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.5
[0.15.4]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.4
[0.15.3]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.3
[0.15.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.2
[0.15.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.1
[0.15.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.15.0
[0.14.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.14.1
[0.14.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.14.0
[0.13.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.13.0
[0.12.5]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.5
[0.12.4]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.4
[0.12.3]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.3
[0.12.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.2
[0.12.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.1
[0.12.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.12.0
[0.11.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.11.0
[0.10.6]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.6
[0.10.5]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.5
[0.10.4]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.4
[0.10.3]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.3
[0.10.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.2
[0.10.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.10.1
[0.9.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.9.0
[0.8.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.8.0
[0.7.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.7.0
[0.6.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.6.0
[0.4.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.1
[0.4.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.4.0
[0.3.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.3.0
[0.1.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.1.0
