# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- **`df_to_xlsx` and `dfs_to_xlsx` release the GIL while writing the archive.** Serialising
  and compressing the workbook touches no Python object, and it measured as the larger half
  of a call — so holding the interpreter lock across it meant threaded exports did not run
  in parallel at all. Four threads now finish a batch in about 43% of single-threaded wall
  time, against 98-102% before; single-threaded cost is unchanged. The gain plateaus near
  2.3x because reading values out of the DataFrame still holds the GIL and cannot be
  detached. `csv_to_xlsx` already released the GIL for its whole conversion and is
  unaffected. No API change, and nothing to opt into.

  `tests/test_concurrency.py` pins both halves: that concurrent writes off one shared frame
  produce correct workbooks — including a failure raised from inside the GIL-free region —
  and that four threads beat one by a wide enough margin. The timing case runs once per
  entry point, so removing either `py.detach` alone reddens exactly that entry point.

  Output is unaffected: the same workbook built before and after the change is identical in
  13 of its 14 archive members, differing only in `docProps/core.xml`, which records the
  creation time and has never been reproducible between runs.

### Documentation
- `docs/stability.md` said one `abi3` wheel per platform covers "every supported version".
  It does not cover free-threaded builds: on `python3.13t`/`python3.14t` there is no usable
  wheel, so `pip` falls back to the source distribution and needs a Rust toolchain. The page
  now says so. Measured on 3.14.7t and 3.15.0rc1t, against the ordinary 3.14.7 as a control.
  No code change — the wheels, and what they support, are unchanged.

## [1.3.0] - 2026-08-23

### Changed
- **`rust_xlsxwriter` 0.98.2 -> 0.99.0**, and the dependency floor moves with it. 0.99.0 is
  the release of upstream [#189](https://github.com/jmcnamara/rust_xlsxwriter/issues/189):
  it validates table names and defined names against Excel's rules, and it requires the two
  kinds to be unique *against each other*, which is what the new pre-check below reports.
  Some calls that used to succeed now raise, in every case because the workbook they wrote
  was one Excel offers to repair — a `defined_names` key containing a character Excel
  forbids (`My-Name`, `Total$`, `x=y`), one of Excel's logical constants (`TRUE`), or one of
  its internally reserved names (`_xlnm.Print_Area`). Table names are unaffected: they are
  sanitized before the crate sees them.
- The crate's message for an empty defined name changed with it, from `Name '' cannot be
  empty in Excel` to `Name cannot be blank`. The `defined_names['<key>']:` prefix that
  xlsxturbo puts in front of it is unchanged.

### Fixed
- **A `table_name` that Excel reads as a cell reference *with trailing text* is now
  sanitized too.** Excel stops reading an R1C1 form once it has the index and ignores
  whatever follows, so `"R2D2"` is the reference `R2` and `"R1_total"` is `R1` — both drew
  Excel's recovery prompt when written verbatim, as did `"C3PO"`, `"R1C1x"` and
  `"R1C16385"`. The check tested the whole string, so all of them passed through untouched
  and reached the saved workbook. They now gain the `_` prefix like `"Q1"`. Only the leading
  index has to exist, which is why `"R1C16385"` is prefixed although that column is past the
  grid; a name that never reaches an index is not a reference and is untouched (`"RCx"`,
  `"Rate1"`). The same rule now applies to `defined_names`, where a reference-shaped key is
  refused rather than rewritten — `"R1C1D"` was in the test suite as an *ordinary* name
  until Excel said otherwise.
- **A `table_name` of `"TRUE"` or `"FALSE"`, in any case, is now sanitized.** They are
  Excel's logical constants and it reserves them as names; a table named `"true"` drew the
  same recovery prompt as one named `"Q1"`. `"TRUE"` becomes `"_TRUE"`. Names that merely
  start with one (`"TRUEX"`, `"Falsehood"`) are untouched.
- **A `table_name` written in decomposed form is no longer mangled.** `sanitize_table_name`
  screens with an allowlist of `is_alphanumeric() || '_'`, and a combining mark is neither, so
  `"Verkäufe"` typed as `"Verka" + U+0308` reached the workbook as `Verka_ufe` and `"がくせい"`
  in NFD as `か_くせい` — silently, with no warning, for names Excel accepts. The name is now
  normalised to NFC before the screen runs, which repairs every mark that has a precomposed
  form. Normalisation is NFC and deliberately not NFKC: NFKC folds `U+FF21` FULLWIDTH A to
  ASCII `A`, which would turn a distinct name into the cell reference `A1` and then into
  `_A1`.

  Marks with **no** composed form are still rewritten — `"ไม่"` (Thai tone mark `U+0E48`) and
  `"हिन्दी"` (Hindi virama `U+094D`) become `ไม_` and `हिन_दी` — because closing that gap means
  widening the allowlist or inverting it to a denylist, which changes what is *accepted* and
  needs an audit in that direction. `marks_without_a_composed_form_are_still_rewritten` pins
  the current behaviour so the change cannot land unnoticed.

  Table names only. Defined names are refused rather than rewritten and never went through
  this screen.

### Added
- `unicode-normalization` as a direct dependency, for the NFC fix above. It was not
  previously in the tree, transitively or otherwise, and brings `tinyvec` and
  `tinyvec_macros` with it.
- **A pre-check for a table name that collides with a `defined_names` key**, raising
  `WorkbookValidationError` before anything is written. Excel requires the two kinds of name
  to be unique against each other and repairs a workbook carrying a table `Sales` beside a
  defined name `Sales` in any scope. rust_xlsxwriter 0.99.0 enforces it, but from inside
  `Workbook::save`, where this library classifies a failure as `FileError` — so without the
  pre-check the caller is told their *file* failed, in a message naming neither the sheet nor
  the option. Names are compared after sanitization and ignoring case, so `table_name="My
  Sales"` collides with a defined name `My_Sales`, and a sheet-scoped key such as
  `"Sheet1!Sales"` collides as much as a global one. A sheet that creates no table — no
  `table_style`, or an empty DataFrame — claims no name and cannot collide.

## [1.2.0] - 2026-08-17

### Changed
- **`rust_xlsxwriter` 0.98.1 -> 0.98.2**, and the dependency floor moves with it: the
  requirement is now `0.98.2`, not `0.98`, because the empty-name screen removed below only
  holds if the fix is present.
- **A `table_name` that collides with a cell reference is now sanitized like every other
  invalid name.** Excel forbids table names that address a cell, and rust_xlsxwriter 0.98.2
  does not check the rule (`Table::set_name` stores verbatim), so `table_name="Q1"` reached
  the saved workbook. `sanitize_table_name` now prefixes `_` — `Q1` becomes `_Q1` — for
  A1-style references inside the grid (case-insensitive, up to `XFD1048576`), the R1C1
  forms (`RC`, `R1C1`, with grid-bounded indices), and the selection shortcuts `R`/`C`,
  which Excel reserves separately. Names *beyond* the grid
  (`AAAA1`, `A1048577`) address no cell, so Excel accepts them and they pass through
  unchanged; the check runs on the already-sanitized string, so `1Q1 -> _1Q1` is not
  prefixed twice. A zero-padded row counts by the row it parses to — measured against
  Excel 16.112, which offers to repair a workbook whose table is named `A01` exactly as
  for `Q1`, so `A01` is prefixed while `A0` (row zero, addressing nothing) passes
  through. Rules table added to `docs/tables.md`. The missing validation is reported
  upstream as
  [rust_xlsxwriter#189](https://github.com/jmcnamara/rust_xlsxwriter/issues/189); the
  sanitize contract here stays either way, since it is `table_name`'s own behaviour and
  not a workaround.
- **A `defined_names` key that is a cell reference now raises `ConfigurationError`** instead
  of reaching the file — rejected rather than renamed, because a silently renamed defined
  name breaks every formula that references it. Sheet-qualified keys (`Sheet1!Q1`) are
  judged on their local part, matching what `define_name` itself validates. The A1/R1C1
  shape detector is written once, in `src/parse/cell_refs.rs`, and shared with the table
  sanitizer.
- **A Python `bool` is now rejected where a number is expected.** `column_widths={True: 20}`
  silently set column B's width and `textboxes={..., "width": True}` made a 1-unit box,
  because `bool` is an `int` subclass and pyo3's integer extraction accepts it — the trap
  the data plane already ordered its branches around (`write.rs`). Rejection covers
  `column_widths` keys (`ConfigurationTypeError`) and every `OptionMap` numeric field —
  chart/textbox/image sizes and offsets, `font_size`, sparkline weights, validation bounds
  (`ConfigurationError`, matching how those fields already classify a wrong-typed string).
  Genuine boolean options (`bold`, `wrap_text`, ...) are untouched.
- **`rich_text` with an empty segment list now raises `ConfigurationError` naming the cell**
  instead of being silently dropped — the same posture as an empty chart `series`, which has
  always errored.
- **A chart-level `data_range`, `values_range`, `values`, `name` or `series_name` beside
  `series` now raises `ConfigurationError`** naming every conflicting key at once, instead
  of being silently ignored while the `series` list won. `categories_range`/`categories`
  beside `series` is deliberately unaffected: the series branch reads it as the documented
  shared fallback for series items that carry none of their own.
- **CSV parse errors report a 1-based row number on both conversion paths.** Both the
  sequential and the parallel converter said `row 0` for the first record; they now agree
  with the `line N` that csv's own error text embeds. Exception classes unchanged
  everywhere above; all of these reject previously-accepted nonsense input and stay inside
  the hierarchy, so pre-0.19 `except ValueError`/`except TypeError` code keeps working.

### Fixed
- **Four apply-time type errors no longer embed a raw `TypeError:` inside their message.**
  `conditional_formats` `'type'`/`'criteria'` and `validations` `'type'`/`'values'` read
  `invalid 'type': TypeError: 'int' object is not an instance of 'str'` — a flattened
  `PyErr` rendered mid-string, the same defect class 1.1.1 fixed for chart series and
  textbox fonts, one layer down. They now go through the `OptionMap` accessors and read
  `'type' must be a string, got int`; a `values` list with a non-string item names the
  offending index (`item 0 is int`), and an explicit `None` for these keys now means unset.
  The guard in `tests/test_errors.py` that should have caught this matched `^\w+Error: `
  anchored to the message head — structurally blind to a class name at offset 42 — and now
  searches the whole message, with the four sites added to its probe matrix.

### Added
- **A source scan pins the "raise only through `crate::errors`" rule.** Zero violations
  existed, and zero tests would have noticed one: a `new_err` added to `lib.rs` would
  escape the hierarchy silently. `tests/test_errors.py` now scans `src/**/*.rs` for PyErr
  constructions (`new_err(`, `PyErr::new`, `PyErr::from_type(`, `create_exception!`)
  outside `src/errors.rs`, with controls asserting the scan found the sources and that the
  one allowed site actually matches — so a broken glob reads as red, not clean. Mentions
  (`is_instance_of::<PyKeyError>`) stay legal; the scan forbids constructions.
- **The two CSV pipelines now share one skeleton, and a test couples what remains apart.**
  Reader construction, workbook setup, formats, record validation and save were a ~50-line
  near-verbatim copy between the sequential and parallel converters — the one duplication
  with no guard. The shared parts are now written once; the row loops stay separate because
  routing the sequential path through the chunk machinery would add a per-row allocation to
  the hot path the parallel benchmark measures (stated at the site). A CLI test asserts
  both paths report the identical 1-based row for the same malformed file.
- **CI installs `maturin>=1.9,<2.0`, matching the build backend's floor.** Four workflow
  lines still said `>=1.4` while `pyproject.toml` pins 1.9 for the PEP 639 license-files
  behaviour; pip's newest-wins resolution masked the contradiction. A guard in
  `tests/test_ci_config.py` now compares every workflow maturin spec against the
  `[build-system]` floor, with a fail-on-zero-matches control — proven necessary: with the
  specs unpinned the floor assertion passed vacuously and only the control went red.

### Removed
- **The empty-`defined_names` screen is gone — upstream fixed the panic.** Until 0.98.1
  `Workbook::define_name` called `chars().next().unwrap()` on the local part of the name, so
  `""` and `"Sheet1!"` aborted the process instead of returning an error, and
  `apply_defined_names` rejected them before the call. 0.98.2 returns
  `ParameterError("Name '' cannot be empty in Excel")` instead
  ([rust_xlsxwriter#186](https://github.com/jmcnamara/rust_xlsxwriter/issues/186), filed
  2026-08-16 and released the next day). Both inputs still raise `ConfigurationError`, and the
  message still names the `defined_names` key that was rejected — the crate reports only the
  local part, so `"Sheet1!"` and `""` produce the same `Name ''` text from it, and the key
  comes from xlsxturbo's own wrapper. The two tests that covered the screen are replaced by
  three that cover the crate's behaviour, including one for that key; pinned back to 0.98.1
  they fail, because a Rust panic reaches Python as `pyo3_runtime.PanicException` rather than
  a `ValueError`.

## [1.1.2] - 2026-08-16

### Added
- **`THIRD-PARTY-LICENSES.md`, shipped in the wheel.** The wheel contains compiled code
  from 92 Rust crates under MIT, Apache-2.0, Unicode-3.0 and Zlib, and every one of those
  licenses requires the copyright notice to travel with the binary; `LICENSE` covers only
  xlsxturbo's own code, so until now the wheel carried none of them. maturin already wrote
  a CycloneDX SBOM into it, but an SBOM records *which* license applies and not the notice
  text those licenses ask for. Generated from the dependency tree by
  `python scripts/gen_third_party_licenses.py --write` (cargo-about, config in `about.toml`,
  template in `scripts/`), and delivered by PEP 639 `license-files`, which puts it and
  `LICENSE` in `<dist>.dist-info/licenses/`.
- `tests/test_third_party_licenses.py` compares the notice against `cargo metadata` in both
  directions, so a new dependency cannot ship without its notice and a hand-edited file is
  caught. It also pins `requires = ["maturin>=1.9"]` together with the `license-files`
  declaration: PEP 639 landed in maturin 1.9.0, and an older backend ignores the key and
  builds a wheel carrying neither notice with nothing going red.

### Removed
- **The `data_bar`-beside-sparkline refusal is gone — upstream fixed the defect.**
  `rust_xlsxwriter` 0.98.1 balances the `<ext>` elements it emits for that pair, so the two
  features can now be used together on one worksheet and the workbook opens. Filing the
  report is what ended it: [rust_xlsxwriter#185](https://github.com/jmcnamara/rust_xlsxwriter/issues/185)
  was opened on 2026-08-15 and released as 0.98.1 the next morning, after the workaround had
  been carried across two releases unreported. Removed with it: `reject_databar_with_sparklines`
  and its call in `convert.rs`, `tests/upstream_defect.rs` (which failed on the version bump
  exactly as it was written to), the `zip` dev-dependency it needed, and the "One combination
  is refused outright" section of `docs/errors.md`. `TestDataBarBesideSparklines` in
  `tests/test_options.py` replaces the guard's tests and asserts the opposite — that the pair
  writes well-formed worksheet XML; pinned back to 0.98.0 it fails with a `ParseError`, which
  is what shows it is measuring the fix rather than passing by default.

### Changed
- The build backend floor moves from `maturin>=1.4` to `maturin>=1.9`, for the PEP 639
  `license-files` support that ships the notice above. Source builds only; installing a
  published wheel is unaffected.
- `rust_xlsxwriter` 0.97.1 → 0.98.1, and eleven transitive crates refreshed within their
  declared ranges (cc, find-msvc-tools, futures-core/task/util, js-sys, portable-atomic,
  and the four wasm-bindgen crates). No API change.
- CI pins `github/codeql-action` at v4.37.6. Its `init`, `autobuild` and `analyze` entry
  points share one config and must move together; Dependabot files one PR per path, so
  each of its three PRs failed with `Loaded a configuration file for version '4.37.6', but
  running version '4.37.4'` and none could pass alone.
- Unreleased work now accumulates under `## [Unreleased]` in this file, and gains its
  version number and date only at release time. The previous convention wrote the version
  in early (`## [1.1.2] - Unreleased`), which `release-notes.sh` matches — it tests
  `## [<version>]` as a prefix and ignores the rest of the line — so a tag pushed before
  the date was filled in would have published a GitHub Release whose notes say Unreleased,
  with no job failing. `## [Unreleased]` matches no tag and fails the release job instead.
  `BUILD.md` step 2 and `tests/test_ci_config.py::TestChangelogHeadings` now say so.

## [1.1.1] - 2026-08-06

### Fixed
- **A `DATA_BAR` conditional format beside sparklines no longer writes a corrupt workbook.**
  The guard that refuses the databar+sparkline pair (a known rust_xlsxwriter 0.97.x defect —
  the writer emits unbalanced `<ext>` elements and Excel reports the file as damaged) matched
  its type string case-sensitively, while the dispatcher that applies conditional formats
  lowercases first. So `data_bar` was refused and `DATA_BAR` slipped past the guard, was
  applied as a real data bar, and produced exactly the silently-corrupt file the guard exists
  to prevent — accepted input, successful return, a workbook Excel calls damaged. The guard
  now normalizes case the same way the dispatcher does, and the guard's test covers the case
  axis as well as the spelling axis.
- **The `cells` option's nested `num_format`, `align_horizontal`, and `align_vertical` values
  are classified, and an explicit `None` reads as unset.** A wrong-typed value — say
  `cells={"A1": {"value": 1, "num_format": 42}}` — raised the binding's own bare `TypeError`,
  outside the documented exception hierarchy; it is now a `ConfigurationTypeError` naming the
  cell, the key, and the offending type (still a `TypeError`, so existing handlers keep
  working). Passing `None` for any of these keys, `wrap_text` included, previously raised the
  same bare error; it now means "not given", matching every other optional value in the API.
- **Apply-time errors for nested option dicts no longer carry a stray class-name prefix.**
  A non-string key in a chart `series` item or a textbox `font` dict produced a message
  beginning `ConfigurationTypeError: ...` on an exception that is not one — the class name of
  an internal error surviving as message text. The message now reads like every other
  configuration error. (The exception class itself is unchanged; aligning it with the
  extract-time `ConfigurationTypeError` changes the builtin base an `except` clause sees,
  which the stability policy reserves for 2.0.0.)
- **The duplicate-table-name pre-check folds names with full Unicode case rules.** It
  compared with an ASCII-only fold, so two table names differing only by non-ASCII case —
  `TÄBLE` and `täble` — passed the pre-check and collided only during the save, as a
  `FileError` naming neither sheet. The pre-check now uses the same case fold as the
  underlying writer's own uniqueness check and reports both sheet names up front.
- **`docs/errors.md` and the exception docstrings now describe save-time classification as it
  is.** Workbook rules the writer only checks while serialising — a duplicate sheet name, a
  chart range naming a sheet that does not exist — surface as `FileError`, and an image file
  that cannot be read surfaces as `ConfigurationError`; the documentation claimed otherwise
  in both directions. The docs and the `FileError`/`WorkbookValidationError` docstrings state
  the actual behavior, with a worked example of each. (Reclassifying the failures themselves
  is a builtin-base change and waits for 2.0.0.)
- **`uv.lock` is current again, and a test now keeps it that way.** The 1.1.0 release bumped
  the version and raised `requires-python` to `>=3.10` without re-running `uv lock`, so the
  tracked lockfile went on recording `xlsxturbo 0.21.0` plus resolution legs for Python 3.9,
  which the project no longer supports. Nothing read the lockfile in CI, so nothing noticed.
  The lockfile is regenerated, and `tests/test_ci_config.py` asserts the locked version
  matches `pyproject.toml` — that test fails the next release that skips the re-lock.
- **Two stale in-source documents told the next reader the opposite of the design.** The
  `ConvertError` rustdoc still described the `From<String>` default whose removal was the
  point of the 0.21.0 redesign, and the module docstring pointed benchmark readers at README
  tables that moved to the documentation site in 0.19.0. Both now match reality, as do
  `BUILD.md`'s release-workflow job lists, which had drifted behind the five-leg smoke test.

- **Every published wheel is now smoke-tested on a runner of its own architecture.** Two of
  the five — Linux `aarch64` and macOS `x86_64` — were built, published to PyPI, and never
  executed anywhere. Both are cross-compiled, and when the pipeline was written no hosted
  runner of either architecture existed. They do now (`ubuntu-24.04-arm` and
  `macos-15-intel`), so the release workflow installs and runs the full test suite against
  all five before publishing.

  Each leg also asserts the platform tag of the wheel it downloaded and prints the runner's
  own architecture. Without that, a mistyped artifact name silently installs some other wheel
  twice and two green legs look exactly like coverage.

  Nothing had compared the build matrix with the smoke-test matrix, which is why the gap
  survived every release. `tests/test_stability_policy.py` now does, in both directions, and
  ties the "Smoke-tested before publish" column on the
  [stability page](https://tstone-1.github.io/xlsxturbo/stability/) to the workflow — so a
  new build target fails the suite until it has a leg, and a leg naming an artifact no job
  uploads fails in an ordinary test run instead of after every wheel has built.
- **`ExportOptions` field types now match the signature they mirror.** Five fields disagreed
  with `df_to_xlsx`. Four had decayed to `Any` inside a container — `column_widths`,
  `merged_ranges`, `hyperlinks`, `checkboxes` — so a bundle gave an IDE nothing for exactly
  the options where the shape is hardest to remember, which is most of the reason the class
  exists.

  The fifth is the one worth naming: `row_heights` was `dict[int, float]` while the function
  accepts `dict[int, int | float]`. That direction is worse than `Any` — a checker rejected
  `row_heights={1: 30}`, which works at runtime, so correct code was reported as wrong.
- **`SECURITY.md` was three minor versions out of date.** It described the project as
  pre-1.0, named `0.18.x` as the supported line, and said what would happen "once 1.0 ships"
  — while the package was at 1.1.0. A support table that names an unsupported version as
  supported is worse than no table. It now states the 1.1.x line, and says plainly why a
  1.0.x security release is the one exception (a Python 3.9 interpreter resolves to 1.0.0
  and cannot take a fix shipped on 1.1).
- Issue-template version placeholders suggested `0.18.0` and `0.17.2`, including a worked
  benchmark example comparing two pre-1.0 releases.

### Changed
- **`dfs_to_xlsx`'s docstring documents `conditional_formats` pattern matching.** Both
  functions share the same implementation, which matches column names and wildcard patterns
  alike; the multi-sheet docstring said "column names" only, under-documenting a working
  feature.
- **Format-key lists are composed, not restated.** The cell-scope key list restated all six
  font-scope keys by hand; the scopes now compose, so a font-level key added once cannot
  diverge the rich-text and cell scopes silently. User-visible key lists in unknown-key
  errors are unchanged, including their order.
- Rust dependencies refreshed within their declared ranges: rust_xlsxwriter 0.97.1 (the
  databar+sparkline defect is verified still present there — the guard and its signal test
  stay), pyo3 0.29.2, clap 4.6.6, and the transitive tail. Test floor for polars raised to
  1.43.2 (Dependabot #31).

### Internal
- `tests/test_options_types_match_the_stub.py` compares every `ExportOptions` field
  annotation against the corresponding `df_to_xlsx` parameter in `xlsxturbo.pyi`, as source
  text, in both directions. Two surfaces describing the same options is the arrangement that
  drifts, and nothing was watching this pair — the field *set* was pinned, the field *types*
  were not.
- `tests/test_public_docs_are_current.py` — every correctness gate in this repository points
  inward, at options, types, errors and workflows. None read the public policy documents, so
  those drifted across five releases without a red build and an external reviewer found them.
  The new tests tie `SECURITY.md` and the issue templates to the declared version.

  Written twice, which is the part worth recording. The first version matched
  `placeholder:` line by line and could not fail its own control: a Python-version
  placeholder counted toward the population, so removing an xlsxturbo one left the count
  above threshold. It also could not see inside a block scalar, which is exactly where one of
  the stale references was hiding. Parsing the issue-form YAML and keying on the field `id`
  fixed both — and only then did the control go red on the mutation written for it.

## [1.1.0] - 2026-07-30

### Removed
- **Python 3.9 support.** It reached upstream end of life in October 2025. `pip` handles the
  drop without any action on your part: a 3.9 interpreter resolves to 1.0.0, which stays on
  PyPI and keeps working. Move to 3.10 or newer to receive further releases.

  A dropped Python version is a minor release under the
  [stability policy](https://tstone-1.github.io/xlsxturbo/stability/) published in 1.0.0, not
  a 2.0.0 event — the reasoning is on that page, and this is the first release to exercise it.

  The floor was not costing nothing. In the two days before this release it blocked pytest 9,
  numpy 2.1+ and polars 1.37+ from the test matrix — each requires 3.10 — and it forced
  `python/xlsxturbo/types.py` to spell every union `Union[str, int]` where `str | int` is the
  natural form.

### Changed
- Wheels are now `abi3-py310` rather than `abi3-py39`; one wheel per platform still serves
  every supported version, 3.10 through 3.14.
- `python/xlsxturbo/types.py` uses PEP 604 unions throughout, including module-level aliases
  such as `PathArg`. Type checking is unaffected; `typing.get_type_hints()` on the option
  shapes now works on every supported version rather than only from 3.10.
- Test-dependency floors raised to what 3.10 makes reachable: `pytest>=9.1.1`,
  `numpy>=2.2.6`, `pandas>=2.3.3`, `polars>=1.43.1`. Test-only; the published wheel has no
  runtime dependencies.
- **CI now runs the test suite against both pandas 2 and pandas 3.** The declared range was
  `<3` while pandas 3 was released and was what development actually ran on, so local runs
  and CI had silently stopped testing the same library. The ceiling moves to `<4` and a
  fourth `python-test` leg installs `requirements-test-pandas2.txt`, which layers the main
  file and overrides only the pandas ceiling — every floor stays single-sourced.

  Worth a job rather than an assumption because pandas 3 changed defaults this library is
  directly exposed to: Copy-on-Write, and PyArrow-backed strings as the default `str` dtype,
  in a library whose whole job is detecting types. Both majors pass, measured on each before
  the change was made.

### Internal
- **Two guards fired on this change, as designed, and that is the result worth recording.**
  The stability page's version table went red against the classifiers and the CI matrix, and
  the Dependabot hold registered in `tests/test_ci_config.py` went red because its reason —
  pytest 9 needs 3.10 — had expired. Neither was an obstacle to work around; each named the
  work that had to accompany the drop, which is what both were built for one day earlier.
- Guards that existed only for 3.9 were **deleted rather than weakened**: the two tests
  pinning the `Union[...]` spelling in `types.py`, and the ruff `keep-runtime-typing` setting
  that stopped `--fix` undoing them. The constraint is now enforced by the language version,
  which is stronger than a test.
- `HELD_BACK_MAJORS` is empty and the Dependabot `ignore` block is gone. The mechanism stays,
  restructured so it is not vacuous with nothing held: the live check is now that no
  major-version `ignore` may exist without a registered expiry condition, which needs no
  entries to do its job.

## [1.0.0] - 2026-07-30

**No behaviour changes. Nothing was renamed, removed, or reinterpreted.** 1.0.0 is the
version that stops reserving the right to.

The API surface was fixed over 0.19.0–0.21.0 — the exception hierarchy, the runtime option
types, `ExportOptions`, and the four design questions the review left open. What remained
for 1.0 was to say what is now promised, and to make the promise checkable rather than
aspirational.

### Added
- **[Stability and support](https://tstone-1.github.io/xlsxturbo/stability/)** — the public
  surface named exhaustively, what does and does not count as a breaking change, the
  deprecation policy, the supported Python and platform matrices, and what is guaranteed
  about generated files.

  The policy in one line: everything reachable from `import xlsxturbo` without a leading
  underscore is covered and will not break before 2.0.0. Anything removed gets a
  `DeprecationWarning` naming its replacement and its removal version, shipped for at least
  one minor release and at least six months, with removal only in a major.

  The page is checked against what it describes rather than kept in step by hand
  (`tests/test_stability_policy.py`): the Python table against the trove classifiers and the
  CI matrix in both directions, `requires-python` against the oldest row, the platform table
  against the wheel targets in `release.yml`, and the surface table against `__all__`.

- **A measured statement about output determinism.** Two exports of identical data are *not*
  identical files — `docProps/core.xml` records the creation time — but every other member
  of the archive is byte-identical across runs. Both halves are pinned by tests, so the
  caveat cannot rot into folklore and a future non-determinism elsewhere cannot hide behind
  it.

  Worth stating because the obvious measurement lies: two exports written inside the same
  clock second hash identically, which reads as full reproducibility. The test waits out a
  second deliberately.

### Changed
- `Development Status :: 5 - Production/Stable` is now accurate. It has been declared since
  well before this release while the version stayed on `0.x`, which is a contradiction —
  a 0.x version reserves the right to break anything.
- Test-dependency floors raised to the versions CI already resolves: `openpyxl>=3.1.5`,
  `pyyaml>=6.0.3`. Test-only, so the published wheel — which has no runtime dependencies —
  is unaffected.

### Known and unchanged
- **Python 3.9 is supported and is past upstream end of life** (October 2025). Dropping it is
  a user-visible change, not a tidy-up, so it will happen in a future minor release and be
  announced. The cost of keeping it is visible in `python/xlsxturbo/types.py`, which writes
  `Union[str, int]` where every other supported version would allow `str | int`.
- **`NaN`, `Inf` and `Infinity` as text still empty their cell.** Documented in
  [DataFrame export](https://tstone-1.github.io/xlsxturbo/dataframe-export/) since 0.21.0 and
  deliberately not changed here: fixing it means changing the value written for an input that
  already works, which is exactly what this release promises not to do outside a major.

### Internal
- `pytest` stays on 8, and the reason is recorded where it can expire. pytest 9
  requires Python 3.10 while this project supports 3.9, so the 3.9 CI jobs cannot resolve it
  at all; nothing in the suite is otherwise incompatible with it. The weekly Dependabot PR is
  silenced by an `ignore` entry — and because a silenced PR simply stops arriving, that entry
  is tied to `requires-python` by a test, in both directions: while 3.9 is supported the
  ignore must exist, and the moment the floor moves past 3.10 the ignore must be gone. Raising
  the Python floor now fails the suite until the hold is released with it.

## [0.21.0] - 2026-07-30

Four open design questions from the 1.0 review, decided rather than inherited. Every change
here is additive: no class was renamed or removed, no builtin base changed, and no existing
`except` clause means anything different than it did.

### Added
- **`OptionError` — one class meaning "anything wrong with what you passed."**
  `ConfigurationError` (bad value) and `ConfigurationTypeError` (wrong type) were siblings, so
  a caller who wanted both had to catch a tuple, while `XlsxTurboError` also swept up
  filesystem failures and unsupported DataFrames — which are not the caller's options at all.

  ```python
  try:
      xlsxturbo.df_to_xlsx(df, path, **user_supplied_options)
  except xlsxturbo.OptionError as exc:
      return {"error": f"bad export options: {exc}"}
  except xlsxturbo.FileError as exc:
      return {"error": f"could not write the file: {exc}"}
  ```

  Added by **reparenting, not renaming**: both classes keep their names and their builtin
  bases, and `OptionError` takes no builtin of its own — one there would land on both children
  and make every bad *value* a `TypeError` too.
- **`FileError.errno` is populated**, so a full disk can be told from a missing directory
  without matching on message text:

  ```python
  except xlsxturbo.FileError as exc:
      if exc.errno == errno.ENOSPC:
          ...
  ```

  `strerror` and `filename` deliberately stay `None`: `OSError` rewrites its own `str()` into
  `[Errno n] strerror: 'filename'` the moment `filename` is set, discarding the message — which
  already carries the path and the context. Setting `errno` alone leaves `str()` byte-identical,
  which is what makes this additive rather than a break.

  On Windows the value is the POSIX equivalent, not the native Win32 code, so a comparison
  against `errno.ENOENT` means the same thing on every platform. Where no meaningful number is
  available it stays `None` — compare against a constant, not for truthiness.

### Fixed
- **`typing.get_type_hints()` now works on Python 3.9** for every option shape. The field
  annotations in `xlsxturbo.types` used PEP 604 unions (`str | None`), which 3.9 cannot
  evaluate; `from __future__ import annotations` hid that at import time, so the failure
  surfaced only for a framework that resolves hints — pydantic, FastAPI, attrs, or anything
  building a schema from a type. 35 annotations are now spelled `Union[...]` / `Optional[...]`,
  verified on a real 3.9 interpreter.

  **Python 3.9 support is unchanged.** Dropping it would have solved the same problem while
  also removing users; the PEP 604 syntax was a maintainer convenience.

### Internal
- **`ConvertError` no longer has a blanket `From<String>`.** It mapped every untagged pipeline
  failure to `Config`, and thence to `ConfigurationError` — so a filesystem failure added later
  silently blamed the caller's options, and nothing failed. Two such misclassifications had been
  found and written down before it was removed. Each of the 14 sites now names its category, and
  a new failure site does not compile until it chooses one.

  This goes further than the review proposed, which was an `Internal` fallback variant. A
  fallback that still exists is still a default, and the default was the bug.
- **Property tests over `src/parse/`**, stated as round-trips, idempotences and equivalences
  to the `str` method each parser claims to implement, plus boundary tests for the Excel
  serial-date scale, the 1900 leap-year gap and the 2^53 integer cutoff.

  Two of the properties did not test what they appeared to, and were only found by mutating
  the code they guard: one could reach its failing input roughly once in seventy thousand
  draws, and one never entered the branch it was written for at all. Both generators were
  narrowed until the interesting case is common. Every property and every boundary test has
  since been shown to go red under a deliberate single-edit mutation of the code it covers.
- **`scripts/coverage_report.py`** reports coverage of the Rust core and the Python layer.
  Deliberately no threshold, in CI or out of it: the useful output is *which* branches are
  unexercised.

  It measures both test suites together because neither is honest alone. `cargo test` on its
  own reports 26% and shows every `src/apply/*.rs` file at zero — those paths are covered
  thoroughly, from Python. The Python suite alone misses the parser branches only the Rust
  property tests reach.

### Documentation
- **The type-detection table now covers the three places a value changes kind.** All three
  were already the implemented behaviour and none was written down: any spelling of a
  non-finite number (`nan`, `+Inf`, `Infinity`, …) becomes an empty cell — which is a trap for
  a *text* column that happens to contain one of those words; dates before 1900-03-01 are
  written as text because Excel numbers them a day out; and integers past 2^53 are written as
  text so no digits are lost.

## [0.20.0] - 2026-07-30

### Added
- **`ExportOptions` — the export options as one reusable, typed object.**
  `df_to_xlsx` takes 27 parameters, which is unpleasant for a call with ten options and gives
  you nowhere to keep a set of options you use more than once:

  ```python
  from xlsxturbo import ExportOptions

  REPORT = ExportOptions(freeze_panes=True, autofit=True, header_format={"bold": True})
  xlsxturbo.df_to_xlsx(df, "out.xlsx", **REPORT.as_kwargs())
  ```

  The same bundle serves the multi-sheet path via `as_sheet_options()`, which drops the two
  workbook-level options a per-sheet dict rejects. It is frozen, so it is safe as a shared
  constant; `dataclasses.replace` derives a variant and `merged_with()` layers a sparse
  override without resetting everything else.

  **Every keyword argument still works, unchanged, and nothing is deprecated.** The compiled
  entry points are untouched: the bundle lowers to the keywords they already accept rather
  than wrapping them. A wrapper would have had to either duplicate all 27 parameters or
  collapse them to `**kwargs` — and showing `**kwargs` where an editor currently shows 27
  named, typed parameters would have cost more discoverability than the object adds.

  An option left unset is omitted when lowering; an option set to `None` is passed through.
  The two are different: per-sheet, `table_style=None` means "no table on this sheet",
  overriding a workbook default.

  Only one new public name. The per-feature shapes (`ChartOptions`, `ValidationOptions`, ...)
  remain `TypedDict`s in `xlsxturbo.types` — they already describe themselves to a type
  checker, and grouped `LayoutOptions`/`TableOptions`-style objects were dropped because the
  natural groups are two to five fields and none earned a public class.

### Changed
- **A `data_bar` conditional format and a sparkline on the same worksheet are now refused.**
  rust_xlsxwriter 0.97.0 emits unbalanced `<ext>` elements for that pair, producing a workbook
  Excel reports as damaged. xlsxturbo now raises `ConfigurationError` naming the sheet, the
  column and two workarounds, before writing anything, rather than producing a file you cannot
  open.

  Reproduced against rust_xlsxwriter alone, with no xlsxturbo code in the path, so it cannot be
  fixed here; 0.97.0 is the latest release. Deliberately narrow: only `data_bar` collides, each
  feature alone is unaffected, and every other conditional-format type still works beside
  sparklines. `tests/upstream_defect.rs` asserts the upstream bug is *still present*, so the day
  it is fixed a test fails and the workaround gets removed instead of outliving its reason.

## [0.19.1] - 2026-07-30

Three defects in 0.19.0's new public surface, all found by an independent review of that
release and all verified against the built extension before being fixed.

### Fixed
- **`docs/errors.md`'s central promise is now true.** It states that every failure xlsxturbo
  itself raises is an `XlsxTurboError`. It was not: the option extractors used a bare
  conversion for *nested* keys and values, so PyO3's own `TypeError` escaped the hierarchy
  unclassified. A non-string `column_formats` key, a non-string `formula_columns` value, a
  bad `merged_ranges` tuple element, and the same shape in `comments`, `images`, `cells`,
  `hyperlinks`, `validations`, `rich_text`, `charts`, `sparklines`, `checkboxes`, `textboxes`
  and `column_widths` all escaped. Every one is now a `ConfigurationTypeError`.

  These errors also name the option they came from. A wrong key type inside any nested option
  dict previously reported that a dict key was bad without saying which option's — nearly
  useless on a call passing a dozen options.

  Argument conversion done by the binding *before* xlsxturbo sees a value still raises a plain
  `TypeError`, unchanged and documented. That now includes `row_heights` and `defined_names`,
  which are typed in the signature rather than read by an extractor.
- **`xlsxturbo.types` declares `__all__`.** `from xlsxturbo.types import *` previously also
  bound `PathLike`, `Literal`, `TypedDict` and `Union`. It now binds exactly the 20 documented
  shapes.
- **Required option fields are marked required.** Nine shapes documented a field as "required
  at runtime but TypedDict doesn't enforce this" — `CommentOptions.text`, `ImageOptions.path`,
  `ChartOptions.type`, `SparklineOptions.range`, `CellValueOptions.value`,
  `ValidationOptions.type`, `ConditionalFormat.type`, `CheckboxOptions.checked` and
  `TextboxOptions.text`. A type checker accepted `images={"D1": {}}`, which raises at runtime.
  It can be enforced on Python 3.9, so it now is, and each requirement is tested against the
  error the runtime actually raises. `ChartSeriesOptions` keeps a documented one-of
  requirement across `values_range`/`values`/`data_range` that a `TypedDict` cannot express.

  **This can newly fail type checking on code that was always broken at runtime.** No runtime
  behaviour changed.

## [0.19.0] - 2026-07-30

### Added
- **A typed exception hierarchy, rooted at `xlsxturbo.XlsxTurboError`.** `except XlsxTurboError` now catches everything this library raises and nothing else. Five subclasses say what kind of failure it was: `ConfigurationError` (an option or argument *value*), `ConfigurationTypeError` (wrong *type*), `InputDataError` (not a supported DataFrame), `FileError` (a filesystem read or write), and `WorkbookValidationError` (well-formed configuration that Excel forbids). All six are exported from the package and covered by the type stubs.

  **This changes nothing about which builtin exception any given failure raises.** Every class also inherits the builtin its failures raised in 0.18 and earlier, so existing `except ValueError` / `except TypeError` handlers keep working untouched — verified by the 93 pre-existing assertions on builtin exception types across the suite, all of which pass unedited. `FileError` additionally inherits `OSError`, which is what makes `except OSError` work around a save for the first time; it keeps its `ValueError` base and its `Failed to save workbook to '<path>': ` message prefix, so message-matching code written against 0.18 is unaffected.

  Two mappings follow history rather than taste, and are documented as such: `InputDataError` is a `ValueError` (not a `TypeError`) because that is what an unsupported DataFrame has always raised, and a dtype problem found deep in the write pipeline arrives as `ConfigurationError` rather than `InputDataError`. `errno`, `strerror` and `filename` on a `FileError` are always `None`; the path is in the message.
- **`xlsxturbo.types` — the option shapes as a real runtime module.** The 20 option `TypedDict`s and `Literal` aliases (`HeaderFormat`, `ChartOptions`, `ValidationType`, `SheetOptions`, ...) moved out of the type stub into `python/xlsxturbo/types.py`, so annotating an option dict no longer needs a `TYPE_CHECKING` guard and an import from the compiled submodule:

  ```python
  from xlsxturbo.types import HeaderFormat

  header: HeaderFormat = {"bold": True, "bg_color": "#DDDDDD"}
  ```

  `xlsxturbo.pyi` imports the shapes rather than declaring its own copies, so there is one home per fact; it keeps the four function signatures, the exception classes and `__version__`, and drops from 650 lines to 377. The old `from xlsxturbo.xlsxturbo import HeaderFormat` still type-checks. `types.py` imports nothing beyond the standard library, so it is safe to import before the extension is built, and `tests/test_types_module.py` fails if the stub's re-export list and the runtime module drift apart.

  One caveat, verified on a real 3.9 interpreter: field annotations are strings (the module uses `from __future__ import annotations`, which is what lets a field be written `bool | str` and still import on 3.9), so `typing.get_type_hints()` on these classes fails on 3.9 and works from 3.10. Static type checking is unaffected on every version.
- **An unsupported DataFrame is now rejected before any file is created.** `df_to_xlsx` and `dfs_to_xlsx` check the frame type at the boundary rather than partway through the write, so a bad input no longer leaves a partially considered output path. Same message and same predicate as before, and DataFrame subclasses are still accepted.
- **`docs/capability-matrix.md`** — which option applies in which mode (`df_to_xlsx`, `dfs_to_xlsx`, `csv_to_xlsx`, per-sheet overrides, `constant_memory`). Generated from the Rust sources by `scripts/gen_capability_matrix.py`, so it cannot drift from the implementation; `tests/test_capability_matrix.py` fails if the committed page is stale. It makes explicit something the README never stated: `csv_to_xlsx` accepts three options and none of the formatting or feature ones — it is a straight-through fast path, not a reduced-feature DataFrame path.
- `CONTRIBUTING.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`, three issue forms and a pull-request template. `SECURITY.md` documents a scope that follows from the library only ever *writing* xlsx: the vulnerability classes affecting spreadsheet readers largely do not apply.
- `docs/roadmap-1.0.md` — the planned work between here and a 1.0 release.
- GitHub Releases now exist for every tag. The release workflow creates them going forward, with wheels and SBOMs attached; the 37 pre-existing tags were backfilled from their changelog sections.

### Changed
- `rust_xlsxwriter` 0.96 → 0.97.
- CI gained `cargo audit`, `pip-audit`, CodeQL (Python **and** Rust), dependency review, and a check that each SHA-pinned action's version comment still names that SHA. All third-party actions are pinned to commit SHAs.
- The release workflow attests build provenance for the published wheels and publishes CycloneDX SBOMs.
- Dependabot now tracks Python dev/build dependencies in addition to Cargo and Actions.
- `python-lint` installs into a project-local `.venv` and runs ruff, bandit and pyright from it, matching the documented local commands. Previously it installed into the runner's system interpreter while pyright's config pointed at `.venv`, so the gate analysed a different environment in CI than locally — and could pass one while failing the other.
- `scripts/` is covered by the ruff and pyright gates; it was previously unchecked.

### Documentation
- **The documentation is now a site.** `README.md` goes from 1289 lines to a landing page; the reference content moved into 16 pages under `docs/`, published to GitHub Pages by a new `docs` workflow. The split was done mechanically and each moved section verified byte-identical in exactly one destination, so no example or caveat was lost or reworded in transit. New pages: `api-reference.md` and `errors.md`.
- `errors.md` documents the exception hierarchy, which builtin each class keeps for backwards compatibility, and the two places the classification is deliberately coarse. Writing this page is what surfaced the I/O-as-`ValueError` wart in the first place; the hierarchy above is the fix.
- **Corrected a false claim on the front page: the command-line tool is not in the PyPI wheel.** The README stated "Available as both Python library and CLI tool" and documented `xlsxturbo in.csv out.xlsx` as if `pip install xlsxturbo` provided it. It never has — maturin packages only the extension module, confirmed by inspecting the published 0.18.0 artifact. The CLI is a Cargo `[[bin]]` target built by `cargo build --release`; the docs now say so, and `csv_to_xlsx` is the supported route from Python.
- `CONTRIBUTING.md` gained sections on working on the docs site and on building a distributable wheel, the latter absorbing the README's former "Building from Source".

## [0.18.0] - 2026-07-25

### Added
- Subclasses of `pandas.DataFrame`/`polars.DataFrame` (e.g. geopandas' `GeoDataFrame`, or a locally defined `class MyFrame(pd.DataFrame)`) are accepted. Detection is an `isinstance` check against the real classes rather than a match on the type's `__module__` prefix, which rejected every subclass — one defined in a script reports `__main__`. Unrelated duck-typed objects are still refused, since identification remains by type rather than by probing for attributes.

### Fixed
- **Saving is now atomic: a failed write can no longer destroy the file already at the output path.** `Workbook::save` truncates the destination *before* it serializes and validates, so any failure partway through — a chart range naming a sheet that does not exist, a full disk, a dropped network share — left a 0-byte file where the previous export had been, destroying it as a side effect of an error that was otherwise reported cleanly. All four save paths (`df_to_xlsx`, `dfs_to_xlsx`, and both CSV conversions) now build the workbook in a temporary file in the destination's own directory and rename it into place only on success, so the destination is always either the old file or the complete new one. When an existing file is replaced its permissions are preserved. Because the staging file is created next to the destination, that directory must exist and be writable.
- A save to a nonexistent directory reports `directory '<dir>' does not exist` instead of a raw OS error.

### Changed
These two tighten validation and can raise where a previous version silently produced output. Both reject values that never had any effect on the resulting file.

- Chart `style` is validated against Excel's documented 1-48 range. Values a `u8` silently accepted (`0`, `200`) reached rust_xlsxwriter, which discarded them and reported it only on stderr where Python cannot see it, so the chart came out with the default style and no error; values above 255 raised a "range 0-255" message that had nothing to do with Excel's limit. Matches the guard sparkline `style` has had since 0.16.1.
- `rich_text` segment formats reject cell-level keys (`border*`, `align_horizontal`, `align_vertical`, `wrap_text`) instead of accepting and silently ignoring them. A segment is an inline run inside one cell, so only font-level properties reach the XML — the accepted-but-inert keys contradicted both the `RichTextFormat` type stub and the parser's own docstring. Format the cell itself via `column_formats` or `cells` instead.

### Documentation
- README documents that writes are atomic and what that implies for the output directory; that `formula_columns` combined with `table_style` places the formula columns outside the Excel table (no banded fill, no autofilter, not covered by `column_widths`/`autofit`); and that duration columns (`pandas.Timedelta`/`numpy.timedelta64`) are written as text, with the recommended conversion.
- The strict-unknown-key note now lists `rich_text` and spells out its font-only key set.
- `AGENTS.md` gives the lint/type/security gate commands for both Windows and macOS/Linux venv layouts; the previous list was Windows-only.

### Internal
- New `TestAtomicSave` test class covering the save-failure contract in both the single- and multi-sheet paths: an existing file survives a failed save byte-for-byte, no file is created where none existed, a successful save still replaces, no staging files are left behind, and permissions are preserved.
- `OptionMap::u8` removed — chart `style` was its only caller and now range-checks via `i64`.
- The format-dict parser's `include_column_options` bool became a three-way `FormatScope` (`Font`/`Cell`/`Column`), making the rich-text key set a scope rather than a special case.

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

[0.18.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.18.0
[0.17.2]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.17.2
[0.17.1]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.17.1
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
