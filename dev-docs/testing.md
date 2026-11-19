# Test environment, coverage and property tests

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

  **A declared version range is only supported at the end CI installs.** pip resolves to the
  newest allowed, so `pandas>=2.3.3,<4` means pandas 3 on every leg and pandas 2 never. Since
  1.1.0 a fourth `python-test` leg installs `requirements-test-pandas2.txt`, which **layers**
  the main file (`-r requirements-test.txt`) and overrides only the pandas ceiling — never
  restates the list, which would be a fifth copy of it. `tests/test_ci_config.py` fails if
  that file stops layering, gains a package, stops being referenced by a workflow, or outlives
  its reason (the pandas range narrowing back to one major). It was scoped to pandas while
  that was the only dependency whose range spanned two majors, with the note that a second
  one would need its own leg and nothing would notice on its own.

  polars became the second on 2026-10-13 (`polars>=1.44.2,<3`, leg file
  `requirements-test-polars1.txt`), and the guard now notices: `OLDER_MAJOR_LEGS` in
  `tests/test_ci_config.py` maps each two-major dependency to its layered file, and
  `test_every_two_major_range_has_a_leg` compares that table with the bounded ranges in
  `requirements-test.txt` in both directions. It reads only requirements that have a
  ceiling, so `numpy` and `pyarrow`, which have none, are outside it.

  How the gap was found is the transferable part: the local `.venv` had pandas 3 while
  `requirements-test.txt` said `<3`, so **local runs and CI had not been testing the same
  library for weeks** and everything was green on both. The usual assumption here — that the
  local environment is a strict superset of CI — was simply false, in the direction nothing
  checks.

  That guard exists because the same bug landed three times in two days: `tests/test_docs_site.py` imported `yaml` (declared in `dev`, broke three CI jobs); fixing those three left a **fourth** copy of the list in `release.yml`'s smoke-test job, which failed the v0.19.0 release after every wheel had already built; and the guard's first run found `numpy` imported by `tests/test_core.py` and never declared anywhere — working only because pandas pulls it in. A comment saying "remember the other copies" was the fix after the first, and it did not survive a day.

- **The release workflow must smoke-test every published wheel on a runner of its own architecture, not a representative subset.** All five legs install the wheel and run the full suite before the publish job. Two of them existed only from 1.1.0 onward: Linux `aarch64` and macOS `x86_64` are cross-compiled and had no hosted runner when the pipeline was written, so they shipped untested for every release before that. The runners are `ubuntu-24.04-arm` and `macos-15-intel` — note `macos-13`, the old Intel image, has been retired, so check the current label before assuming. Each leg asserts the platform tag of the wheel it downloaded; without that a mistyped `wheel-artifact` installs some other wheel twice and two green legs read as coverage. `tests/test_stability_policy.py` compares the build matrix and the smoke-test matrix in both directions and ties them to the table in `docs/stability.md`, so a new build target cannot be published untested and a leg naming a non-existent artifact fails locally rather than after every wheel has built.

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

The script installs an instrumented extension into the venv, and a later test or benchmark
run would measure it without any check noticing, so it ends, on failure as well, by
rebuilding with plain `maturin develop --release` and prints `[OK]` or `[WARN]` saying which
build is installed; `tests/test_coverage_script.py` holds it to that.

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

## Replacing a dependency: run both side by side before the old one goes

`chrono` became unmaintained (RUSTSEC-2026-0543, `cargo audit --deny warnings` red on `main`)
and was replaced on 2026-11-17 by `src/parse/civil.rs`. A date crate decides which CSV cells
become dates, and that is output the stability policy covers, so "the tests still pass" was
not the bar. The bar was: no input gives a different result. What made that checkable:

- **A second date crate was not the replacement.** xlsxturbo needs three things: is this
  year-month-day a date, which day number is it, and does a text match one of ten fixed
  patterns. jiff's years end at 9999 where chrono's end at 262142, and its parser has none of
  chrono's leniencies, so it would have supplied two functions and changed results. The
  leniencies are not documented anywhere as a list; they were read from chrono's
  `format/parse.rs` and `format/scan.rs`: white space of any kind is skipped before every
  number, a number has one digit or more, `%Y` takes four digits or any count after a sign,
  `%.f` matches nothing without a `.`, uses nine digits and drops the rest, second 60 is a
  leap second, and a space in the pattern matches no space at all (`2024-01-1510:30:00` is a
  datetime).
- **The comparison ran with the old crate still in the build**, as a test module that called
  both and compared `f64::to_bits`: every year-month-day from year -262146 to 262145 with
  months 0 to 13 and days 0 to 32 (191,491,529 valid dates), 899,496 combinations of hour,
  minute, second and nanosecond, 43 million timestamps including both ends of the range, and
  80 million generated texts under all ten patterns. No difference.
- **The first text generator was nearly useless and its own count showed it.** It built texts
  from random parts and 128,214 of 40 million matched a pattern, so the accepting side of
  the parser was hardly run. The second generator starts from a pattern, fills it with valid
  numbers and changes about one part in eight; 18 million of 80 million matched. The test
  asserts both counts (more than 10 million accepted, more than 10 million refused), so a
  generator that only produces refusals fails.
- **14 mutations of the new code were declared with the test each should fail, and all 14
  failed it**, among them a third digit for the month, no 400-year rule, no space skip, a
  lost leap second, one more year in the range and a truncating division for timestamps
  before 1970.
- **The comparison cannot stay, because the old crate has to leave the lockfile for the audit
  to pass.** So each rule got a named test in `civil.rs` first, and the same mutations plus
  five more ran against those: 18 of 19 failed, and the survivor was a range check no caller
  can reach, which then got its own test. Do the mutation run twice, once per oracle; the
  second run is the one that says what protects the code from now on.

A date pattern added to `DATE_PATTERNS_*` or `DATETIME_PATTERNS` with a field other than
`%Y %m %d %H %M %S %.f` never matches; `every_pattern_in_use_is_one_this_module_reads` fails
for it.

## A stale build fails the suite

**`tests/test_build_is_current.py` fails when the built extension is older than any tracked Rust source, `Cargo.toml`, `Cargo.lock` or `pyproject.toml`.** A stale `.so` does not error, it measures the wrong code with full confidence: on 2026-09-02 a whole review ran against 1.3.0 code believing it was HEAD, and a later gate run blamed a timing test for what was the build. CI builds immediately before pytest, so the test only ever fires locally, where `maturin develop --release` is the answer. A `git checkout` that touches a `.rs` file trips it on purpose. `pyproject.toml` is compared against the installed `.dist-info/RECORD` rather than the extension file, because cargo never reads it: after a `pyproject.toml`-only change `maturin develop --release` compiles nothing and leaves the extension's mtime alone.

A comment-only edit to `Cargo.toml` trips it as well, and `maturin develop --release` did not clear it (observed once, 2026-11-10): cargo judged the crate fresh and left the extension's mtime alone. `cargo clean -p xlsxturbo --release` and a rebuild cleared it.

## Reading a workbook back is not opening it in Excel

**The suite checks output with openpyxl and by reading the XML. Neither parses a value the way Excel does, so a workbook Excel refuses can pass every test.** From 0.10.0 until 2026-11-10 a `decimal` validation with one bound wrote `f64::MAX` (1.797e308) as the missing one. Excel's largest number is 9.99999999999999E+307, and it refused to open the file; the same rule with an explicit `1e300` opened, so the value was the cause, not the 309-digit length. Every decimal test passed both bounds, and openpyxl read the one-sided file without complaint.

So a change that writes a number, a formula or a name Excel has to interpret (validations, conditional formats, defined names, table names) is checked in Excel once, by hand, before it ships. On Windows that takes seconds through COM: a new `Excel.Application` with `Visible` and `DisplayAlerts` off, `Workbooks.Open(path, 0, $true)`, then read the object back (`Range.Validation.Type`, `.Operator`, `.Formula1`). With alerts off, a file Excel would offer to repair fails to open with `Unable to get the Open property of the Workbooks class`. Run a file known to open in the same session as the control, and close only the instance the probe created. `scripts/excel_open_check.ps1 -Control <good.xlsx> <files>` does the open check that way and refuses to give a verdict when the control fails; its first ad-hoc predecessor passed `[Type]::Missing` for the optional arguments, which pwsh COM rejects, so every file failed, the control included (2026-11-15). The COM error alone does not prove a repair prompt, because every failed open raises one: a 0-byte file gives "the file format or file extension is not valid" through the same path. So the script opens each refused file a second time in a new visible Excel process with alerts on and prints the text of the dialog a user would see (`We found a problem with some content in ...` is the repair prompt); that step needs the logged-on desktop session and `-NoDialog` skips it (2026-11-17, after 16 refusals in a pivot table probe were each confirmed by their dialog). The script tells "could not check" apart from a verdict, because each such case was once printed as one: a `-Path` item that matches nothing (it was dropped, and the run said `1 of 1 opened` for two paths given), Excel missing or failing through COM (it exited 1, the code for a refusal), and Excel rejecting the call while it is busy, as it is while a pivot table refreshes on load (`RPC_E_CALL_REJECTED` was printed as `REFUSED`; the same workbook opened in the next run). Now every `-Path` item is resolved on its own, literally when the file exists so a name with `[` or `]` is found; a rejected call is retried and one still rejected is reported as `BUSY`; and each of those cases exits 2. Exit 1 means at least one file was refused and nothing was left unchecked. A stand-in value for a missing bound is the shape to distrust: write the one-sided rule, not a range up to the type's limit.
