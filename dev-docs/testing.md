# Test environment, coverage and property tests

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

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
