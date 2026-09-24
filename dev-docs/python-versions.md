# New CPython releases, free threading and output determinism

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

### A new CPython needs no code change — and the `abi3` promise has one hole

Checked against 3.15.0rc1 on 2026-08-26, Windows x86_64. **The published wheel works on a
Python released after it was built, which is the whole point of `abi3-py310`**: pip resolves
`cp310-abi3-win_amd64` on 3.15rc1, the extension imports, and a pass over `df_to_xlsx`
(header format, column widths, table style, freeze panes, column formats, a `data_bar`
conditional format, formula columns, comments, a chart), `dfs_to_xlsx`, `csv_to_xlsx`,
`ExportOptions` and the whole exception hierarchy all behave. The sdist also builds from
source under 3.15rc1 with pyo3 0.29 — no `PYO3_USE_ABI3_FORWARD_COMPATIBILITY` needed.

⚠ **The `abi3` wheel does not cover free-threaded builds, and `docs/stability.md` promised
that it covered "every supported version".** On `python3.14t` — a version that page lists as
supported — `pip install xlsxturbo` finds **no usable wheel** and falls back to the sdist,
which needs a Rust toolchain. **The control is what makes this a fact about free threading
rather than about the machine**: the identical command on the ordinary 3.14.7 installs
`cp310-abi3-win_amd64` in a second. Both were measured, on 3.14.7t and 3.15.0rc1t; the sdist
builds and runs correctly on both, in about a minute. The page now says so. Whether
free-threaded builds are *supported* is a promise nobody has made yet — do not add one to
that page without asking.

**The code is ready for free threading; the ecosystem is not, and the interesting number is
on the ordinary build.** Measured on 3.14.7t, 32 cores:

- **pyo3 declares the module free-threading-safe by default, and nobody here opted in.**
  `#[pymodule]` with no `gil_used` argument compiles to `Py_MOD_GIL_NOT_USED` —
  `pyo3-macros-backend/src/module.rs` reads `options.gil_used.is_some_and(...)`, so *absent*
  means `false` means "does not need the GIL". Confirmed at runtime: `sys._is_gil_enabled()`
  stays `False` after `import xlsxturbo`. Know this before assuming a `t` build silently
  re-enables the GIL — it does not, and the assertion is the macro's, not ours.
- **It survives the assertion.** 502 passed, 2 skipped on 3.14t (the 4 collection errors are
  polars, below). Eight threads exporting one *shared* DataFrame produce 40 files that all
  verify. `PyOnceLock<ErrorTypes>` is the only `static` in `src/` — no `static mut`, no
  `RefCell`, no `unsafe` — and it is the type pyo3 provides for exactly this. ⚠ **Nothing can
  race its initialisation, so do not write a test claiming to**: `errors::register` calls
  `get_or_init` during module import, so it is filled before any Python code can run. A test
  doing this covers concurrent *reads* and that `except` still matches by identity, which is
  worth having under a docstring that says so.
- ⚠ **polars ships no free-threaded wheels, for the same `abi3` reason as ours.** pandas 3.0.5
  and numpy 2.5.2 do. So a `cp3XXt` wheel would put users on an interpreter where one of the
  two supported frame libraries cannot be installed from a wheel at all. That, not any defect
  here, is the argument against shipping one today.
- **`df_to_xlsx` holds the GIL for the whole call.** The single `py.detach` in the crate is
  `csv_to_xlsx` (`src/lib.rs:257`). Threaded `df_to_xlsx` measures **1.03x on an ordinary
  build and 6.0x free-threaded** at 8 threads; `csv_to_xlsx` measures **~5.9x on both**. That
  control is what makes this a statement about the detach and not about the interpreter.

**That measurement is why `df_to_xlsx` and `dfs_to_xlsx` now detach for the save**, at
`convert.rs` and `lib.rs` (the third `save_workbook` call, in `convert_csv`, already runs
inside `csv_to_xlsx`'s detach — do not wrap it again). Threaded exports on an **ordinary**
build went from 1.06x to **2.36x** at 4 and 8 threads, single-thread cost unchanged,
reproduced over two interleaved A/B passes. The plateau is Amdahl on the extraction half,
which reads Python objects and cannot be detached; it puts the save at ~58% of the call.

Two things about testing it, because a GIL release is only observable through timing:

- **The correctness tests do not cover the change and are not meant to.**
  `TestConcurrentWrites` stays green with both `py.detach` wrappers deleted — verified. It
  covers the *hazard* the detach creates (Rust running while the interpreter is unlocked),
  which is the half that could corrupt a file. Do not read its green as cover for the detach.
- **`TestSaveReleasesTheGil` is the guard, and it was mutation-checked per call site.** The
  threshold is 0.80x with best-of-3 legs and a 4-core skip, chosen to sit in the middle of a
  measured 0.43-vs-0.98 gap rather than close to either side. If it ever flakes, widen the
  threshold — do not delete it.
- ⚠ **A wall-clock assertion is invalid under `-Cinstrument-coverage`, and this repo runs
  the whole pytest suite that way.** On an instrumented build four threads are *slower* than
  one — 1.693s against 1.347s on a 32-core machine where the ordinary build gives 0.43x — so
  the comparison measures the instrumentation at any threshold. `TestSaveReleasesTheGil`
  therefore skips when `XLSXTURBO_COVERAGE` is set, which `scripts/coverage_report.py` exports
  for every pytest pass it drives. **The local suite and every CI test leg passed; only the
  Coverage job caught it**, which is the job whose failures are easy to wave through as
  informational. Its `AGENTS.md` line already said what to do: a Coverage failure means the
  measurement broke. Read it.
- ⚠ **`scripts/coverage_report.py` runs pytest TWICE, and the first fix keyed on a marker only
  one of them sets.** Step 3 is the instrumented pass with `LLVM_PROFILE_FILE`; step 5 measures
  the pure-Python layer under coverage.py — with no env of its own, and against the instrumented
  extension step 3 left in the venv. Keying the skip on `LLVM_PROFILE_FILE` therefore fixed
  half of it and left step 5 making a wall-clock assertion under *two* sources of slowdown.
  Hence the dedicated `COVERAGE_MARKER` constant, set on both. **The tell was that both passes
  reported the same counts as an ordinary run**, and it was nearly missed because the log had
  been read through `tail -20`, which showed step 5's summary and hid step 3's. When one script
  runs a suite more than once, a single summary line is not the job's result.
- ⚠ **Two detaches need two mutations, and the first draft covered only one of them.** The
  timing test originally exercised `df_to_xlsx` alone, so deleting the `lib.rs` wrapper —
  half the change — left the **entire suite green at 739 tests**. Found by an independent
  review asking which single-site mutation survives, not by the both-sites mutation that had
  already been run and looked convincing. It is now parametrized over both entry points, and
  each single-site mutation reddens exactly its own case (`dfs_to_xlsx` 1.00x, `df_to_xlsx`
  0.94x) and leaves the other green. This is the repo's own *"a check bound to one caller
  covers only that caller"* trap: when one change touches N call sites, mutate each alone.

Two more things that review established, both worth keeping:

- **The output is unchanged, and that is measured rather than argued.** The same workbook
  built by the stock 1.3.0 wheel and by the patched build has 14 zip members, 13 of them
  byte-identical; the only difference is `docProps/core.xml`, the creation-timestamp part
  this repo already documents as the one non-reproducible member.
- **`save_workbook` resolves a relative `output_path` against the process cwd twice** — once
  for `NamedTempFile::new_in(dir)`, once for `tmp.persist(dest)` — so releasing the GIL lets
  another thread `os.chdir` between them. Real by reading; **it does not reproduce as a
  behaviour difference.** 300 exports against 67,656 concurrent chdir flips gave no error, no
  stray temp file, and the same "lands in whichever directory is current" outcome as the
  stock build, because a same-volume rename across directories simply succeeds. That probe
  also measured the change working from the other side: the flipping thread got 48 turns
  during 60 stock exports and 13,080 during 60 patched ones.

Two things that decide the work when a new CPython ships:

- **Declaring support is two documentation edits and no code.** Add the trove classifier and
  the `docs/stability.md` row; `tests/test_stability_policy.py` compares them in both
  directions and stays red until they agree. A CI leg is *not* required — the "Run in CI"
  column may say `no`, which is already the page's own argument for 3.11 and 3.13. Do not add
  a row for a `t` build: the parser reads the first table under that heading and compares
  cell 0 against the classifiers, so a `3.14t` row fails the suite. Prose after the table is
  safe, and that is where the free-threading note went.
- **A CI leg is gated on pandas, not on xlsxturbo.** On 3.15rc1 neither pandas nor pyarrow had
  a `cp315` wheel, so 17 of 28 test modules could not even be collected and ~90% of the suite
  was unmeasurable. polars and openpyxl did ship 3.15 wheels, which is what made the
  functional pass and the 24 pandas-free repo tests possible at all. Expect this gap every
  year: the library is ready for a new CPython months before its test dependencies are.

### Output is deterministic except for one part, and the obvious measurement says otherwise

Two exports of the same frame differ, because `docProps/core.xml` records the creation time.
Every other member of the archive is byte-identical across runs.

The trap is that writing both files and hashing them reports **identical** — that timestamp
has one-second resolution, so a loop with no delay measures nothing. It was one step from
being published as "output is byte-reproducible".
`TestGeneratedFileDeterminism` waits 1.1 s deliberately, and mutating that wait to zero turns
both of its tests red, which is what proves neither passes by accident. If a reproducible-build
option is ever added, it belongs here as an opt-in — and it is additive, so it needs no major.
