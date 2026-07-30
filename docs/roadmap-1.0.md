# Roadmap to 1.0

**Complete. All six phases landed; 1.0.0 shipped 2026-07-30.** This file is kept as the
record of what was decided and why — the D-numbered decisions and the per-phase "what this
found" sections are the parts worth reading later. It is no longer a plan, and nothing here
is waiting to be done: what survived unbuilt is in [Deferred, with
reasons](#deferred-with-reasons), and the promises 1.0 now carries are in
[`docs/stability.md`](stability.md).

Working plan for the 0.19 -> 1.0 cycle. `ROADMAP.md` tracks *feature* gaps; this file
tracks the engineering work needed to turn a feature-complete library into a stable,
adoptable public package.

Written 2026-07-30 following an independent external project review. The review's findings
were fact-checked against the tree at `v0.18.0`; the baseline it produced is recorded under
[Baseline](#baseline-verified-2026-07-30) so a later session can tell what has since moved.

**How the phases were run:** ordered and mostly sequential — Phase 3 depended on Phase 2's
public surface, and Phase 5 followed both. Phases 0 and 1 were independent of everything.
Each phase ended at the gate set in `AGENTS.md`; phases 2 and 3 additionally got a deep diff
review before the next phase built on them.

---

## Decisions already taken

Recorded here because they are not derivable from the code, and re-deciding them costs more
than reading them.

### D1 — Option types live in a runtime module, not in the stub

`python/xlsxturbo/types.py` becomes the authoritative home for the option `TypedDict`s and
`Literal` aliases. `xlsxturbo.pyi` keeps only the compiled functions' signatures and imports
the shapes from `xlsxturbo.types`.

This splits responsibility so each fact has exactly one home:

| File | Owns | Kind |
|------|------|------|
| `python/xlsxturbo/types.py` | option shapes (`HeaderFormat`, `ChartOptions`, ...) | runtime module, authoritative |
| `python/xlsxturbo/xlsxturbo.pyi` | the four compiled function signatures | stub, authoritative |
| `python/xlsxturbo/__init__.pyi` | thin re-export of the runtime surface | stub, mirrors `__init__.py` |

Rejected alternative: keep the stub authoritative and generate `types.py` from it. It needs
build-time codegen, a committed generated file and a freshness test — and the generator would
have to rewrite syntax, because:

**The constraint that decides this.** `xlsxturbo.pyi` declares
`PathArg = str | PathLike[str]` at module level. A `.pyi` is never executed, so PEP 604
unions are free there. `requires-python` is `>=3.9`, where evaluating that expression raises
`TypeError` — so a runtime module cannot copy the stub verbatim, and a codegen doing the
`|` -> `Union[...]` transform is fragile for no gain. `types.py` therefore uses
`Union[...]`/`Optional[...]` for module-level aliases, with `from __future__ import
annotations` so `TypedDict` field annotations stay unevaluated.

No new runtime dependency: the stub imports only `Literal` and `TypedDict`, both available
from `typing` on 3.9. `os.PathLike[str]` is subscriptable at runtime from 3.9. Do **not**
reach for `typing_extensions` — the package has zero runtime dependencies and that is worth
keeping.

Consequence: touchpoint 6 of the feature checklist in `AGENTS.md` changes meaning. Update
`AGENTS.md` in the same commit that lands `types.py`.

### D2 — New exception classes multiply inherit from the current builtins

`ConfigurationError(XlsxTurboError, ValueError)` and friends, so that every existing
`except ValueError` / `except TypeError` keeps working. Without this the hierarchy is a
breaking change and cannot ship in a minor release.

**The principle held; the class list in this entry did not — see [D6](#d6--the-hierarchy-and-boundary-as-actually-built).**

### D3 — Classify errors at the boundary, not by refactoring internals

Internal functions return `Result<_, String>`; the conversion to Python exceptions happens at
roughly a dozen sites in `src/lib.rs`. Classify there, by call site. Do **not** introduce an
internal error enum in this cycle — it touches all of `apply/`, `parse/` and `convert.rs` for
the last fraction of precision, and the public contract does not depend on it.

**"Do not introduce an internal error enum" held. "Roughly a dozen sites in `src/lib.rs`" was
wrong — see [D6](#d6--the-hierarchy-and-boundary-as-actually-built).**

### D4 — Options objects lower to kwargs in Python

The dataclasses convert to the existing keyword arguments in a thin Python wrapper before
crossing into the extension. `RawOptions` / `extract_options` and the `define_options!`
machinery stay untouched. This is what bounds the cost; a second Rust extraction path would
not.

### D5 — `docs/strategic-recommendations-plan.md` is untracked

It was the only tracked file under `docs/` — an internal planning memo was what a visitor to
the public repo found there. Now gitignored alongside `docs/reviews/`. `docs/` from here on
holds user-facing documentation only, which is what Phase 1 fills it with.

### D6 — The hierarchy and boundary as actually built

D2 and D3 were written before anyone counted the raise sites. Counting them changed four
things. Recorded here rather than by editing D2/D3, because the reasoning is the useful part.

**The boundary is two files, not one.** `src/lib.rs` has 12 conversion sites; **`src/extract.rs`
has 38 more**, raising `PyTypeError`/`PyValueError` directly. D3's "roughly a dozen sites in
`src/lib.rs`" undercounted by a factor of four. Everything below `extract.rs` — `apply/`,
`parse/`, `convert.rs`, `write.rs`, `workbook.rs` — really is uniformly `Result<_, String>`,
so that half of D3 was right and is what makes the rest tractable.

**`UnsupportedFeatureError` is dropped: nothing raises it.** D2 gave it "e.g. constant_memory
conflicts" as its purpose, but a `constant_memory` conflict is a `RuntimeWarning` and the call
succeeds — by design, and `CONSTANT_MEMORY_SAFE_OPTIONS` plus a guard test exist to keep that
deliberate. An exception class no site raises is dead public API that can never be removed.

**`WorkbookWriteError(XlsxTurboError, OSError)` becomes
`FileError(XlsxTurboError, OSError, ValueError)`.** Two reasons, and the first is a
compatibility bug in D2's own terms:

- Without a `ValueError` base it is a **breaking change**. Save failures are `ValueError`
  today (`docs/errors.md` documents the exact prefix), so `except ValueError` around a save is
  code that exists in the wild, and `(XlsxTurboError, OSError)` would stop catching it — the
  precise outcome D2 was written to prevent. Verified legal: `(XlsxTurboError, OSError,
  ValueError)` has no instance-layout conflict, gives a clean single-arg `str()`, and leaves
  `errno`/`strerror`/`filename` as `None`.
- `WorkbookWriteError` is the wrong *name*: `csv_to_xlsx` also fails on **reading** its input
  (`Failed to open input file: ...`), which is filesystem I/O but not a workbook write.

**`ConfigurationTypeError` is added, and it is not a refinement — it is required.** A single
`ConfigurationError(XlsxTurboError, ValueError)` cannot cover `extract.rs`, because ~20 of its
sites raise `TypeError` ("expected dict for 'x', got int") and ~18 raise `ValueError`. One
`ValueError`-based class over both breaks `except TypeError`. The alternative considered and
rejected was `ConfigurationError(XlsxTurboError, ValueError, TypeError)`: it is layout-legal,
but it makes every bad option *value* also a `TypeError`, and the sibling pair costs one class
to keep the mapping 1:1 with today's behaviour. **The whole hierarchy is therefore additive:
no call that raises `ValueError` today raises `TypeError` tomorrow, or vice versa** — which is
a property a test can pin, and `tests/test_errors.py` does.

**`InputDataError` is a `ValueError`, not a `TypeError` — and the existing suite is what
proved it.** The list above was implemented with `InputDataError(XlsxTurboError, TypeError)`,
on the reasoning that "you passed a list, not a DataFrame" is a type problem. It is, but the
builtin is not a matter of taste: `tests/test_core.py::TestDataFrameSubclasses::
test_unrelated_object_still_rejected` asserts `pytest.raises(ValueError, match="Unsupported
DataFrame type")`, and it went red. Frame detection has always raised `ValueError`, because it
reaches the boundary through the write pipeline. So the same bug the `FileError` naming
caught in D2 was reproduced one paragraph later in D6, by the author of that paragraph —
which is the useful part of the record. Two habits follow:

- **Derive the builtin base from a grep of the current behaviour, never from what the failure
  ought to be.** Taste is exactly the input that produces a breaking change here.
- **Run the full suite before believing a compatibility claim.** 93 existing
  `pytest.raises(ValueError|TypeError)` assertions across the suite are a behaviour record
  for pre-0.19; they are load-bearing for this phase and cost nothing to consult. The final
  state passes all 381 pre-existing tests **with zero test edits**, which is the actual proof
  that the hierarchy is additive. Any future change here that needs a test edited to stay
  green is a breaking change, whatever the changelog says.

As built:

```
XlsxTurboError(Exception)
├── ConfigurationError(XlsxTurboError, ValueError)        # an option or argument VALUE is invalid
│   └── WorkbookValidationError(ConfigurationError)       # valid config, but Excel forbids it
├── ConfigurationTypeError(XlsxTurboError, TypeError)     # an option or argument has the wrong TYPE
├── InputDataError(XlsxTurboError, ValueError)            # the object passed is not a supported frame
└── FileError(XlsxTurboError, OSError, ValueError)        # filesystem read or write failure
```

**The one refactor D3 did not anticipate: `save_workbook` is not at the boundary.** It is
called from inside the pipeline (`convert.rs:94`, `:169`, `:767`) and its error arrives at the
*same* `map_err` as an option-validation failure, so classification by call site cannot tell
them apart — and `FileError` is the most valuable class in the hierarchy, so leaving it
imprecise defeats the point. Fixed with a two-variant tag at that one seam:

```rust
pub(crate) enum ConvertError { Config(String), File(String) }
impl From<String> for ConvertError { /* -> Config */ }
```

`From<String>` means every existing `?` on a `Result<_, String>` keeps compiling unchanged and
lands in `Config`; only the three `save_workbook` calls and the one `File::open` are tagged
`File` explicitly. This is **not** the internal error enum D3 forbids — it does not reach
`apply/`, `parse/` or `write.rs`, and it has two variants rather than one per failure mode.

**A second boundary call was needed, for the same reason.** Frame detection also lives inside
the pipeline (`types.rs:460`, reached from `write_sheet_data`), so `InputDataError` was
*unraisable* through `df_to_xlsx` on the first build — `df_to_xlsx([1, 2, 3], path)` raised
`ConfigurationError`, verified by running it. Fixed by calling the same predicate,
`types::is_polars_dataframe`, at the top of both entry points via
`require_supported_dataframe`. Calling the existing function rather than re-implementing the
check is what stops the two sites drifting; the side benefit is that an unsupported input now
fails before any file is created.

That leaves a check whose inner twin already exists, which normally argues for deleting one of
them. Not here: the inner call is the real implementation and the outer one exists to
*classify*, which is information the inner site cannot produce. The mutation that deletes the
boundary call is in the harness and turns
`test_class_is_raised_by_a_real_call[InputDataError]` red, so the outer call is demonstrably
load-bearing rather than defensive decoration.

Accepted imprecision — **two places, and this list is meant to be complete**, so that nobody
mistakes either for an oversight and nobody assumes the classification is finer than it is:

1. A **dtype problem** raised deep inside `write_sheet_data` arrives as `ConfigurationError`,
   not `InputDataError`. `InputDataError` covers only frame detection, which is now at the
   boundary.
2. A **mid-write CSV failure** (`convert.rs`, `"Write error at ({}, {}): {}"`) is untagged, so
   `From<String>` puts it in `ConvertError::Config` and it surfaces as `ConfigurationError` —
   despite having nothing to do with configuration. Same root cause as the trap noted on the
   type: untagged is Config.

Chasing either is exactly the "last fraction of precision" D3 declined. The reason to write
them both down is that a partial list of known imprecisions reads as a complete one, and the
next maintainer will trust it.

### D7 — Phase 3's dataclass names collide with 0.19.0's TypedDicts

**Status: open. Blocks Phase 3.** Found before writing any of it, by checking the proposed
names against the shipped package rather than against the plan.

The external review proposed eight dataclasses. Three of those names shipped publicly in
`xlsxturbo.types` two hours earlier, in 0.19.0:

| Name | Already in `xlsxturbo.types` as | Phase 3 would make it |
|------|--------------------------------|-----------------------|
| `SheetOptions` | the per-sheet options mapping accepted by `dfs_to_xlsx` | a dataclass of the same fields |
| `ValidationOptions` | **one column's** validation rule config | a group of validation-related kwargs |
| `ChartOptions` | **one chart's** config | a group of chart-related kwargs |

`ExportOptions`, `LayoutOptions`, `TableOptions`, `FormattingOptions` and `MediaOptions` are
free. None of the eleven `types.py` names is re-exported at package top level today — they are
reached as `xlsxturbo.types.X` — which bounds the damage but does not remove it.

The last two rows are the dangerous ones: same name, same package, **different meaning**,
distinguished only by which module you imported from. Renaming the dataclasses avoids the
clash and buys a worse problem — two parallel vocabularies for one set of concepts.

**`SheetOptions` is the finding that matters, though.** The TypedDict *is already* the
structured per-sheet options surface, public and type-checked since 0.19.0. A dataclass with
the same name and the same fields differs only in construction syntax. That is not obviously
worth a permanent touchpoint.

**Recommendation — build the entry-point object, not the feature-level ones.** The real
discoverability problem the review identified is that `df_to_xlsx` takes **27 parameters**;
it is not that a chart config is hard to spell, since `ChartOptions` is a typed TypedDict with
per-field documentation already. So:

- Build `ExportOptions` (the ~24 option kwargs, grouped), optionally with
  `LayoutOptions` / `TableOptions` / `FormattingOptions` / `MediaOptions` as nested fields
  **only if each earns its keep** — a group of two fields does not.
- Do **not** build dataclass `ValidationOptions` / `ChartOptions` / `SheetOptions`. They exist,
  they are typed, and they are the right shape already.

That drops the phase from eight new public names to between one and five, removes all three
collisions, and makes the eighth-touchpoint tax proportional to what it buys. The coverage
guard and D4's lower-to-kwargs design are unaffected.

### D8 — What the independent review found

The Phase 2 aftermath said a self-review catches "the whole approach is wrong" badly and that
an independent read was worth having before 1.0. It was run (`codex exec`, read-only, on
`833a889..c5e5cd9`, aimed at design rather than line-level defects). Every claim below was
re-verified here against the built extension before being recorded — an outside reviewer's
stated confidence is not evidence.

**Verdict: the surface is not safe to freeze unchanged.** Three defects shipped in 0.19.0 and
four are 1.0 design questions.

**All three defects are fixed in 0.19.1. All four design questions are resolved in 0.21.0**,
each one decided rather than inherited; the decisions and their reasoning are recorded inline
below. One consistency question surfaced while fixing the first defect and is recorded with
it below.

#### Defects in 0.19.0

**1. `docs/errors.md` states a guarantee that is false.** It promises *"every failure
xlsxturbo itself raises is an `XlsxTurboError`"*. It is not. The custom extractors still use
bare `extract()?` for nested keys and values, so PyO3's plain `TypeError` propagates
unclassified. Six of six probes escaped: a non-string `column_formats` key
(`src/extract.rs:369`), a non-string `formula_columns` value (`:422`), a bad `merged_ranges`
tuple element (`:445`), and the same shape in `comments`, `images`, `cells` and `hyperlinks`.

These are not argument-conversion failures happening before the library sees the value — they
are validations inside xlsxturbo's own extractors, which is exactly what the sentence claims
to cover.

**The test that should have caught this structurally cannot.**
`test_base_catches_everything_the_library_raises` iterates five hand-picked triggers, one per
exported class. It proves those five reach the hierarchy; it says nothing about the
population, so it passes at full strength while an entire family escapes. This is the
"a consistency check proves agreement, never completeness" failure, and it is the single
strongest argument for having run this review: a self-review shares the author's mental model
of what the test covers.

**Fixed in 0.19.1.** An `extract_typed!` macro classifies every nested conversion and names
the option; `pydict_to_hashmap` — the shared inner loop for every nested option dict — took a
required `context` parameter, because its message previously said a dict key was bad without
saying which option's. `TestNestedExtractionStaysInTheHierarchy` drives a 16-option probe
matrix whose population is **derived from `inspect.signature`**, so option N+1 cannot arrive
with an unclassified extractor and no failing test. Both directions were mutated: reverting one
site to a bare `extract()?` fails exactly the aimed-at probe, and parking a real extractor
option in the `SIGNATURE_CONVERTED` exclusion list fails the completeness check — the exclusion
list cannot swallow a gap.

**A consistency question fell out of it, and is open for 1.0.** `row_heights` and
`defined_names` are declared in the PyO3 signature as `HashMap<u32, f64>` and
`HashMap<String, String>`, so the binding converts them and a wrong inner type is a plain
`TypeError`. `column_widths` is declared `&Bound<PyAny>` and read by an extractor, so the
identical mistake there is a `ConfigurationTypeError`. Both are correct under the documented
carve-out, and from Python the three options look the same. Either make the two raw and
classify them, or accept the split — but decide it rather than inheriting it from which Rust
type someone reached for. Documented in `docs/errors.md` meanwhile.

**2. `xlsxturbo.types` has no `__all__`.** `PathLike`, `Literal`, `TypedDict` and `Union` are
all exported by `from xlsxturbo.types import *`. `tests/test_types_module.py` hides them
behind a hardcoded exclusion list, so it validates a cleaner namespace than users get, and
every future typing helper needs another exclusion.

**Fixed in 0.19.1.** `__all__` is declared and authoritative, `_runtime_shapes()` reads it, and
three new guards check it from both sides plus one that actually executes
`from xlsxturbo.types import *` — because a test that re-derives the answer from `__all__`
would pass even if the module stopped declaring one. Mutated both ways: dropping a name and
adding a nonexistent one each fail the aimed-at test.

**3. Every `TypedDict` is `total=False`, including where the field is mandatory.** Verified by
running them: `images={'D1': {}}`, `cells={'D1': {}}`, `charts={'D2': {}}`,
`sparklines={'D2': {}}` and `comments={'D1': {}}` each raise `ConfigurationError: missing
'...' key` at runtime, while a type checker accepts all five. The static contract is weaker
than the real one, in the module whose entire purpose is to state the real one. Python 3.9 does
not force this — a required base `TypedDict` plus a `total=False` subclass works there. Worth
fixing **before** Phase 3, or the dataclass and dict APIs disagree about what is required from
their first release together.

**Fixed in 0.19.1**, for nine shapes, each verified on a real 3.9 interpreter rather than
assumed. Each requirement is asserted twice — that the shape marks the field required, *and*
that the runtime rejects a dict without it — since either half alone is a contract with one
side missing. A partition check forces every shape into exactly one of "has required fields",
"fully optional", or the single documented conditional case (`ChartSeriesOptions`, whose
one-of a `TypedDict` cannot express), so a new shape cannot default to unexamined.

Pyright then found **four existing tests** passing dicts that omit a now-required key. All four
were tests deliberately checking the missing-key error, so each got a marker saying so — which
is the fix working, not collateral damage.

#### 1.0 design questions

**4. There is no class meaning "any bad option." — RESOLVED in 0.21.0, additively.** `ConfigurationError` (values) and
`ConfigurationTypeError` (types) are siblings — confirmed, `issubclass` is `False` — so a
caller wanting all configuration problems must catch a tuple, and `XlsxTurboError` also
catches file and input-data failures. This is the real category flaw.

The review proposed renaming `ConfigurationError` to `ConfigurationValueError` beneath a new
abstract parent. **A purely additive fix is available and better:** introduce
`OptionError(XlsxTurboError)` and reparent both classes under it. No rename, no removal, every
existing `except` clause keeps working, and `except OptionError` becomes possible.

**Shipped exactly as proposed.** `OptionError` takes no builtin base of its own, deliberately:
a builtin there would land on *both* children, so a `ConfigurationTypeError` would silently
also be a `ValueError` and the value/type split would stop meaning anything to `except`.
`FORBIDDEN_BASES` pins that.

The interesting part was not the class but the test. `test_every_concrete_class_has_a_trigger`
is what kept a dead `UnsupportedFeatureError` from shipping, and `OptionError` is *deliberately*
never raised — so it would have walked straight past that rule simply by being called abstract.
Instead `ABSTRACT` maps each abstract class to exactly the triggered classes it must catch,
asserted in both directions and again at runtime by actually catching them. An abstract class
that catches nothing raisable now fails.

**5. `FileError(XlsxTurboError, OSError, ValueError)`. — RESOLVED in 0.21.0: keep the base,
populate the field.** The layout hazard is theoretical here
— it works, because `ValueError` adds no fields to `OSError`'s struct. The demonstrated
oddity is `OSError`'s argument handling, which `FileError` fully inherits: `FileError('boom')`
has `errno is None`, while `FileError(2, 'x')` sets `errno` and changes `str()` to
`[Errno 2] x`. So the class is nominally an `OSError` with its structured fields permanently
unset, and `tests/test_errors.py` pins that as a contract.

The 1.0 question is whether `ValueError` stays. Dropping it and populating `errno`/`filename`
from the underlying OS error is the cleaner class — and it breaks `except ValueError` on file
failures, which D6 chose deliberately. That is a real trade, to be taken knowingly at 1.0
rather than inherited by default.

**The trade turned out to be false, because the two halves are separable.** Measured before
deciding: setting `errno` alone leaves `str()` **untouched**, while setting `filename` makes
`OSError.__str__` switch to `[Errno n] strerror: 'filename'` and *discard the message* — which
is where this library puts the path and the context, so populating it would have been strictly
worse. `errno` is now populated, the other two stay `None` on purpose, and the `ValueError`
base stays. The complaint the question was really about — an `OSError` whose structured fields
are permanently empty — is answered with no break at all.

Two things this needed that were not obvious. `save_workbook` stringified its `io::Error`
immediately, so the number had to be threaded out (`FileFailure`) rather than recovered. And
**`raw_os_error()` is only an `errno` on Unix**: on Windows it is a Win32 code from a different
numbering, where `ERROR_PATH_NOT_FOUND` is 3 and 3 as POSIX means `ESRCH`, "no such process".
Reporting that would be worse than reporting nothing, because it is wrong in a way that looks
right. Unix passes the number through; elsewhere it is classified via `io::ErrorKind`, the
portable view the standard library has already computed.

**6. `From<String> -> ConvertError::Config` is the wrong default. — RESOLVED in 0.21.0, more
strongly than proposed.** D6 already records two
misclassifications caused by it and files them under "last fraction of precision". The review
sharpens this usefully: the problem is not the two known instances but the *direction* — every
new untagged failure silently becomes `ConfigurationError`, i.e. blamed on the user, unless an
author remembers an invisible rule. **Make the fallback `Internal` instead**, so an omission
fails visibly rather than masquerading as bad configuration. That is cheap and does not require
the error-enum refactor D3 declined.

**Shipped as "no fallback at all", which is the better answer and was cheaper than it looked.**
The blanket `From<String>` and its `&str` twin are removed, so every site names `Config` or
`File` and a new failure site **does not compile** until it chooses. An omission is a build
error rather than a wrong exception in production — a stronger guarantee than any runtime
classification could give — and it avoids inventing an `Internal` variant, which would have
needed a public exception class nothing can raise, tripping the reachability rule from
question 4. A fallback that still exists is still a default, and the default was the bug.

The cost was measured before committing to it, by deleting the impl and counting what stopped
compiling: **14 sites, all in `convert.rs`**, not the ~70 that a naive grep for `?;` suggests.

Guarded by `TestConvertErrorHasNoDefaultCategory`, which reads the source, because the property
is the *absence* of code and nothing observable at runtime distinguishes that from a codebase
that never took the shortcut. It carries a control asserting both variants are still
constructed: deleting every `File` construction would satisfy the absence check perfectly while
restoring the exact behaviour it exists to prevent.

**7. The Python 3.9 annotation split. — RESOLVED in 0.21.0: keep 3.9, fix the annotations.** `typing.get_type_hints()` on these classes fails on
3.9 and works from 3.10 — documented in D1 and accepted. The review's point is that freezing
*documented broken introspection* at 1.0 buys nothing, and frameworks do introspect
annotations. Either drop 3.9 and use PEP 604 throughout, or keep 3.9 and spell field unions
with `Union[...]` too. Decide at 1.0; it resolves itself when 3.9 support ends.

**Decided: keep 3.9.** Dropping it would have solved the same problem while also removing
users — PEP 604 syntax is a maintainer convenience, not a user benefit — so 35 field
annotations were rewritten to `Union[...]`/`Optional[...]` instead. Verified on a real 3.9.25
interpreter: 15 shapes, all resolving.

The guard is the interesting part, because the obvious one is useless. A test that calls
`typing.get_type_hints` passes on 3.10+ **whether or not** the annotations use `|`, so a
regression introduced on a developer machine goes green locally and red only in the 3.9 CI leg.
The discriminating check is a source scan for PEP 604 unions, which fails on every version; the
runtime check is kept beside it for what a source scan cannot see, such as a forward reference
to a name that no longer exists. Both were mutated: reintroducing one `str | None` fails the
source scan and, as predicted, leaves the runtime check green.

#### Corrected

The review also reported that `tests/test_errors.py` claims picklability without testing it.
The test-quality half is right — it asserts `__module__`/`__qualname__`, which is a proxy, and
`EXPECTED_BASES` is a third hand-written copy of the hierarchy. But **pickling does work**:
round-tripping `FileError`, `ConfigurationError` and `ConfigurationTypeError` each returns the
same class object. A nitpick, not a defect.

---

## Phase 0 — Public-project hygiene

No code changes. Highest return per hour in the plan, and it addresses the weakest area in
the review.

### Supply-chain CI (`.github/workflows/ci.yml`)

- [x] `cargo audit` on every PR and push, with `--deny warnings`. Installed with
      `cargo install` rather than run through a third-party action: a job whose purpose
      is deciding what to trust should not widen the trusted set to do it
- [x] `pip-audit --strict --skip-editable` over the dev and build dependency set
- [x] CodeQL for **both `python` and `rust`**, as a `fail-fast: false` matrix
- [x] `actions/dependency-review-action` on pull requests, `fail-on-severity: moderate`
- [x] Pin every third-party action to a commit SHA with a trailing **full-version**
      comment (`# v7.0.1`, not `# v7`), plus an `action-pins` CI job that resolves each
      labelled version to a commit and compares it against the pinned SHA
- [x] Add an explicit `toolchain: stable` to every `dtolnay/rust-toolchain` use. That
      action takes the toolchain from its **ref name**, which a pinned SHA no longer
      carries — and a SHA from the `nightly` branch would look identical while silently
      changing the compiler
- [x] Top-level `permissions: contents: read`, widened per-job only where needed
      (`security-events: write` for CodeQL, `pull-requests: write` for dependency review)

**Correction to an earlier claim in this plan:** it said CodeQL has no Rust analyzer and
scoped the job to Python. That was wrong — `github/codeql-action` has `rust` in
`src/languages/index.ts`, and the experimental flag is only required below CodeQL 2.22.1.
Python-only would have covered mostly the thin binding layer and the tests while ignoring
the bulk of the codebase. Both languages now run. The Rust analyzer is newer than the
Python one, so if its leg proves unstable, drop that matrix entry rather than the job.

### Release provenance (`.github/workflows/release.yml`)

- [x] `actions/attest-build-provenance` over `dist/*`, with `attestations: write` added
      alongside the existing `id-token: write`. Runs **before** publish, so the
      attestation covers the exact bytes that go to PyPI
- [x] New `sbom` job: `cargo cyclonedx` for the Rust tree (the crates actually linked
      into the extension — this is the meaningful one) plus `cyclonedx-py` for the
      Python environment, which is thin by design since the wheel has no runtime deps
- [x] New `github-release` job creating the GitHub Release for each tag, with the wheels
      and SBOMs attached

That last item was not in the original plan and is worth more than the backfill below:
it stops the problem recurring. Every tag that builds and publishes now also gets its
Release, so tags and Releases cannot drift apart again.

### Dependabot (`.github/dependabot.yml`)

- [x] Added the `pip` ecosystem for the dev and build toolchain. Noted in the file that
      `uv.lock` is tracked but **not** updated by Dependabot — `uv lock` is still a
      manual step after merging one of those PRs, as `AGENTS.md` requires

### Community files

- [x] `CONTRIBUTING.md` — five-minute setup, the exact gate commands with their flags,
      and the touchpoint checklist
- [x] `SECURITY.md` — private reporting route, supported-version table, and an explicit
      scope section. Worth reading before triaging a report: xlsxturbo *writes* xlsx and
      never parses it, so the reader-side vulnerability classes largely do not apply
- [x] `CODE_OF_CONDUCT.md` — Contributor Covenant 2.1
- [x] Issue forms: `bug_report.yml` (platform, Python version and a minimal reproducer
      all required), `performance_regression.yml` (requires **interleaved** A/B rounds,
      because two back-to-back runs cannot separate a regression from thermal drift),
      `feature_request.yml` (requires the real task and the current workaround), plus a
      `config.yml` routing security reports away from public issues
- [x] `.github/pull_request_template.md`

### Release backfill

- [x] Created GitHub Releases for all 37 tags (2026-07-30). No assets attached, notes
      from the script, authored by `tstone-1`, `v0.18.0` marked latest
- [x] `.github/scripts/release-notes.sh` — the slicing logic, shared by the workflow and
      the backfill so there is one implementation to trust
- [x] Verified the script against **all 37 tags**: every one yields a non-empty section,
      a bogus version exits non-zero, and the boundary case stops correctly

**Correction to an earlier claim in this plan:** it said `CHANGELOG.md` uses a clean
`## [X.Y.Z] - YYYY-MM-DD` format throughout. It does not — 37 headings are bracketed and
**three are not** (`## 0.10.0`, `## 0.5.0`, `## 0.2.0`). A regex built naively from the
version turned `[0.18.0]` into a character class that matched `## 0.10.0` instead, and
returned plausible-looking notes from the wrong release. Hence fixed-string `index()`
matching in the script, and a terminator for both heading forms — without the second one,
v0.4.1's slice runs straight through `## 0.5.0`.

Arithmetic worth recording: 40 changelog sections, 37 tags. The three untagged versions
are exactly the three unbracketed ones, so the backfill covers 37 and nothing is missing.

Verified after the backfill: paginated count 37, `[.[].author.login] | unique` is
`["tstone-1"]`, no release carries assets, none is a draft or prerelease, no body is
empty, `releases/latest` is `v0.18.0`, and all 37 remote tag SHAs still match local
(nothing was moved or invented). Use the `--paginate` form for the count — the plain one
defaults to 30 items and under-reports.

Two hazards handled, worth keeping for any future backfill:

- **`gh release create` invents the tag if it is not on the remote**, pointing it at the
  default branch's HEAD. So the tag/SHA parity check runs *before* the loop, not after.
- **Attach nothing.** `gh` globs the filesystem, not git's index, so `.gitignore` is no
  protection: `python/xlsxturbo/*.pdb` sits in the working tree after any local
  `maturin develop`, and a Windows PDB embeds absolute source paths. The forward-looking
  `github-release` job attaches `dist/*` safely only because it runs on a clean runner.

### Housekeeping

- [x] Untrack `docs/strategic-recommendations-plan.md` (see D5)

**Gate:** CI green including every new job.

Verified locally before the first push:

- all three YAML files parse; the four issue forms parse with no duplicate field ids
- `cargo audit --deny warnings` — exit 0, 95 crate dependencies, no advisories
- `pip-audit --strict` over the 29 resolved dev/build dependencies — no vulnerabilities
- `.github/scripts/release-notes.sh` against all 37 tags, plus a failure control

One trap while doing that, worth not repeating: `python -m pip freeze` in a uv-created
venv returns nothing, because uv does not install `pip` into it. Piped into `pip-audit`
that produced a confident "No known vulnerabilities found" over an **empty** dependency
list — a check that could not fail, reading exactly like one that passed. Use
`uv pip freeze`, and always print the count of what is about to be audited.

Not verifiable locally, so the first push is the real test: both CodeQL legs, dependency
review (PR-only), the provenance attestation, the SBOM job, and the `github-release` job.

**Release:** none — infrastructure only.

---

## Phase 1 — Documentation split and capability matrix

- [x] MkDocs Material site, published to GitHub Pages by `.github/workflows/docs.yml`.
      Builds on every PR with `--strict`, deploys only from `main`. Docs dependencies are
      pinned in `requirements-docs.txt`, separate from the `dev` extras — building the
      site needs neither Rust nor the compiled extension
- [x] Cut `README.md` from 1289 lines to 121: what it does, install, one DataFrame
      example, one CSV example, performance ratios, the four limitations that most often
      surprise people, and links into the site
- [x] Move the remaining README content into pages. The set grew from the 12 planned to
      16, because four sections were large enough to earn their own page rather than being
      folded into a neighbour: `conditional-formatting`, `data-validation`, `cells`
      (arbitrary writes, hyperlinks, comments, checkboxes) and `performance`. The planned
      `getting-started` page was dropped instead — `index.md` does that job, and a separate
      page would have duplicated the install-and-first-example content the README already
      carries
- [x] **Capability matrix** — `docs/capability-matrix.md`, covering df / dfs / csv,
      per-sheet overridability, and constant-memory behaviour
- [x] **Generated, not hand-maintained** — `scripts/gen_capability_matrix.py` parses
      `SHEET_OPTION_NAMES`, `define_options!`, `CONSTANT_MEMORY_SAFE_OPTIONS`,
      `warn_constant_memory_skips` and the three `#[pyo3(signature)]` blocks
- [x] `tests/test_capability_matrix.py` — freshness check plus per-parser guards, and
      `scripts/` added to the ruff and pyright gates (it was an ungated corner)

Two generator properties worth keeping if this is ever rewritten:

- **Every parser raises rather than returning an empty list.** A structural audit that
  finds no instances of its own pattern reports nothing to flag, which is indistinguishable
  from a clean result.
- **Every parsed parameter must be a Python identifier.** That check is what caught the
  first signature parser, whose non-greedy regex spanned from the file's first pyo3
  attribute through to the requested function and yielded "parameters" like
  `) -> PyResult<(u32` — while still emitting a table that looked entirely plausible.

The tests pin the parsers against the **sources**, not against the generated page, so the
two cannot drift together. Both were shown to go red before being trusted: mutating the
page fails the freshness check, and both parser guards raise rather than returning quietly.

Two facts make the matrix cheap to build and worth building:

- `CONSTANT_MEMORY_SAFE_OPTIONS` is three entries (`column_widths`, `header_format`,
  `column_formats`). Everything else complex is skipped with a `RuntimeWarning`, and the
  README currently explains this in prose across ~15 lines.
- `csv_to_xlsx` takes five parameters (`input_path`, `output_path`, `sheet_name`, `parallel`,
  `date_order`) and no feature options at all. So the CSV column is almost entirely "not
  applicable" — an API asymmetry no current document states plainly. Surface it rather than
  letting the matrix hide it.

**Gate:** docs build clean; matrix freshness test passes; existing tests unaffected.
**Release:** none.

### Phase 1 aftermath

**The split was scripted, not retyped.** A one-off migration script cut the README into
atomic sections, assigned each to exactly one page or to an explicit drop list with a
reason, and then asserted three things: that the assignments partition the section list,
that each moved section's body appears **byte for byte** in exactly one destination, and
that every intra-README anchor link was rewritten the expected number of times. 33
sections moved, 5 deliberately did not. Two link assertions failed on the first run and
were real: one anchor occurred four times where the plan said three, and the "target page"
exemption had conflated *where a link sits* with *where its target lives*. Replacing that
with "the anchor must resolve to a heading on the destination page" is a strictly better
check, and it is the one that would catch a future page rename.

**A false claim on the front page, found because the split forced a check.** The README
advertised "Available as both Python library and CLI tool" and documented the CLI as
though `pip install xlsxturbo` provided it. It does not: maturin packages only the
extension module. This was settled by unzipping the published 0.18.0 wheel — no console
script, no binary — not by reading `Cargo.toml`, which says `default = ["cli"]` and would
have supported the wrong conclusion. Documentation-only fix, recorded in the changelog; if
shipping the binary is ever wanted, that is a packaging decision and a separate change.

**`mkdocs build --strict` does not protect against publishing a private file.** `docs/`
holds two files that are deliberately untracked (an internal planning memo, review notes)
plus the tracked-but-internal roadmap, and MkDocs publishes the whole directory. They are
named in `mkdocs.yml`'s `exclude_docs`, and `tests/test_docs_site.py` asserts that list
stays in step with `.gitignore`. What makes this worth writing down is the control:
removing one `exclude_docs` entry put the memo into `site/` **and `--strict` still exited
0**. The real protection is that deployment happens only from the workflow, from a clean
checkout with no untracked files present — which is why `CONTRIBUTING.md` says not to run
`mkdocs gh-deploy` by hand.

**A new test import cannot be validated locally, and declaring it as a dev dependency does
not help.** `tests/test_docs_site.py` imports `yaml`; `pyyaml` was added to the `dev`
extras, the full local suite passed, and three `python-test*` jobs then failed with
`ModuleNotFoundError`. Those jobs install an explicit minimal package list and never
install `dev` — so the local `.venv` is a strict superset of what CI has, and any new
third-party import in `tests/` is invisible until it fails there. Recorded in `AGENTS.md`,
since the fix is to edit three `pip install` lines that nothing points you at.

**One of the new tests was vacuous and the mutation harness caught it.** The orphan-page
check asked `git ls-files` which pages exist. Every new page was still uncommitted, so it
examined almost nothing and passed for the same reason an empty audit passes — the
mutation that removed a page from the nav survived. Rewritten against the filesystem,
which is what MkDocs actually publishes, and which cannot be empty. 5/5 mutations caught
afterwards. Same family as the generator's empty-list guard above, arriving through a
different door: there the population came from a regex, here from git.

**Phase 0 aftermath, all resolved — recorded because each was a gate behaving differently
than expected rather than a code bug:**

- `pip-audit --strict` and `--skip-editable` contradict each other: `--strict` fails on any
  skipped dependency and `--skip-editable` creates one. Fixed by auditing an explicit
  `pip freeze --exclude-editable` list, with the dependency count asserted non-empty.
- `python-lint` installed into the runner's system interpreter while `[tool.pyright]`
  points at `.venv`, so pyright analysed a different environment than the documented local
  command. It passed locally and failed in CI on the same tree. The job now builds `.venv`
  and runs all three tools from it, matching `CONTRIBUTING.md` exactly.
- A local `pytest` run showed 16 failures in `test_media.py` that were **not** real: the
  `.venv` extension was a week stale, reporting 0.17.2 against a 0.18.0 `Cargo.toml`. CI
  passing every `python-test` job on the same commit is what identified it.
  `maturin develop --release` cleared it. Note `test_version.py`'s drift guard cannot catch
  this — both sides of its comparison come from the same stale build.

All CI jobs green as of `8cc6620`, including both CodeQL legs. `Dependency review` only runs
on pull requests, and was confirmed passing on the two Dependabot PRs.

**The first Dependabot action bump exposed a hole in the pinning scheme**, and it is the
kind that never announces itself. Dependabot moved `dependency-review-action` to v5.0.0's
SHA and left the comment reading `# v4`, because it rewrites that comment only when it
matches the tag being replaced — and `# v4` was a shorthand, not the `v4.9.0` tag. A correct
SHA carrying a wrong human label, on a workflow that runs perfectly. Fixed three ways: full
versions in every comment, a `check-action-pins.sh` CI job comparing label to SHA, and the
false claim in `dependabot.yml` corrected to what Dependabot actually does.

Two implementation notes for that checker, both learned by getting them wrong first:

- Ask for **the commit at the tag**, not the tags at the SHA. `git/refs/tags` returns the
  tag *object's* SHA for annotated tags, which makes a correct pin look wrong — it did for
  `codeql-action`, whose pin is v4.37.4.
- It **fails when it checks nothing.** An extraction pattern that stops matching after a
  formatting change would otherwise report success over an empty set.

Dependency bumps merged this cycle: `rust_xlsxwriter` 0.96 → 0.97 (full local gate set green
on it: fmt, clippy, 80 Rust tests, 375 Python tests) and `dependency-review-action` v4 → v5
(a node20 → node24 runtime bump, no input changes, nothing blocking it).

---

## Phase 2 — Runtime types and exception hierarchy (0.19.0)

Both additive; both prerequisites for a credible 1.0. They ship together because they belong
to the same new public surface.

### Runtime types (see D1)

- [x] Create `python/xlsxturbo/types.py` with the option `TypedDict`s and the `Literal`
      aliases, 3.9-safe. **20 shapes, not 22** — the estimate counted two names that were
      never separate types. Moved with a script that asserts each definition appears byte for
      byte in the destination and that the original file is fully accounted for, rather than
      retyped.
- [x] **The 3.9 claim is verified, not assumed.** `uv python install 3.9`, then import the
      module on it: `PathArg = Union[...]` imports, `str | PathLike[str]` would not, and
      `get_type_hints()` fails there exactly as documented while working on 3.10+. Worth the
      three minutes — D1 reasoned this out correctly on paper, but the whole design rests on
      it and paper is not an interpreter.
- [x] Reduce `xlsxturbo.pyi` to the function signatures + exception classes, importing shapes
      from `xlsxturbo.types` with the redundant-alias form so `from xlsxturbo.xlsxturbo import
      HeaderFormat` still type-checks for code written before the move. 650 → 377 lines.
- [x] Update `__init__.pyi`'s module docstring — it instructed users to import the helpers
      from `xlsxturbo.xlsxturbo` under `TYPE_CHECKING`, which this phase obsoleted
- [x] Update `AGENTS.md` touchpoint 6 to match — it is now the largest touchpoint, since a new
      option must land in `types.py`, the stub's import block, the stub's `__all__`, both
      function signatures and `SheetOptions`
- [x] `CHANGELOG.md` entry superseding the note that these are stub-only types
- [x] `tests/test_types_module.py` — 27 tests coupling the stub's re-export list to the runtime
      module, since nothing in the type checker's world notices if those drift. All shown to
      fail: 6 mutations, 6 caught. **Two of the six survived the first run, and both were the
      tests' fault rather than the code's** — a sanity guard tight enough to fire on a single
      dropped name, which made the module fail to import so the intended comparison never ran;
      and `assert "from __future__ import annotations" in source`, which passed after the
      statement was deleted because `types.py` *documents* that import in its own docstring.
      The second is the "bookkeeping satisfies its own detector" trap, arriving through a
      docstring. Both are fixed and commented in place.

### Exception hierarchy (see D2, D3, and **D6 for what actually shipped**)

- [x] Define the hierarchy. The planned shape is below, kept for the record; the built shape
      and the four reasons it differs are in [D6](#d6--the-hierarchy-and-boundary-as-actually-built).

```
XlsxTurboError(Exception)
├── ConfigurationError(XlsxTurboError, ValueError)       # bad option key or value
├── UnsupportedFeatureError(XlsxTurboError, ValueError)  # e.g. constant_memory conflicts
├── InputDataError(XlsxTurboError, TypeError)            # unrecognised frame, bad dtype
├── WorkbookWriteError(XlsxTurboError, OSError)          # save_workbook, file I/O
└── WorkbookValidationError(XlsxTurboError, ValueError)
```

- [x] Register the classes on the extension module and export them from `__init__.py` /
      `__init__.pyi`. Built in `src/errors.rs` by calling the `type` metaclass, because
      `create_exception!` accepts only one base and every class here needs two or three.
      `__module__` is set to `xlsxturbo` so `repr()` and pickling follow the import path.
- [x] Classify the conversion sites by call site. **50 sites, not ~12**: 12 in `src/lib.rs`
      and 38 in `src/extract.rs` (22 `TypeError` -> `ConfigurationTypeError`, 16 `ValueError`
      -> `ConfigurationError`). Plus the two boundary additions D6 describes.
- [x] `tests/test_errors.py` — 51 tests. Asserts, per class, the new type, the legacy builtin
      base, **and the builtins it must *not* be** (without that third table, a class that
      inherited everything would pass). Also asserts every class is raised by a real call, and
      that the exported set and the triggered set are equal — so a class cannot be added,
      documented and never raised.
- [x] Every one of those tests shown to fail: 6 mutations of production code, 6 caught. Two
      harness bugs found on the way, both worth recording because each produced a *wrong
      verdict* rather than an obvious crash — a bare test name that does not match pytest's
      `name[param]` id read as `SURVIVED`, and `shutil.copy2` preserving mtime on restore left
      cargo with nothing to rebuild, so one mutation was scored against the *previous*
      mutation's binary. Same family as the padded-column parser already on record: the
      harness needs its own controls.
- [x] Document in the Phase 1 `errors` page

**Gate:** full set — `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`,
`cargo test`, `maturin develop --release`, ruff, bandit, pyright, pytest. Then a deep diff
review before Phase 3. **Release:** 0.19.0.

### Phase 2 aftermath

Both halves landed and the full gate set is green: `cargo fmt` 0, clippy 0, 80 Rust tests,
**460 Python tests** (up from 381 — 52 in `test_errors.py`, 27 in `test_types_module.py`),
ruff 0, bandit 0, pyright `0 errors, 0 warnings, 0 informations`, capability matrix current,
action pins `checked=14 failed=0`, `cargo audit` clean, `mkdocs build --strict` clean. Twelve
mutations run across the two new test modules, twelve caught.

Versions are bumped to 0.19.0 in `Cargo.toml`, `pyproject.toml` and `Cargo.lock`, the
changelog is dated, and everything through `885ddb0` is pushed with CI green.

**The phase-gate review is done** — deep, on the range `833a889..885ddb0`. Verdict APPROVE:
0 blockers, 4 warnings, 3 nitpicks, all seven fixed in the follow-up commit. Two things from
it are worth carrying forward rather than leaving in a gitignored report:

- **Three of the four warnings were documentation that lost a race with a fix in the same
  range** — a Rust doc comment still promising `TypeError` after `InputDataError` was moved to
  `ValueError`, this very section asserting the release had not happened, and the D6
  imprecision note listing one imprecision when there were two. None touched runtime
  behaviour, and all three would have misled the next reader about what is true. That is the
  bill for a high prose-to-code ratio (`src/errors.rs` is 223 lines, roughly 90 of them
  comment). The prose earns its place; the lesson is that **when a fix changes a decision,
  grep for every place that decision is written down** — not just the one you edited.
- **A self-review has a known blind spot and it should be stated, not assumed away.** This
  range was authored and reviewed in the same session, which catches mechanical defects and
  contract drift well and "the whole approach is wrong" badly, because the approach and the
  reviewer share an author. Worth an independent read (`codex exec`, or a human) before 1.0
  freezes this surface.

**0.19.0 is published** — PyPI has 5 wheels and an sdist, the GitHub release carries 8 assets,
and a disposable-environment install confirms the version, all six exception classes,
`xlsxturbo.types` importable from the wheel, a real export, and `FileError` catchable as both
`OSError` and `ValueError`.

The release failed on its first tag, and that failure is the durable part:

- **The reported symptom was not the problem.** Three smoke-test jobs reported
  `ModuleNotFoundError: No module named 'yaml'`, which reads as one missing dependency. It was
  not: reproducing the job locally showed **16** failures, not 3. `tests/test_docs_site.py` and
  `tests/test_capability_matrix.py` audit *repository* files, and the smoke test runs
  `pytest tests/` against an installed wheel from outside the checkout, with only `tests/`
  copied. Both modules were added after v0.18.0, so this was the first release that ran them.
  Installing pyyaml alone would have converted three import errors into sixteen assertion
  failures. Both now skip via `tests.helpers.repo_checkout_available()`; inside a checkout a
  missing file stays a hard failure.
- **The dependency list had drifted three times, so it became a file and a test.** It was
  inlined in four jobs; the fix after the first drift was a comment saying "remember the other
  copies", which did not survive a day. It is now `requirements-test.txt` plus
  `tests/test_ci_config.py`, which fails if a `tests/` import is undeclared or a workflow
  re-inlines the list. That guard's **first run found a third instance**: `tests/test_core.py`
  imports `numpy`, declared nowhere, working only because pandas pulls it in.
- **Moving the tag was safe and worth doing.** PyPI returned 404 and `gh release view` said
  "release not found", so nothing had been published and no version number was burned. Verify
  both before moving a release tag; if either says otherwise, bump instead.

Three things worth knowing before Phase 3 starts, none of them obvious from the diff:

- **The verification habit that paid off twice.** Building against a real 3.9 interpreter
  (`uv python install 3.9`) and inspecting the built wheel each caught an assumption that
  would otherwise have shipped. This is the same lesson as the Phase 1 CLI claim, which was
  settled by unzipping the published artifact rather than by reading `Cargo.toml`. Cheap;
  do it again for Phase 3's `options.py`.
- **The `pytest.raises(ValueError|TypeError)` count is now a number to watch, not just a fact.**
  It was 93 when the hierarchy landed, and every one of them is a pre-0.19 behaviour record.
  Phase 3 lowers dataclasses to kwargs, so its errors flow through exactly these paths.
- **`tests/test_errors.py` already encodes the anti-dead-API rule** (exported set == triggered
  set). Phase 3's coverage guard should be built the same way: derive both sides, compare them,
  and never hardcode the expected list in the test.

---

## Phase 3 — Structured options objects (0.20.0)

The expensive item, deliberately fourth. It is the review's top recommendation and its
highest-cost one; sequencing it here means it lands on top of the type and exception surfaces
it wants to reference.

**Blocked on a naming decision that 0.19.0 created — see D7.** Three of the eight names below
are already public as TypedDicts in `xlsxturbo.types`, and two of those three mean something
different there. Resolve D7 before writing `options.py`.

- [x] `python/xlsxturbo/options.py`: **`ExportOptions` only**, per D7. The other seven names
      were dropped — three collided with shipped `TypedDict`s and the rest were groups of two
      to five fields that did not earn a public class
- [x] Lower to kwargs in Python (see D4) — and **not** by wrapping the entry points, which is
      a deviation worth reading: a wrapper would have to duplicate 27 parameters or collapse
      them to `**kwargs`, and showing `**kwargs` where an editor shows 27 typed parameters
      costs more discoverability than the object adds. `as_kwargs()` / `as_sheet_options()`
      instead, leaving the compiled functions untouched
- [x] Keep every existing kwarg supported indefinitely. **Nothing deprecated**
- [x] Coverage guard: `TestCoverage` derives the option list from `inspect.signature` and
      fails in both directions. Mutated three ways — dropping a field, adding a field naming
      no real option, and emptying the workbook-only exclusion set — each caught by its
      aimed-at test
- [x] `AGENTS.md`: eighth touchpoint added, with what makes it affordable stated

**Unplanned find:** the full-bundle test wrote all 24 options at once and produced a file no
XML parser would read. Bisected to `data_bar` + `sparklines` on one worksheet, then reproduced
against **rust_xlsxwriter 0.97.0 alone** — no xlsxturbo code in the path — which emits three
`<ext>` elements and closes two. Upstream, unfixable here, and 0.97.0 is the latest release.
**Resolved by refusing the combination.** `apply::reject_databar_with_sparklines` raises
`ConfigurationError` before anything is written, naming the sheet, the column and two
workarounds. Refusing beats warning because a corrupt workbook is normally discovered by
whoever it was sent to, not by the person who wrote it.

`ConfigurationError` and deliberately *not* `WorkbookValidationError`: the latter is documented
as well-formed configuration that **Excel** forbids, and Excel is perfectly happy with a data
bar beside a sparkline. It is this writer that cannot produce it.

Two consequences worth carrying forward:

- **The guard makes the defect unreachable from Python, so no Python test can see it fixed.**
  `tests/upstream_defect.rs` drives rust_xlsxwriter directly and asserts the bug is *still
  present*; when upstream fixes it, that test fails and the guard gets deleted rather than
  outliving its reason. It has a control so a worse regression cannot be misread as the known
  one. `zip` was added as a dev-dependency for it — already in the tree via rust_xlsxwriter,
  with `default-features = false` because the default set pulls a `pbkdf2` version that does
  not resolve.
- **Over-reach is the expensive failure for a guard like this, not under-reach.** Refusing a
  combination that is actually fine removes a feature silently, because users read the error as
  their own mistake. Mutating the condition to refuse *any* conditional format beside sparklines
  is caught by `test_the_guard_is_narrow`, which pins the adjacent cases that must keep working.

**Accept the cost explicitly.** Adding a permanent eighth touchpoint per feature is a real
tax, paid for discoverability and typability. It is worth paying, but it is not free, and the
guard above is what keeps it from becoming a correctness problem as well as a cost.

**Gate:** full set plus the new guard; deep diff review. **Release:** 0.20.0.

---

## Phase 4 — Property tests and coverage visibility

- [x] `proptest` over `src/parse/` — 31 properties in `src/parse/proptests.rs`: the
      `A1 <-> (row, col)` round-trip, range corner ordering, colour parsing as an *iff*,
      table-name validity and idempotence, each pattern form as an equivalence to the `str`
      method it claims to implement, and the Excel serial scale as linearity from one anchor
- [x] Boundary tests: `src/parse/boundaries.rs` (the 1900 leap-year gap, Excel's first and
      last representable days, leap seconds), `src/write.rs` tests (the 2^53 cutoff, both
      `i64` extremes), and `tests/test_boundaries.py` end-to-end for all three
- [x] Coverage reporting: `scripts/coverage_report.py`, plus an informational CI job.
      No threshold anywhere

**Deliberately dropped from the review's proposal:** fuzzing the CSV parser. That path
delegates to the `csv` crate, which is fuzzed upstream — real setup cost, little marginal
value.

**Every test added here must be shown to go red before it is trusted.** Mutate deliberately,
one edit at a time, and write down which mutation each property is meant to catch *before*
running it. A property that holds by construction passes exactly like one that discriminates.

**Release:** rolls into whatever ships next — folded into the unreleased 0.20.0.

### What the mutation pass actually found

18 single-edit mutations, 15 against `src/parse/` and 3 against the compiled extension. All
were eventually caught, but not on the first attempt, and the two failures are the point of
the exercise:

- **A property can be correct and still never reach its case.** `six_ascii_chars_are_a_color_iff_all_hex`
  generated six printable ASCII characters and asserted the `iff`. Removing the
  `is_ascii_hexdigit` guard it exists to defend left it green: the inputs that discriminate
  (a sign character followed by five hex digits) are about one in seventy thousand of that
  space, so 256 cases never drew one. Narrowing the alphabet to hex digits plus `+` and `-`
  makes a near-miss appear in roughly one case in nine. A companion property whose generator
  *is* the case — sign, then five hex digits — was added beside it, because a generator that
  merely makes the case likely can still get unlucky.
- **A property can never enter the branch it was written for.** `sanitized_table_names_are_always_valid`
  asserted the 255-character cap over `".*"`, which produces short strings. Swapping the
  truncate and prepend steps — which pushes a 255-character digit-leading name to 256 —
  survived untouched. It now has a sibling generating names that straddle the cap.

**A property also caught its own author.** `strings_keep_their_original_padding` failed on
`"\tNan\t"`, which looks like a data-loss bug and is not: `Nan` is a valid `f64` literal, and
a non-finite number is deliberately written as an empty cell. The test encoded an assumption
the code documents the opposite of. Fixed by excluding numeric literals from the generator
and pinning the excluded case as its own property, which covers eight spellings where the
existing unit test covered one.

**Two design notes worth keeping:**

- `proptest-regressions/` is gitignored here, against the usual advice. Mutation-testing the
  suite makes proptest save a seed for every property that correctly went red, describing
  code that no longer exists. A genuine failing case is promoted to a named test instead,
  which states the input where it can be read rather than hiding it behind a hash.
- Coverage is measured from **both** suites, merged. `cargo test` alone reports 26% and shows
  every `src/apply/*.rs` at zero — a number that would be actively misleading, since those
  paths are exercised thoroughly from Python. Python alone misses the parser branches only
  the Rust tests reach. Together: 92.96% of lines in the Rust core, 100% of the Python layer.
  `cargo-llvm-cov` is not used, despite being the better tool for the Rust half, because its
  `report` subcommand cannot be pointed at an extra object file — and the extension module is
  exactly that.

### Carried into Phase 5

- **`NaN`/`Inf` spellings silently empty a text cell.** Documented and intentional, but the
  words `NAN`, `Inf`, `Infinity` are also ordinary text, and a column containing one loses it
  with no warning. Changing it is a behaviour change, so it belongs with the other surface
  decisions rather than in a test phase. Now at least documented in `docs/dataframe-export.md`.
- **The column bound is enforced by this crate, the row bound by rust_xlsxwriter.** Both raise
  `ConfigurationError` and both messages are clear, so this is cosmetic — but the two messages
  have different shapes for the same class of mistake.

---

## Phase 5 — 1.0.0

Only after phases 2 and 3, since both change the public surface. Shipping 1.0 and then
wanting an options object is the wrong order.

Landed as **1.0.0**, entirely in `docs/stability.md` plus the guard suite behind it. No
source behaviour changed — which was the point: by the time this phase started, phases 2-4
had already moved everything that needed moving, so 1.0 is the release that stops reserving
the right to move it again.

- [x] Stable names and semantics for `df_to_xlsx`, `dfs_to_xlsx`, `csv_to_xlsx`
- [x] Published deprecation policy — replacement first, `DeprecationWarning` naming the
      replacement *and* the removal version, at least one minor release and at least six
      months, removal only in a major
- [x] Supported Python and platform matrix stated as a promise
- [x] Stable exception model (delivered in Phase 2)
- [x] Stable interpretation of existing options
- [x] Documented compatibility guarantees for generated XLSX files
- [x] `Development Status :: 5 - Production/Stable` becomes accurate rather than
      contradictory

### What this phase found

- **The obvious reproducibility measurement lies.** Writing the same frame twice and hashing
  both files reported *identical* — which would have gone onto the page as "output is
  byte-reproducible". It was an artifact of both writes landing in the same clock second;
  `docProps/core.xml` embeds the creation time to one-second resolution. Separated by 1.1 s
  the files differ, and exactly one archive member is responsible. Both facts are now pinned
  by tests, and the test that proves non-reproducibility is the control for the one that
  names the single differing part.

  Same family as the timing traps in the personal agent-memory policy: nothing in the first
  measurement read a clock, and the whole result was about one.

- **A support matrix is four separate sources pretending to be one.** `requires-python`, the
  trove classifiers, the `ci.yml` interpreter matrix, and the `release.yml` wheel matrix all
  say part of it, none of them says all of it, and each moves without anyone opening the
  page. `tests/test_stability_policy.py` compares the page against all four, in both
  directions — a version added to CI and nowhere else fails just as loudly as a row deleted
  from the page.

- **Documenting the promise surfaced that 3.9 is already past EOL** (October 2025, and it is
  now July 2026). Kept deliberately — `requires-python = ">=3.9"` is a promise, and dropping
  it is a user-visible change rather than a tidy-up — but the page now states when it goes
  and why, instead of leaving a reader to infer that 3.9 is current.

- **Nine mutations, nine caught**, including the one that matters most: removing the
  deliberate 1.1 s wait turns the determinism pair red, so neither of them can be passing by
  accident.

---

## Deferred, with reasons

| Item | Why not in this cycle |
|------|----------------------|
| Benchmark CI with historical trends | Highest setup cost, weakest signal. On shared runners the noise routinely exceeds the ~20% regression it exists to catch. A local interleaved A/B is more trustworthy; see the measurement rules in the personal agent-memory policy. |
| Single-sourcing the version | The CI drift guard the review asks for already exists: `tests/test_version.py::test_version_matches_package_metadata` compares `version()` (compiled from `CARGO_PKG_VERSION`) against `importlib.metadata.version()` (wheel metadata from pyproject's static `version`, no `dynamic` key), on three operating systems. Single-sourcing is tidiness only. |
| Arrow tables / DataFrame interchange protocol | A genuinely good suggestion, and a feature. Feature work behind an unfixed API surface is what this plan exists to stop. Revisit after 1.0. |
| Append mode | Turns a writer into a partial workbook editor. `ROADMAP.md`'s existing reasoning stands. |
| Positioning as an openpyxl replacement | Discards the library's identity: a constrained, fast conversion path, not an in-memory workbook object model. |
| Internal `Result<_, String>` -> error enum | See D3. Large refactor, last fraction of precision, no public-contract dependency. |

---

## Baseline (verified 2026-07-30)

Point-in-time facts the plan was built on. **Re-verify before acting on any of them** — they
were true at `v0.18.0` and several are the targets of the phases above.

| Fact | Value |
|------|-------|
| `df_to_xlsx` parameters | 28, with `#[allow(clippy::too_many_arguments)]` |
| `csv_to_xlsx` parameters | 5, no feature options |
| Tests | 367 Python, 80 Rust |
| `README.md` | 1289 lines |
| Tracked files under `docs/` | 1 before this change, 0 after |
| GitHub tags / Releases | 37 / 0 |
| Error types raised | 25 `PyValueError`, 25 `PyTypeError`, 1 `PyKeyError`; no custom classes |
| Runtime-importable names | 4 (`csv_to_xlsx`, `df_to_xlsx`, `dfs_to_xlsx`, `version`) |
| Option `TypedDict`s | 22, stub-only |
| `CONSTANT_MEMORY_SAFE_OPTIONS` | 3 entries |
| Supply-chain CI | none of cargo-audit, pip-audit, CodeQL, dependency-review, SBOM, attestations |
| Actions pinning | all on floating tags |
| Dependabot ecosystems | `cargo`, `github-actions` (no `pip`) |
| Property tests / fuzzing / coverage | none |
| Already in place and credited | trusted PyPI publishing, wheel smoke-test before publish, version drift guard |
