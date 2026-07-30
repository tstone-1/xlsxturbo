# Roadmap to 1.0

Working plan for the 0.19 -> 1.0 cycle. `ROADMAP.md` tracks *feature* gaps; this file
tracks the engineering work needed to turn a feature-complete library into a stable,
adoptable public package.

Written 2026-07-30 following an independent external project review. The review's findings
were fact-checked against the tree at `v0.18.0`; the baseline it produced is recorded under
[Baseline](#baseline-verified-2026-07-30) so a later session can tell what has since moved.

**How to use this file:** phases are ordered and mostly sequential — Phase 3 depends on
Phase 2's public surface, and Phase 5 must follow both. Phases 0 and 1 are independent of
everything and can be done in any order. Tick boxes as work lands. Each phase ends at the
gate set in `AGENTS.md`; phases 2 and 3 additionally get a deep diff review before the next
phase builds on them.

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

### D3 — Classify errors at the boundary, not by refactoring internals

Internal functions return `Result<_, String>`; the conversion to Python exceptions happens at
roughly a dozen sites in `src/lib.rs`. Classify there, by call site. Do **not** introduce an
internal error enum in this cycle — it touches all of `apply/`, `parse/` and `convert.rs` for
the last fraction of precision, and the public contract does not depend on it.

### D4 — Options objects lower to kwargs in Python

The dataclasses convert to the existing keyword arguments in a thin Python wrapper before
crossing into the extension. `RawOptions` / `extract_options` and the `define_options!`
machinery stay untouched. This is what bounds the cost; a second Rust extraction path would
not.

### D5 — `docs/strategic-recommendations-plan.md` is untracked

It was the only tracked file under `docs/` — an internal planning memo was what a visitor to
the public repo found there. Now gitignored alongside `docs/reviews/`. `docs/` from here on
holds user-facing documentation only, which is what Phase 1 fills it with.

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

- [ ] Create `python/xlsxturbo/types.py` with the 22 option `TypedDict`s and the `Literal`
      aliases, 3.9-safe
- [ ] Reduce `xlsxturbo.pyi` to the four function signatures, importing shapes from
      `xlsxturbo.types`
- [ ] Update `__init__.pyi`'s module docstring — it currently instructs users to import the
      helpers from `xlsxturbo.xlsxturbo` under `TYPE_CHECKING`, which this phase obsoletes
- [ ] Update `AGENTS.md` touchpoint 6 to match
- [ ] `CHANGELOG.md` entry superseding the note that these are stub-only types

### Exception hierarchy (see D2, D3)

- [ ] Define the hierarchy:

```
XlsxTurboError(Exception)
├── ConfigurationError(XlsxTurboError, ValueError)       # bad option key or value
├── UnsupportedFeatureError(XlsxTurboError, ValueError)  # e.g. constant_memory conflicts
├── InputDataError(XlsxTurboError, TypeError)            # unrecognised frame, bad dtype
├── WorkbookWriteError(XlsxTurboError, OSError)          # save_workbook, file I/O
└── WorkbookValidationError(XlsxTurboError, ValueError)
```

- [ ] Register the classes on the extension module and export them from `__init__.py` /
      `__init__.pyi`
- [ ] Classify the ~12 conversion sites in `src/lib.rs` by call site: `save_workbook` ->
      `WorkbookWriteError`, option extraction -> `ConfigurationError`, frame detection ->
      `InputDataError`
- [ ] `tests/test_errors.py` asserting, for each class, both the new type **and** the legacy
      builtin base — the legacy assertion is what proves D2 held
- [ ] Document in the Phase 1 `errors` page

**Gate:** full set — `cargo fmt --check`, `cargo clippy --all-targets -- -D warnings`,
`cargo test`, `maturin develop --release`, ruff, bandit, pyright, pytest. Then a deep diff
review before Phase 3. **Release:** 0.19.0.

---

## Phase 3 — Structured options objects (0.20.0)

The expensive item, deliberately fourth. It is the review's top recommendation and its
highest-cost one; sequencing it here means it lands on top of the type and exception surfaces
it wants to reference.

- [ ] `python/xlsxturbo/options.py`: `ExportOptions`, `SheetOptions`, `LayoutOptions`,
      `TableOptions`, `FormattingOptions`, `ValidationOptions`, `MediaOptions`, `ChartOptions`
- [ ] Lower to kwargs in Python (see D4)
- [ ] Keep every existing kwarg supported indefinitely, documented as the low-level form.
      **No deprecation in this phase**
- [ ] Coverage guard, in the spirit of `tests/test_option_coverage.py`: fail when a kwarg has
      no corresponding options-object field, so the two surfaces cannot drift
- [ ] Update `AGENTS.md`: the feature checklist grows an eighth touchpoint

**Accept the cost explicitly.** Adding a permanent eighth touchpoint per feature is a real
tax, paid for discoverability and typability. It is worth paying, but it is not free, and the
guard above is what keeps it from becoming a correctness problem as well as a cost.

**Gate:** full set plus the new guard; deep diff review. **Release:** 0.20.0.

---

## Phase 4 — Property tests and coverage visibility

- [ ] `proptest` over `src/parse/` — the strongest fit in the codebase: cell references,
      ranges, colors, table names, already 63 unit tests deep with cheaply stated invariants
      (`A1 <-> (row, col)` round-trip, range-normalisation idempotence, color parsing total
      over the accepted alphabet)
- [ ] Boundary tests: dates, Excel serial values, integers near 2^53
- [ ] Coverage reporting: `cargo-llvm-cov` for Rust, `coverage.py` for the Python layer.
      Publish as information, not as a threshold — the useful question is which
      error-handling branches are unexercised, which is exactly where Phase 2's
      classification wants evidence

**Deliberately dropped from the review's proposal:** fuzzing the CSV parser. That path
delegates to the `csv` crate, which is fuzzed upstream — real setup cost, little marginal
value.

**Every test added here must be shown to go red before it is trusted.** Mutate deliberately,
one edit at a time, and write down which mutation each property is meant to catch *before*
running it. A property that holds by construction passes exactly like one that discriminates.

**Release:** rolls into whatever ships next.

---

## Phase 5 — 1.0.0

Only after phases 2 and 3, since both change the public surface. Shipping 1.0 and then
wanting an options object is the wrong order.

- [ ] Stable names and semantics for `df_to_xlsx`, `dfs_to_xlsx`, `csv_to_xlsx`
- [ ] Published deprecation policy
- [ ] Supported Python and platform matrix stated as a promise (already `>=3.9`, abi3,
      `manylinux_2_28` — this is documentation, not change)
- [ ] Stable exception model (delivered in Phase 2)
- [ ] Stable interpretation of existing options
- [ ] Documented compatibility guarantees for generated XLSX files
- [ ] `Development Status :: 5 - Production/Stable` becomes accurate rather than
      contradictory — it currently sits on an 0.x version

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
