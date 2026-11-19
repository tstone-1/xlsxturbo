# Build, test and release

Moved verbatim from `AGENTS.md` on 2026-11-15, which keeps one index line per section.

## The project-local venv

This repo uses a project-local `.venv` (an exception to any central-venv convention). Test deps (`pytest pandas polars openpyxl`) must be installed there; if they are missing, `pytest` can resolve to a system Python with a stale extension and report bogus signature mismatches. Verify the interpreter in the pytest header is `.venv\Scripts\python.exe`; recover with `uv pip install -e ".[dev]" -r requirements-test.txt -r requirements-docs.txt` and rebuild via `maturin develop --release`.

## Never run a bare uv run

**Never run a bare `uv run` here; always `uv run --no-sync`** (or `--no-project` for a throwaway `--with` environment). A plain `uv run` first syncs `.venv` exactly to `uv.lock` with no extras, which uninstalls pytest, ruff, pyright, maturin, pandas, polars and mkdocs (54 packages by `uv sync --dry-run`) and replaces the `maturin develop` build. That is the usual way the test deps go missing: during the 1.5.1 release, BUILD.md's own rebuild step (then without `--no-sync`) removed pytest one line before the suite was due to run. `tests/test_ci_config.py::TestUvRunDoesNotResyncTheVenv` fails on a bare `uv run` in `BUILD.md`, `AGENTS.md`, `CONTRIBUTING.md` or `dev-docs/*.md`.

## A new third-party import in a test goes in requirements-test.txt

**A new third-party import in a test cannot be validated locally — add it to `requirements-test.txt`.** The local `.venv` holds the whole `dev` extras, but the CI test jobs and the release smoke test install only `requirements-test.txt`, so the local environment is a strict superset of CI and a test importing anything outside that file passes locally and fails only in CI. Declaring it in `[project.optional-dependencies] dev` does **not** fix it; those jobs never install `dev`. `tests/test_ci_config.py` fails if an import is undeclared or if a workflow re-inlines the list.

## The release smoke test runs outside the checkout

**The release smoke test runs `pytest tests/` against an installed wheel from outside the checkout, with only `tests/` copied.** So a test module that reads *repository* files — `mkdocs.yml`, the capability matrix, the generator script, the workflows — has nothing to read there and must skip via `tests.helpers.repo_checkout_available()`, not fail. This is invisible in ordinary CI and surfaces only when a tag is pushed: `test_docs_site.py` and `test_capability_matrix.py` were both added after v0.18.0, so v0.19.0 was the first release to run them and all 16 failed. Inside a checkout a missing file stays a hard failure — the guard distinguishes "no repository here" from "the repository is broken". Reproduce the job locally before tagging: `cp -r tests "$TMP/smoke" && cd "$TMP/smoke" && python -m pytest tests/ -q`.

## Confirming a release on PyPI

To confirm a release on PyPI, query the version-specific endpoint `https://pypi.org/pypi/xlsxturbo/<version>/json` (authoritative within seconds). The aggregate `https://pypi.org/pypi/xlsxturbo/json` `info.version` field lags several minutes behind (CDN cache) and can still show the previous version; trust the publish job's green status over it.

## Building against a crate release inside the 3-day window

The project resolves only releases at least three days old: `pyproject.toml` sets
`[tool.uv] exclude-newer = "3 days"` for the Python side, and on a machine whose Cargo
configuration sets `[registry] global-min-publish-age = "3 days"` (Cargo 1.100 or later),
`cargo update` does not resolve a crate version published less than three days ago and
moves a locked version that young back to an older one. To write and test the code for
such a release before then (done for rust_xlsxwriter 0.103.0 on 2026-11-19), use a
throwaway worktree, never the repo:

1. `git worktree add --detach <scratch>/wt HEAD`, then `git diff HEAD --binary | git -C <scratch>/wt apply`.
2. In the worktree's `Cargo.toml`, raise the requirement to the new version *and* add
   `[patch.crates-io] <crate> = { git = "<repo url>", tag = "v<version>" }`. Without the
   raised requirement cargo prints "patch ... was not used" and keeps the old version,
   because `"0.102.0"` means `^0.102`, which excludes 0.103.
3. `cargo update -p <crate>`, give the worktree its own venv
   (`uv venv`, then `uv pip install -e ".[dev]" -r requirements-test.txt`) and run every gate there.
4. Copy the changed files back except `Cargo.toml`'s `[patch]` block and `Cargo.lock`, and
   check each copy with `cmp`. The repo then fails `cargo build` with "failed to select a
   version" until the window ends, when one `cargo update -p <crate>` takes the registry copy.

   For 0.103.0 the window was not waited out, so that `main` could hold the work with a
   lockfile every clone can build: once the comparison below held,
   `CARGO_RESOLVER_INCOMPATIBLE_PUBLISH_AGE=allow cargo update -p rust_xlsxwriter` took the
   registry copy the same day, and every gate ran again in the repo. That variable overrides
   the minimum publish age for one command.

Building runs the new crate's code on the building machine early, which is what the window exists
to delay: check first that it has no `build.rs` and no new required dependencies.
Regenerate `THIRD-PARTY-LICENSES.md` after the registry update, not from the worktree:
the git source sorts the crate to another place in its license block, which
`tests/test_third_party_licenses.py` does not see (it compares names and versions) and
`scripts/gen_third_party_licenses.py --check` does. Before the copy-back, compare the
published `.crate` with the tag: its checksum against the index, its
`.cargo_vcs_info.json` commit against the tag, and every file with `cmp`
(`Cargo.toml.orig` against the tag's `Cargo.toml`).
