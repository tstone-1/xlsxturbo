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
