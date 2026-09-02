# Contributing to xlsxturbo

Thanks for your interest. xlsxturbo is a Rust extension module for Python, so the
setup has two halves; the whole thing should take about five minutes.

## Setup

You need a [Rust toolchain](https://rustup.rs/) (stable), Python 3.10 or newer, and
[uv](https://docs.astral.sh/uv/).

```bash
git clone https://github.com/tstone-1/xlsxturbo
cd xlsxturbo

# This repo uses a project-local .venv on purpose.
uv venv
uv pip install -e ".[dev]"

# Build the Rust extension into the venv.
maturin develop --release

# Confirm it worked.
python -c "import xlsxturbo; print(xlsxturbo.version())"
```

`--release` matters. A debug build of the extension is 20-50x slower, which makes
the test suite crawl and makes any timing you look at meaningless.

## Checks

Run these before opening a pull request. The flags are not optional: `cargo clippy`
without `--all-targets -- -D warnings` is a weaker check than the gate, and `cargo test`
without `--release` is a different build from the one CI runs.

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test --release
maturin develop --release
pytest tests/ -q
```

[BUILD.md](BUILD.md) holds the full pre-push checklist, copied verbatim from
`.github/workflows/ci.yml` and kept in step with it. When the two disagree, BUILD.md is
the one that was reconciled against the workflow; this list is the short form. Two of
these commands once drifted from the gate here, which is why the longer list exists and
why this one now points at it rather than trying to be a second copy.

Python lint, type and security gates:

| Gate | Windows | macOS / Linux |
|------|---------|---------------|
| ruff | `.venv\Scripts\ruff.exe check python tests benchmarks scripts` | `.venv/bin/ruff check python tests benchmarks scripts` |
| bandit | `.venv\Scripts\bandit.exe -c pyproject.toml -r python` | `.venv/bin/bandit -c pyproject.toml -r python` |
| pyright | `.venv\Scripts\pyright.exe` | `.venv/bin/pyright` |

If a tool is missing, install the dev extras rather than reaching for a system copy:
`uv pip install -e ".[dev]"`. A stale extension in the venv is the usual cause of
confusing signature mismatches -- rebuild with `maturin develop --release` and check
that the interpreter in pytest's header is the project `.venv`.

If you change the `dev` dependencies, run `uv lock` -- the lockfile is tracked.

## Documentation

The site under `docs/` is MkDocs Material, published to GitHub Pages from `main` by
`.github/workflows/docs.yml`. Its dependencies are pinned separately in
`requirements-docs.txt`, because building the site needs neither Rust nor the compiled
extension:

```bash
pip install -r requirements-docs.txt
mkdocs serve          # live preview on http://127.0.0.1:8000
mkdocs build --strict # what CI runs; fails on a broken link or an orphan page
```

Two things to know before editing:

- **`docs/capability-matrix.md` is generated.** Never edit it by hand. Regenerate with
  `python scripts/gen_capability_matrix.py --write`; `tests/test_capability_matrix.py`
  fails if the committed page is stale.
- **Do not run `mkdocs gh-deploy`.** Deployment happens only from the workflow, which
  builds from a clean checkout. A local deploy would publish whatever untracked files
  happen to be sitting in your `docs/` directory. `mkdocs.yml`'s `exclude_docs` names the
  ones that are known about, and `tests/test_docs_site.py` checks that list stays in step
  with `.gitignore` -- but the reliable protection is not deploying by hand.

A new page must be added to the `nav` in `mkdocs.yml`, or the test suite will fail it as
unreachable.

## Building a wheel

`maturin develop --release` installs into the venv, which is what you want while working.
To produce a distributable wheel instead:

```bash
maturin build --release   # writes to target/wheels/
```

The wheel contains only the Python extension module. The `xlsxturbo` command-line binary
is a separate Cargo `[[bin]]` target that is **not** packaged -- it is built by a plain
`cargo build --release` and has never shipped on PyPI.

## Adding a feature: the touchpoint checklist

An option threads through several layers of Rust and Python, and missing one is usually
a compile error or a named test failure rather than a silent bug -- but not always.
**The checklist is in [AGENTS.md](AGENTS.md#adding-a-feature---the-touchpoint-checklist)**,
with the reasoning behind each step and the guard test that enforces it. Follow it there
rather than from a summary: this section used to restate the steps, and the copy drifted
-- it still sent option shapes to the stub two releases after they moved to
`python/xlsxturbo/types.py`.

What the checklist will not tell you, because it is about mechanics rather than taste:

- **Reproducibility is a per-option decision.** Anything keyed by cell reference must be
  an `IndexMap`; iteration order feeds straight into the generated XML.
- **Errors carry their context.** A failure names the option and the key that caused it
  (`charts['D2']: ...`), so a caller with twenty entries knows which one to fix.
- **Every option needs a `constant_memory` classification**, and the default is
  "skipped with a warning". A guard test refuses to let you skip the decision.

Then update the relevant page under `docs/` and `CHANGELOG.md`. New per-feature examples
go on the docs page, not in `README.md` -- the README is a landing page, and it stopped
carrying per-option examples when the site was split out in 0.19.0.

## Tests

Tests read the generated workbook back rather than asserting on internal state. When
you add one, satisfy yourself that it can actually fail -- break the code on purpose,
one edit at a time, and check that your test is among what goes red. A test that
passes because it asserts something true by construction is indistinguishable from a
test that works.

## Pull requests

- Target `main`.
- Keep the change focused; unrelated cleanups are much easier to review separately.
- Say what you verified and on which platform. "Tests pass on Linux" is more useful
  than "tests pass".
- New behaviour needs a `CHANGELOG.md` entry under `## [Unreleased]`.

CI runs the Rust suite, the Python suite on Linux/Windows/macOS across several Python
versions, the lint gates, `cargo audit`, `pip-audit`, CodeQL and dependency review.

## Reporting bugs

Use the issue forms -- they ask for platform, Python version and a minimal input,
which are the three things that otherwise cost a round trip. For a security issue,
see [SECURITY.md](SECURITY.md); please don't open a public issue for it.

## Licence

By contributing you agree that your contribution is licensed under the MIT Licence,
as in [LICENSE](LICENSE).
