# What this changes

<!-- One or two sentences. If it fixes an issue, "Fixes #123". -->

# Why

<!-- What was wrong or missing. Skip if it is obvious from the above. -->

# How it was verified

<!--
Say what you actually ran and on which platform. "Tests pass on Linux, Python 3.12"
beats "tests pass". If you added a test, say what you did to confirm it can fail --
a test that passes for the wrong reason looks exactly like one that works.
-->

# Checklist

- [ ] `cargo fmt --check`
- [ ] `cargo clippy --all-targets -- -D warnings`
- [ ] `cargo test --release`
- [ ] `maturin develop --release` then `pytest tests/ -q`
- [ ] ruff, bandit and pyright clean (see CONTRIBUTING.md for the commands)
- [ ] `CHANGELOG.md` updated under `## [Unreleased]`
- [ ] The relevant page under `docs/` updated, if behaviour or options changed
- [ ] Option shapes updated in `python/xlsxturbo/types.py` (the stub imports them),
      if options changed
- [ ] `uv lock` re-run, if dev dependencies changed

<!--
Adding a new option? It threads through several layers -- follow the checklist in
AGENTS.md ("Adding a Feature") rather than from memory. Two steps that are easy to
miss: the option name must go into SHEET_OPTION_NAMES, and you must decide its
constant_memory classification. Guard tests enforce both, so CI will tell you, but
knowing up front is quicker.

Per-feature examples belong on the docs page for that feature, not in README.md.
-->
