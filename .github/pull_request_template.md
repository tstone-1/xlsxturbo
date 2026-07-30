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
- [ ] `cargo test`
- [ ] `maturin develop --release` then `pytest tests/ -q`
- [ ] ruff, bandit and pyright clean (see CONTRIBUTING.md for the commands)
- [ ] `CHANGELOG.md` updated under `## [Unreleased]`
- [ ] `README.md` updated, if behaviour or options changed
- [ ] Type stubs updated in `python/xlsxturbo/xlsxturbo.pyi`, if options changed
- [ ] `uv lock` re-run, if dev dependencies changed

<!--
Adding a new option? It threads through seven places -- the checklist is in
AGENTS.md ("Adding a Feature"). Two that are easy to miss: the option name must go
into SHEET_OPTION_NAMES, and you must decide its constant_memory classification.
Guard tests enforce both, so CI will tell you, but knowing up front is quicker.
-->
