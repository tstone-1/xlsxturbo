# xlsxturbo Agent Instructions

## Shared Memory Policy

- `AGENTS.md` is the canonical shared memory for Codex and other coding agents in this repository.
- Claude Code loads this file through `.claude/CLAUDE.md`.
- Durable project knowledge, workflows, commands, architecture notes, and recurring pitfalls belong here.
- Do not store durable project knowledge only in Claude auto memory.
- Keep entries concise, specific, and verifiable. Prefer updating existing sections over appending duplicate notes.

## Git Workflow

- Only commit and push when explicitly asked by the user.
- Do not include Claude-related or AI-generated footers in commit messages.
- Before commit or push, run `cargo update` to check for Rust dependency updates.
- Follow `BUILD.md` before release or push-ready work.

## Account Enforcement

- Before any commit, run: `git config user.email "48162401+tstone-1@users.noreply.github.com"` and `git config user.name "tstone-1"`.
- Before any push, run: `gh auth switch --user tstone-1`.
- Do not use unrelated work or organization accounts in this repository.

## Build, Test, and Release

- Use `uv` for Python dependency and command execution.
- Standard local checks: `cargo fmt --check`, `cargo clippy -- -D warnings`, `cargo test`, `maturin develop --release`, then `pytest tests/`.
- Plain `cargo test` must work outside maturin. Keep `pyo3/extension-module` enabled through `pyproject.toml` / maturin, not directly in `Cargo.toml`.
- Release versions are SemVer and must match in `Cargo.toml` and `pyproject.toml`; update `CHANGELOG.md` before release commits.
- Before tagging a release, verify the latest GitHub Actions CI on `main` is passing and no relevant Dependabot PRs are unreviewed.

## Python Lint, Type, and Security Gates

The Python tree (`python/`, `tests/`, `benchmarks/`) must stay clean under ruff, bandit, and pyright, with docstrings and type annotations on all functions. Config lives in `pyproject.toml`; the tools are in the `dev` optional-deps. Run from the repo root using the project-local `.venv`:

- `.venv/Scripts/ruff.exe check python tests benchmarks`
- `.venv/Scripts/bandit.exe -c pyproject.toml -r python`
- `.venv/Scripts/pyright.exe`
- `.venv/Scripts/python.exe -m pytest tests/ -q`

Scoping notes (intentional, do not "fix" by widening):
- pyright runs `typeCheckingMode = "standard"` project-wide, with the shipped library raised to strict via the top-level `strict = ["python/xlsxturbo"]` path list. Do not use `executionEnvironments` + `typeCheckingMode` for this — that key is silently ignored by pyright 1.1.x.
- bandit scans `python/` only; tests and benchmarks are excluded (asserts and non-crypto `random` data generation are expected there).
- ruff per-file-ignores: `S101` in tests; `S404/S603/S607/S311/T201` in benchmarks. Google docstring convention.
- When changing the `dev` deps, run `uv lock` (the lockfile is tracked).

## Benchmarks

- The main comparison suite is `benchmarks/benchmark.py`; use `--markdown` to regenerate the README performance table and `--json` for machine-readable output.
- The parallel CSV conversion suite is `benchmarks/benchmark_parallel.py`.
- README performance numbers are system-specific and should identify the machine, OS, Python version, and run methodology.
