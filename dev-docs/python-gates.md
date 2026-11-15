# Python gates

Moved verbatim from `AGENTS.md` on 2026-11-15, which keeps one index line per section.

## Python Lint, Type, and Security Gates

The Python tree (`python/`, `tests/`, `benchmarks/`, `scripts/`) must stay clean under ruff, bandit, and pyright, with docstrings and type annotations on all functions. Config lives in `pyproject.toml`; the tools are in the `dev` optional-deps. These same three gates also run in CI (`python-lint` job in `.github/workflows/ci.yml`). Run from the repo root using the project-local `.venv`:

On Windows the venv's executables live in `.venv\Scripts\`, on macOS/Linux in `.venv/bin/` — this repo is worked on from both, so use the pair for the machine you are on:

| Gate | Windows | macOS / Linux |
|------|---------|---------------|
| ruff | `.venv\Scripts\ruff.exe check python tests benchmarks scripts` | `.venv/bin/ruff check python tests benchmarks scripts` |
| bandit | `.venv\Scripts\bandit.exe -c pyproject.toml -r python` | `.venv/bin/bandit -c pyproject.toml -r python` |
| pyright | `.venv\Scripts\pyright.exe` | `.venv/bin/pyright` |
| pytest | `.venv\Scripts\python.exe -m pytest tests/ -q` | `.venv/bin/python -m pytest tests/ -q` |

If a tool is missing from the venv, install the dev extras (`uv pip install -e ".[dev]"`) rather than reaching for a system copy; `uvx <tool>` also works for a one-off check and needs no venv.

Scoping notes (intentional, do not "fix" by widening):
- pyright runs `typeCheckingMode = "standard"` project-wide, with the shipped library raised to strict via the top-level `strict = ["python/xlsxturbo"]` path list. Do not use `executionEnvironments` + `typeCheckingMode` for this — that key is silently ignored by pyright 1.1.x.
- pyright sets `pythonPlatform = "Linux"`, the platform of the `python-lint` CI job, so the gate gives the same answer on Windows and macOS. Code under a `sys.platform == "win32"` branch is therefore not type-checked.
- bandit scans `python/` only; tests and benchmarks are excluded (asserts and non-crypto `random` data generation are expected there).
- ruff per-file-ignores: `S101` in tests; `S404/S603/S607/S311/T201` in benchmarks. Google docstring convention.
- When changing the `dev` deps, run `uv lock` (the lockfile is tracked).
- pandas-stubs rejects `pd.to_datetime([..., pd.NaT, ...])` (mixed `list[str | NaTType]`); use the string `"NaT"` instead — pandas parses it to NaT, keeping test data identical.
