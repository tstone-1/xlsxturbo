# Build & Release Guide

## Prerequisites

- Rust toolchain (stable): https://rustup.rs/
- Python 3.10+ (Python 3.14.6 is pinned for local development)
- uv: https://docs.astral.sh/uv/

## Local Development

### Setup

```bash
# Clone and enter directory
git clone https://github.com/tstone-1/xlsxturbo.git
cd xlsxturbo

# Create/sync the pinned environment and build the extension
uv sync --extra dev
uv run maturin develop --release
```

### Running Tests

```bash
# Rust unit tests
cargo test

# Python integration tests
uv run pytest tests/
```

### Code Quality Checks

```bash
# Format check (must pass before commit)
cargo fmt --check

# Linter (must pass with no warnings)
cargo clippy --all-targets -- -D warnings

# Format code (if check fails)
cargo fmt
```

## Pre-Push Checklist

Before pushing to main or creating a PR, verify all checks pass locally.

**These commands are copied verbatim from `.github/workflows/ci.yml`, flags included.**
That matters: a checklist that is weaker than the gate it exists to satisfy buys false
confidence, and the failure then arrives after the push. If you change a gate in CI,
change it here in the same commit. Two of these once drifted — `ruff` was missing
`scripts` and `cargo test` was missing `--release` — so a clean local run could still
go red on `main`. The reconciliation that fixed those then gave steps 6-8 only the
Unix venv path, which fails on Windows: **a checklist is only as good as its worst
platform**, so keep both columns below in step.

Platform-independent steps:

```bash
# 1. Format check                                    (CI: lint)
cargo fmt --check

# 2. Linter, warnings are errors                     (CI: lint)
cargo clippy --all-targets -- -D warnings

# 3. Rust tests, in release mode as CI runs them     (CI: test)
cargo test --release

# 4. Build the extension into the venv
maturin develop --release

# 9. Capability matrix is not stale                  (CI: python-test, via pytest)
python scripts/gen_capability_matrix.py --check

# 10. Supply chain                                   (CI: cargo audit / pip-audit)
cargo audit --deny warnings

# 11. Action pins still name their SHAs              (CI: Action pin comments)
bash .github/scripts/check-action-pins.sh
```

Steps 5-8 run tools from the project-local `.venv`, whose executables live in
`Scripts/` on Windows and `bin/` on macOS/Linux. This repo is worked on from both, so
take the column for the machine you are on — the same pair `AGENTS.md` documents:

| # | Gate | CI job | Windows | macOS / Linux |
|---|------|--------|---------|---------------|
| 5 | pytest | `python-test*` | `.venv\Scripts\python.exe -m pytest tests/ -q` | `.venv/bin/python -m pytest tests/ -q` |
| 6 | ruff — note `scripts` | `python-lint` | `.venv\Scripts\ruff.exe check python tests benchmarks scripts` | `.venv/bin/ruff check python tests benchmarks scripts` |
| 7 | bandit | `python-lint` | `.venv\Scripts\bandit.exe -c pyproject.toml -r python` | `.venv/bin/bandit -c pyproject.toml -r python` |
| 8 | pyright | `python-lint` | `.venv\Scripts\pyright.exe` | `.venv/bin/pyright` |

Steps 1-9 must succeed before pushing. Steps 10-11 rarely fail from a code change, but
they are gates on `main`, so a release must not skip them.

CI additionally runs `pip-audit`, CodeQL for Python and Rust, and — on pull requests
only — dependency review. Those need no local equivalent.

## Release Process

### 1. Update Version

Update version in both files (must match):

- `Cargo.toml`: `version = "X.Y.Z"`
- `pyproject.toml`: `version = "X.Y.Z"`

Follow SemVer:
- MAJOR: Breaking API changes
- MINOR: New features (backward compatible)
- PATCH: Bug fixes (backward compatible)

### 2. Update CHANGELOG.md

Add entry for new version with:
- Date
- Summary of changes
- Breaking changes (if any)

### 3. Commit Version Bump

```bash
git add Cargo.toml pyproject.toml CHANGELOG.md
git commit -m "Release X.Y.Z: <one-line summary of what ships>"
git push origin main
```

The `Release X.Y.Z: ...` subject is the convention every release since 0.17.0 has used;
`.github/scripts/release-notes.sh` does not depend on it, but the history is easier to
read when it is consistent.

### 4. Check Dependabot PRs

Before releasing, review open Dependabot PRs:

1. Go to: https://github.com/tstone-1/xlsxturbo/pulls
2. Check for open Dependabot PRs (dependency updates)
3. For each PR, decide:
   - **Merge** if CI passes and update is safe
   - **Close** if update causes issues or is not needed yet
   - **Defer** to next release (document why)

Don't release with unreviewed dependency PRs piling up.

### 5. Verify CI Passes

**IMPORTANT:** Before creating a release tag, verify GitHub Actions succeed.

1. Go to: https://github.com/tstone-1/xlsxturbo/actions
2. Check the latest push to `main`
3. Verify all CI jobs are green:
   - **CI / test (push)** - Rust tests pass
   - **CI / python-test, python-test-windows, python-test-macos (push)** - pytest passes against a maturin-built wheel on each OS
   - **CI / lint (push)** - Format and clippy pass
   - **CI / python-lint (push)** - ruff, bandit, and pyright pass

Do NOT proceed if CI is failing.

### 6. Create Release Tag

```bash
git tag vX.Y.Z
git push origin vX.Y.Z
```

### 7. Verify Release Workflow

After pushing the tag:

1. Go to: https://github.com/tstone-1/xlsxturbo/actions
2. Watch the **Release** workflow
3. Verify all jobs succeed:
   - **linux** (x86_64, aarch64)
   - **windows** (x64)
   - **macos** (x86_64, aarch64)
   - **sdist**
   - **smoke-test** (ubuntu/windows/macos) - pytest against the built wheels
   - **Publish to PyPI**

### 8. Verify PyPI Publication

1. Go to: https://pypi.org/project/xlsxturbo/
2. Verify new version appears
3. Test installation in a disposable environment: `uv run --with xlsxturbo==X.Y.Z python -c "import xlsxturbo; print(xlsxturbo.__version__)"`

## Troubleshooting

### CI Lint Fails

```bash
# Check what needs formatting
cargo fmt --check

# Auto-fix formatting
cargo fmt

# Check clippy warnings
cargo clippy --all-targets -- -D warnings
```

### Release Workflow Fails

1. Check which job failed in GitHub Actions
2. Common issues:
   - **Build fails**: Check Cargo.toml dependencies
   - **PyPI publish fails**: Check PyPI trusted publisher settings
   - **Wheel build fails**: Check maturin configuration

### Maturin Develop Doesn't Update

If changes aren't reflected after `maturin develop`:

```bash
# Resync and rebuild the editable extension
uv sync --extra dev
uv run maturin develop --release
```

## GitHub Actions Summary

| Workflow | Trigger | Jobs |
|----------|---------|------|
| CI | Push/PR to main | `test` (cargo test), `python-test` / `python-test-windows` / `python-test-macos` (pytest against a maturin-built wheel per OS), `lint` (fmt + clippy), `python-lint` (ruff + bandit + pyright) |
| Release | Push tag `v*` | Build wheels (linux/win/mac) + sdist + `smoke-test` (pytest against the built wheels) + PyPI publish |
