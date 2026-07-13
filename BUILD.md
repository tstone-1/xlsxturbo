# Build & Release Guide

## Prerequisites

- Rust toolchain (stable): https://rustup.rs/
- Python 3.9+ (Python 3.14.6 is pinned for local development)
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
cargo clippy -- -D warnings

# Format code (if check fails)
cargo fmt
```

## Pre-Push Checklist

Before pushing to main or creating a PR, verify all checks pass locally:

```bash
# 1. Format check
cargo fmt --check

# 2. Linter (no warnings)
cargo clippy -- -D warnings

# 3. Rust tests
cargo test

# 4. Build release
uv run maturin develop --release

# 5. Python tests
uv run pytest tests/

# 6. Ruff (Python lint)
uv run ruff check python tests benchmarks

# 7. Bandit (Python security)
uv run bandit -c pyproject.toml -r python

# 8. Pyright (Python types)
uv run pyright
```

All 8 steps must succeed before pushing.

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
git commit -m "chore: bump version to X.Y.Z"
git push origin main
```

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
cargo clippy -- -D warnings
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
