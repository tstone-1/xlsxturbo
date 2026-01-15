# Build & Release Guide

## Prerequisites

- Rust toolchain (stable): https://rustup.rs/
- Python 3.8+
- maturin: `pip install maturin`

## Local Development

### Setup

```bash
# Clone and enter directory
git clone https://github.com/tstone-1/xlsxturbo.git
cd xlsxturbo

# Install in development mode
pip install -e ".[dev]"
# Or with maturin:
maturin develop --release
```

### Running Tests

```bash
# Rust unit tests
cargo test

# Python integration tests
python tests/test_features.py

# Or with pytest
pytest tests/
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
maturin develop --release

# 5. Python tests
python tests/test_features.py
```

All 5 steps must succeed before pushing.

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
3. Verify both workflows show green checkmarks:
   - **CI / test (push)** - Rust tests pass
   - **CI / lint (push)** - Format and clippy pass

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
   - **Publish to PyPI**

### 8. Verify PyPI Publication

1. Go to: https://pypi.org/project/xlsxturbo/
2. Verify new version appears
3. Test installation: `pip install xlsxturbo==X.Y.Z`

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
# Uninstall and reinstall
pip uninstall xlsxturbo -y
pip install .
```

## GitHub Actions Summary

| Workflow | Trigger | Jobs |
|----------|---------|------|
| CI | Push/PR to main | `test` (cargo test), `lint` (fmt + clippy) |
| Release | Push tag `v*` | Build wheels (linux/win/mac) + PyPI publish |
