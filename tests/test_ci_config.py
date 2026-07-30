"""Tests coupling the CI/release workflows to the test suite's real dependencies.

The local `.venv` holds the whole `dev` extra, but the CI test jobs and the
release smoke test install a deliberately minimal set. That makes the local
environment a strict superset of CI, so **a new third-party import in `tests/`
cannot be validated locally** -- it passes here and fails only in CI, or worse,
only in a release.

That has now happened twice in two days:

1. `tests/test_docs_site.py` imported `yaml`, which was declared in the `dev`
   extra. The three `python-test*` jobs never install `dev`, so all three went
   red on `ModuleNotFoundError`.
2. Fixing those three left a fourth copy of the same list in
   `release.yml`'s smoke-test job, which failed the v0.19.0 release after every
   wheel had already built.

A comment saying "remember the other copies" was the fix after (1), and it did
not survive one day. This module is that rule written as code instead.
"""

from __future__ import annotations

import ast
import importlib.util
import re

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="workflow tests audit repository files, which a wheel install does not carry",
)

TESTS_DIR = REPO_ROOT / "tests"
REQUIREMENTS = REPO_ROOT / "requirements-test.txt"
WORKFLOWS = REPO_ROOT / ".github" / "workflows"

# Import name -> distribution name, for the few that differ. Guarded below
# against rot: an entry naming something `tests/` no longer imports fails.
DISTRIBUTION_NAMES = {"yaml": "pyyaml"}

# Packages that must never be re-inlined into a workflow's `pip install`.
# Listing them by name rather than pattern-matching every install line keeps the
# check specific: workflows legitimately install other things (maturin, pip-audit).
MUST_COME_FROM_REQUIREMENTS = ("pandas", "polars", "openpyxl", "pytest", "pyyaml")


def _is_third_party(name: str) -> bool:
    """Whether `name` resolves to an installed third-party distribution.

    Resolved at runtime rather than against a hardcoded stdlib list, because
    `sys.stdlib_module_names` does not exist on Python 3.9 and this repo still
    supports it -- a check that quietly degrades on the oldest supported version
    is the one least likely to be noticed.
    """
    if name in {"xlsxturbo", "tests"}:
        return False
    try:
        spec = importlib.util.find_spec(name)
    except (ImportError, ValueError):
        return False
    if spec is None or spec.origin is None or spec.origin in {"built-in", "frozen"}:
        return False
    return "site-packages" in spec.origin.replace("\\", "/")


def _third_party_imports() -> set[str]:
    """Every third-party top-level module imported anywhere under `tests/`."""
    names: set[str] = set()
    for path in sorted(TESTS_DIR.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                names.add(node.module.split(".")[0])
    # A parser that matches nothing reads exactly like a clean result.
    assert len(names) >= 5, f"only parsed {len(names)} imports from tests/: {sorted(names)}"
    return {name for name in names if _is_third_party(name)}


def _declared_requirements() -> set[str]:
    """Distribution names declared in requirements-test.txt, lowercased."""
    lines = REQUIREMENTS.read_text(encoding="utf-8").splitlines()
    names = {
        re.split(r"[<>=!~;\[]", line.strip(), maxsplit=1)[0].strip().lower()
        for line in lines
        if line.strip() and not line.lstrip().startswith("#")
    }
    assert names, "requirements-test.txt declared nothing"
    return names


class TestTestDependenciesAreDeclared:
    """Every package `tests/` imports is installable in CI."""

    def test_every_third_party_import_is_in_requirements(self) -> None:
        """No test can import something the CI jobs will not have installed."""
        imported = _third_party_imports()
        declared = _declared_requirements()
        missing = {
            name
            for name in imported
            if DISTRIBUTION_NAMES.get(name, name).lower() not in declared
        }
        assert not missing, (
            f"tests/ imports {sorted(missing)}, which requirements-test.txt does not "
            f"declare. CI installs only that file, so these will fail there and pass "
            f"locally. Declared: {sorted(declared)}"
        )

    def test_distribution_name_map_has_no_stale_entries(self) -> None:
        """The import-name alias map does not outlive the imports it maps.

        Without this, a removed import leaves an alias behind and the map slowly
        becomes a list of things that used to be true.
        """
        imported = _third_party_imports()
        stale = set(DISTRIBUTION_NAMES) - imported
        assert not stale, f"DISTRIBUTION_NAMES maps {sorted(stale)}, which tests/ no longer imports"


class TestWorkflowsUseTheSharedRequirements:
    """No workflow re-inlines the dependency list that drifted twice."""

    def test_no_workflow_inlines_a_test_dependency(self) -> None:
        """Test packages are installed via `-r requirements-test.txt`, not by name."""
        offenders: list[str] = []
        for workflow in sorted(WORKFLOWS.glob("*.yml")):
            for number, line in enumerate(workflow.read_text(encoding="utf-8").splitlines(), 1):
                if "pip install" not in line or "requirements-test.txt" in line:
                    continue
                for package in MUST_COME_FROM_REQUIREMENTS:
                    if re.search(rf'["\s]{re.escape(package)}[<>=!~"\s]', line):
                        offenders.append(f"{workflow.name}:{number} installs {package} inline")
        assert not offenders, (
            "these lines re-inline a dependency that belongs in requirements-test.txt, "
            f"which is how the list drifted before: {offenders}"
        )

    def test_the_jobs_that_run_pytest_install_the_requirements(self) -> None:
        """Every workflow that runs pytest also installs the shared requirements.

        Counts rather than names the jobs, because job names change and the
        property that matters is that none of them runs the suite bare.
        """
        for workflow in sorted(WORKFLOWS.glob("*.yml")):
            text = workflow.read_text(encoding="utf-8")
            if "pytest" not in text:
                continue
            assert "requirements-test.txt" in text or ".[dev]" in text, (
                f"{workflow.name} runs pytest without installing requirements-test.txt"
            )
