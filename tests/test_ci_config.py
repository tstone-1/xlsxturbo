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
import yaml

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="workflow tests audit repository files, which a wheel install does not carry",
)

TESTS_DIR = REPO_ROOT / "tests"
REQUIREMENTS = REPO_ROOT / "requirements-test.txt"
PANDAS2_REQUIREMENTS = REPO_ROOT / "requirements-test-pandas2.txt"
WORKFLOWS = REPO_ROOT / ".github" / "workflows"
DEPENDABOT = REPO_ROOT / ".github" / "dependabot.yml"
PYPROJECT = REPO_ROOT / "pyproject.toml"
UV_LOCK = REPO_ROOT / "uv.lock"

# Majors deliberately held back, mapped to the Python 3.x minor their next major
# requires. The hold is legitimate only while this project supports something
# older; once it does not, the entry -- and the Dependabot ignore behind it --
# must go.
#
# Empty as of 1.1.0: `pytest` lived here while the floor was 3.9, and raising it
# to 3.10 released the hold, exactly as the test below demanded.
HELD_BACK_MAJORS: dict[str, int] = {}

# Import name -> distribution name, for the few that differ. Guarded below
# against rot: an entry naming something `tests/` no longer imports fails.
DISTRIBUTION_NAMES = {"yaml": "pyyaml"}

# Packages that must never be re-inlined into a workflow's `pip install`.
# Listing them by name rather than pattern-matching every install line keeps the
# check specific: workflows legitimately install other things (maturin, pip-audit).
MUST_COME_FROM_REQUIREMENTS = ("pandas", "polars", "openpyxl", "pytest", "pyyaml")


def _is_third_party(name: str) -> bool:
    """Whether `name` resolves to an installed third-party distribution.

    Resolved at runtime rather than against a stdlib name list. `sys.
    stdlib_module_names` is available on every supported version since 1.1.0
    dropped 3.9 -- the reason this was written this way -- but it answers a
    weaker question: "not in the standard library" also includes a local module
    or a namespace package, neither of which belongs in a requirements file.
    Checking that the resolved origin is under `site-packages` answers the one
    that matters.
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


def _build_backend_maturin_floor() -> tuple[int, int]:
    """The maturin floor `pyproject.toml`'s `[build-system]` requires."""
    text = PYPROJECT.read_text(encoding="utf-8")
    match = re.search(r'^requires = \["maturin>=(\d+)\.(\d+)', text, re.MULTILINE)
    assert match, "pyproject.toml declares no maturin build requirement"
    return int(match.group(1)), int(match.group(2))


def _workflow_maturin_floors() -> list[tuple[str, int, tuple[int, int]]]:
    """Every `maturin>=X.Y` spec a workflow installs, as (file, line, floor)."""
    found: list[tuple[str, int, tuple[int, int]]] = []
    for workflow in sorted(WORKFLOWS.glob("*.yml")):
        for number, line in enumerate(workflow.read_text(encoding="utf-8").splitlines(), 1):
            for match in re.finditer(r"maturin>=(\d+)\.(\d+)", line):
                found.append(
                    (workflow.name, number, (int(match.group(1)), int(match.group(2))))
                )
    return found


class TestWorkflowMaturinFloorsMatchTheBuildBackend:
    """A CI job that installs an older maturin than `pyproject.toml` requires.

    `[build-system] requires = ["maturin>=1.9,<2.0"]` is load-bearing: PEP 639
    `license-files` is what puts `THIRD-PARTY-LICENSES.md` into the wheel, and
    maturin gained it in 1.9.0. An older backend ignores the key and builds a
    wheel with no notice and no error (see `tests/test_third_party_licenses.py`,
    which pins the pyproject floor itself).

    The four `pip install "maturin>=..."` lines in `ci.yml` sat at `>=1.4` while
    pyproject said `>=1.9`, so every CI job was free to resolve a backend that
    could not honour the declaration the release depends on. Nothing compared
    the two numbers, because they live in different files and neither is wrong
    on its own.
    """

    def test_the_scan_finds_the_specs_it_audits(self) -> None:
        """Control: a sweep matching nothing reads exactly like a clean result.

        Every assertion below is a loop over what this scan returns, so an
        empty scan -- a renamed workflow, a reworded install line, a regex that
        stopped matching -- would pass them all while auditing nothing.
        """
        floors = _workflow_maturin_floors()
        assert floors, (
            "no 'maturin>=X.Y' spec was found in any workflow, so the floor comparison "
            "below is vacuous; the scan, not the workflows, is what to fix"
        )
        assert any(name == "ci.yml" for name, _, _ in floors), (
            "ci.yml declares no pinned maturin floor, but it is the workflow that "
            f"builds the extension for every test job; found only {sorted({n for n, _, _ in floors})}"
        )

    def test_no_workflow_installs_an_older_maturin_than_the_backend_needs(self) -> None:
        """Every workflow floor is at least the `[build-system]` floor."""
        backend_floor = _build_backend_maturin_floor()
        stale = [
            f"{name}:{number} installs maturin>={floor[0]}.{floor[1]}"
            for name, number, floor in _workflow_maturin_floors()
            if floor < backend_floor
        ]
        assert not stale, (
            f"pyproject.toml's build backend requires maturin>="
            f"{backend_floor[0]}.{backend_floor[1]}, which is what carries the PEP 639 "
            f"license files into the wheel, but these lines allow an older one: {stale}"
        )


class TestBothPandasMajorsAreExercised:
    """A declared range is only supported at the end that CI actually runs.

    `requirements-test.txt` allows `pandas>=2.3.3,<4`, and pip resolves to the
    newest every time. Without a second leg, pandas 2 would be declared
    supported and never executed -- which is not a hypothetical: it was the state
    of this repository until 1.1.0, in the opposite direction. The ceiling said
    `<3` while pandas 3 was released and was what the developer machine ran, so
    local runs and CI had silently stopped testing the same library.

    pandas 3 changed defaults this project is directly exposed to (Copy-on-Write,
    PyArrow-backed strings as the default `str` dtype), so "type detection
    behaves the same on both" is a claim that deserves a job rather than an
    assumption. Both majors pass today; this keeps that measurable.

    Scoped to pandas deliberately: it is the only dependency here whose declared
    range currently spans two majors. A second one would need its own leg, and
    nothing in this file would notice -- said out loud because a guard that reads
    as general when it is specific is worse than one that admits its scope.
    """

    def test_the_layered_file_constrains_only_pandas(self) -> None:
        """The pandas-2 file overrides one ceiling and re-uses every other floor.

        Layering is what keeps the floors single-sourced. A file that restated
        the dependency list would be the fourth copy of it, which is the drift
        this module exists to prevent.
        """
        lines = [
            line.strip()
            for line in PANDAS2_REQUIREMENTS.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        ]
        assert lines[0] == "-r requirements-test.txt", (
            f"{PANDAS2_REQUIREMENTS.name} must layer requirements-test.txt, not restate it; "
            f"its first directive is {lines[0]!r}"
        )
        overrides = lines[1:]
        assert overrides == ["pandas<3"], (
            f"{PANDAS2_REQUIREMENTS.name} should override the pandas ceiling and nothing "
            f"else, but declares {overrides}"
        )

    def test_a_workflow_actually_installs_it(self) -> None:
        """The leg exists in CI, not merely the file.

        The silent failure this guards: deleting the matrix `include` entry
        leaves the file in the tree, looking like coverage that no longer runs.
        """
        installed_by = [
            workflow.name
            for workflow in sorted(WORKFLOWS.glob("*.yml"))
            if PANDAS2_REQUIREMENTS.name in workflow.read_text(encoding="utf-8")
        ]
        assert installed_by, (
            f"{PANDAS2_REQUIREMENTS.name} is not referenced by any workflow, so pandas 2 is "
            f"declared supported in requirements-test.txt and never run"
        )

    def test_the_declared_range_still_spans_two_majors(self) -> None:
        """The leg's reason still holds.

        The other direction, and the one that expires: if the pandas ceiling is
        ever lowered back to `<3`, this file and its CI leg are redundant and
        should go rather than linger as a second way of testing the only major
        that is left.
        """
        text = REQUIREMENTS.read_text(encoding="utf-8")
        match = re.search(r"^pandas>=(\d+)[^,]*,<(\d+)", text, re.MULTILINE)
        assert match, "could not read a bounded pandas requirement from requirements-test.txt"
        floor_major, ceiling_major = int(match.group(1)), int(match.group(2))
        assert ceiling_major - floor_major > 1, (
            f"requirements-test.txt now allows only pandas {floor_major}.x, so "
            f"{PANDAS2_REQUIREMENTS.name} and its CI leg are redundant -- remove both, and "
            f"this test with them"
        )


def _project_version() -> str:
    """The version declared in pyproject.toml's `[project]` table."""
    text = PYPROJECT.read_text(encoding="utf-8")
    section = re.split(r"^\[project\]$", text, maxsplit=1, flags=re.MULTILINE)
    assert len(section) == 2, "pyproject.toml has no [project] table"
    body = re.split(r"^\[", section[1], maxsplit=1, flags=re.MULTILINE)[0]
    match = re.search(r'^version\s*=\s*"([^"]+)"', body, re.MULTILINE)
    assert match, "could not read a version out of pyproject.toml's [project] table"
    return match.group(1)


def _locked_version(package: str) -> str:
    """The version `uv.lock` records for `package`.

    Read with a regex rather than `tomllib`, which arrived in 3.11 while this
    project still supports 3.10, and rather than a third-party TOML parser,
    which would have to be declared in requirements-test.txt for one field.
    """
    blocks = UV_LOCK.read_text(encoding="utf-8").split("[[package]]")
    assert len(blocks) > 1, "uv.lock declares no packages"
    for block in blocks[1:]:
        if not re.search(rf'^name = "{re.escape(package)}"$', block, re.MULTILINE):
            continue
        match = re.search(r'^version = "([^"]+)"$', block, re.MULTILINE)
        assert match, f"uv.lock's {package} entry records no version"
        return match.group(1)
    raise AssertionError(f"uv.lock has no [[package]] entry for {package}")


class TestTheLockfileTracksTheProjectVersion:
    """`uv.lock` pins this project itself, so a release has to re-lock.

    AGENTS.md already says to run `uv lock` when the dev dependencies change,
    and a version bump is such a change -- the lockfile records `xlsxturbo`'s
    own version. The 1.1.0 release did not re-lock, so the tracked lockfile went
    on resolving 0.21.0, and since the same commit raised `requires-python` to
    `>=3.10` it also kept resolution legs for an interpreter this project no
    longer supports. Nothing noticed: no job and no other test reads `uv.lock`.

    A failure here means one thing -- a version was bumped without `uv lock`.
    Run it and commit the result.
    """

    def test_the_locked_version_matches_pyproject(self) -> None:
        """The `xlsxturbo` entry in uv.lock names the version being shipped."""
        locked, declared = _locked_version("xlsxturbo"), _project_version()
        assert locked == declared, (
            f"uv.lock locks xlsxturbo {locked} while pyproject.toml declares {declared}. "
            f"Run `uv lock` and commit the result."
        )


def _oldest_supported_python_minor() -> int:
    """The 3.x minor named by `requires-python` in pyproject.toml."""
    text = PYPROJECT.read_text(encoding="utf-8")
    match = re.search(r'^requires-python\s*=\s*"[^"]*?>=\s*3\.(\d+)', text, re.MULTILINE)
    # A parser that matches nothing would make every hold below look justified.
    assert match, "could not read a `>=3.x` floor out of requires-python in pyproject.toml"
    return int(match.group(1))


def _majors_dependabot_ignores() -> set[str]:
    """Packages whose major updates Dependabot is configured to skip."""
    config = yaml.safe_load(DEPENDABOT.read_text(encoding="utf-8"))
    ignored: set[str] = set()
    for update in config["updates"]:
        for entry in update.get("ignore", ()):
            if "version-update:semver-major" in entry.get("update-types", ()):
                ignored.add(str(entry["dependency-name"]).lower())
    return ignored


class TestDeferredMajorUpgradesExpireWithTheirReason:
    """A silenced Dependabot PR stops arriving, so its reason needs a deadline.

    This worked once already: `pytest` 9 needs Python 3.10, the floor was 3.9,
    and the weekly PR was silenced by an `ignore` entry. An ignore is invisible
    once written -- the PR simply never comes back -- so nothing would have
    announced that the reason had expired. Tying it to `requires-python` meant
    raising the floor to 3.10 in 1.1.0 failed this test until the hold was
    released with it, which is exactly what happened.

    **`HELD_BACK_MAJORS` is currently empty, so two of the three tests below are
    dormant** -- they iterate an empty mapping and pass without examining
    anything. That is honest rather than clean: a zero population is not a clean
    result, and saying so is the point of this paragraph. The live one is
    `test_no_silenced_major_lacks_a_recorded_expiry`, which needs no entries to
    do its job and is what stops the next hold being written down nowhere.
    """

    def test_no_silenced_major_lacks_a_recorded_expiry(self) -> None:
        """Every Dependabot major-version ignore is registered here.

        The direction that stays live with no holds recorded, and the one that
        matters: an ignore added to `.github/dependabot.yml` on its own silences
        a PR forever with nothing to expire it. Registering the package in
        `HELD_BACK_MAJORS` is what subjects it to the test below.
        """
        unrecorded = _majors_dependabot_ignores() - set(HELD_BACK_MAJORS)
        assert not unrecorded, (
            f".github/dependabot.yml silences major updates for {sorted(unrecorded)} with "
            f"no entry in HELD_BACK_MAJORS, so nothing will ever say the hold expired. Add "
            f"each one, mapped to the Python 3.x minor its next major requires."
        )

    def test_each_hold_matches_the_supported_python_range(self) -> None:
        """Every held-back major is silenced exactly while its reason holds.

        Dormant while `HELD_BACK_MAJORS` is empty; see the class docstring.
        """
        oldest = _oldest_supported_python_minor()
        ignored = _majors_dependabot_ignores()
        for package, needs_minor in HELD_BACK_MAJORS.items():
            if oldest < needs_minor:
                assert package in ignored, (
                    f"{package}'s next major needs Python 3.{needs_minor} and this "
                    f"project supports 3.{oldest}, so the weekly PR cannot be merged. "
                    f"Add a semver-major ignore for {package} to .github/dependabot.yml, "
                    f"or drop it from HELD_BACK_MAJORS if the hold no longer applies."
                )
            else:
                assert package not in ignored, (
                    f"this project now requires Python 3.{oldest}, so {package}'s next "
                    f"major (which needs 3.{needs_minor}) is installable and the hold is "
                    f"obsolete. Remove the ignore from .github/dependabot.yml, raise the "
                    f"floor in requirements-test.txt, and drop the entry here."
                )

    def test_every_hold_names_a_package_the_suite_installs(self) -> None:
        """The hold list does not outlive the dependency it is about.

        Same rot as DISTRIBUTION_NAMES above: an entry for a package the suite
        no longer uses reads as an active decision and is nothing of the kind.
        """
        declared = _declared_requirements()
        stale = {package for package in HELD_BACK_MAJORS if package not in declared}
        assert not stale, (
            f"HELD_BACK_MAJORS holds back {sorted(stale)}, which requirements-test.txt "
            f"no longer declares"
        )


CHANGELOG = REPO_ROOT / "CHANGELOG.md"

# `## [1.2.3] - 2026-08-15`, `## [Unreleased]`, or the legacy unbracketed
# `## 0.10.0 - 2026-01-16` that 0.2.0, 0.5.0 and 0.10.0 still use.
_VERSION_HEADING = re.compile(r"^## \[(?P<version>\d[^\]]*)\](?P<rest>.*)$")


def _version_headings() -> list[tuple[str, str]]:
    """Every bracketed version heading in CHANGELOG.md, as (version, rest-of-line)."""
    return [
        (match.group("version"), match.group("rest"))
        for line in CHANGELOG.read_text(encoding="utf-8").splitlines()
        if (match := _VERSION_HEADING.match(line))
    ]


class TestChangelogHeadings:
    """A version heading must carry a date, because the release script matches a prefix.

    `.github/scripts/release-notes.sh` selects a section with
    `index($0, "## [<version>]") == 1` -- a prefix test that ignores the rest of
    the line. So `## [1.1.2] - Unreleased` is matched by tag `v1.1.2` exactly as
    a dated heading would be, and the release job publishes a GitHub Release
    whose notes came from a heading saying Unreleased. Nothing fails.

    Measured both ways against this repo's own CHANGELOG: the `- Unreleased`
    form exits 0 and prints the section; `## [Unreleased]`, which carries no
    version, exits 1 with "no CHANGELOG section found" and fails the release job
    before anything is published. Unreleased work therefore accumulates under
    `## [Unreleased]` and gets its number and date at release time (BUILD.md
    step 2). This test is that rule written as code, because the unsafe form is
    the one a habit from other repos produces.
    """

    def test_the_file_has_version_headings_to_check(self) -> None:
        """The emptiness control.

        `_VERSION_HEADING` matching nothing would make every assertion below
        pass over an empty list, which reads exactly like a clean result.
        """
        headings = _version_headings()
        assert len(headings) > 30, (
            f"only {len(headings)} bracketed version headings found in CHANGELOG.md; "
            f"the heading pattern has stopped matching and the checks below are inert"
        )

    def test_no_version_heading_is_marked_unreleased(self) -> None:
        """The failure this class exists for."""
        marked = [
            version
            for version, rest in _version_headings()
            if "unreleased" in rest.lower()
        ]
        assert not marked, (
            f"CHANGELOG.md has version heading(s) {marked} marked Unreleased. "
            f"release-notes.sh matches '## [<version>]' as a prefix, so tagging that "
            f"version publishes a release whose notes say Unreleased, and no job fails. "
            f"Use '## [Unreleased]' with no version until release time; BUILD.md step 2 "
            f"is where it gains its number and date."
        )

    def test_every_version_heading_carries_a_date(self) -> None:
        """A heading with a version but no date is half-renamed, which is the same bug."""
        undated = [
            version
            for version, rest in _version_headings()
            if not re.match(r"^ - \d{4}-\d{2}-\d{2}$", rest)
        ]
        assert not undated, (
            f"CHANGELOG.md version heading(s) {undated} lack a ' - YYYY-MM-DD' date. "
            f"A released version is dated; unreleased work belongs under '## [Unreleased]'."
        )


# `uses: <owner>/<repo>/<path>@<40-hex sha> # v<version>` in a workflow.
_CODEQL_PIN = re.compile(
    r"uses:\s*(github/codeql-action/[\w-]+)@([0-9a-f]{40})\s*#\s*(v[\d.]+)"
)


def _codeql_pins() -> dict[str, tuple[str, str]]:
    """Map each pinned `codeql-action` entry point to its (SHA, version comment)."""
    pins: dict[str, tuple[str, str]] = {}
    for workflow in sorted(WORKFLOWS.glob("*.yml")):
        for path, sha, version in _CODEQL_PIN.findall(
            workflow.read_text(encoding="utf-8")
        ):
            pins[path] = (sha, version)
    return pins


class TestCodeqlActionPinsMoveTogether:
    """`init`, `autobuild` and `analyze` must be pinned to one SHA.

    They are three entry points of one action reading one config file, and its
    ``getConfig`` throws ``Loaded a configuration file for version 'X', but
    running version 'Y'`` on any mismatch — so a workflow mixing versions fails
    CodeQL, whichever two agree.

    Dependabot files a separate PR per path, so ungrouped it produces three PRs
    that each move one pin. None can go green alone, and merging them in
    sequence reddens CodeQL on `main` until the last lands. That happened for
    4.37.4 -> 4.37.6 (PRs #32-34) and again for 4.37.6 -> 4.37.7 (#35-36),
    where the `open-pull-requests-limit` was full and the third PR was never
    filed at all — so merging both open ones would have left `analyze` behind
    with nothing to say so.

    `dependabot.yml`'s `groups:` entry is what makes them arrive as one PR.
    This is what fails if they drift regardless — a hand edit, a revert, or a
    group that stops matching.
    """

    def test_all_three_entry_points_are_pinned(self) -> None:
        """The emptiness control: a pattern that matches nothing passes vacuously."""
        pins = _codeql_pins()
        assert {
            "github/codeql-action/init",
            "github/codeql-action/autobuild",
            "github/codeql-action/analyze",
        } <= set(pins), (
            f"expected all three codeql-action entry points to be pinned in "
            f"{WORKFLOWS}, found {sorted(pins)}; the pin pattern has stopped "
            f"matching and the check below is inert"
        )

    def test_every_codeql_pin_names_the_same_sha(self) -> None:
        """The failure this class exists for."""
        pins = _codeql_pins()
        shas = {sha for sha, _ in pins.values()}
        assert len(shas) == 1, (
            f"codeql-action pins point at {len(shas)} different SHAs: "
            f"{ {path: sha for path, (sha, _) in pins.items()} }. They read one "
            f"shared config file and CodeQL fails on a version mismatch, so all "
            f"of them move in one commit or none do."
        )

    def test_every_codeql_pin_carries_the_same_version_comment(self) -> None:
        """A matching SHA under two labels means one comment is lying."""
        pins = _codeql_pins()
        versions = {version for _, version in pins.values()}
        assert len(versions) == 1, (
            f"codeql-action pins carry {len(versions)} different version comments: "
            f"{ {path: version for path, (_, version) in pins.items()} }. The SHAs "
            f"agree or the test above would have caught it, so at least one comment "
            f"is stale — which is invisible in a green build."
        )

    def test_dependabot_groups_the_codeql_action_paths(self) -> None:
        """Without the group the three PRs come back, one per entry point."""
        config = yaml.safe_load(DEPENDABOT.read_text(encoding="utf-8"))
        actions = [
            update
            for update in config["updates"]
            if update["package-ecosystem"] == "github-actions"
        ]
        assert len(actions) == 1, "expected one github-actions Dependabot entry"
        patterns = [
            pattern
            for group in actions[0].get("groups", {}).values()
            for pattern in group.get("patterns", ())
        ]
        assert any("codeql-action" in pattern for pattern in patterns), (
            "dependabot.yml no longer groups github/codeql-action*, so its three "
            "entry points will arrive as three PRs again, none of which can pass "
            "on its own"
        )
