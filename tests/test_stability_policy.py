"""Tests coupling `docs/stability.md` to the things it makes promises about.

A stability page is a document that claims to describe configuration held
elsewhere: the classifiers in `pyproject.toml`, the interpreter matrix in
`ci.yml`, the wheel targets in `release.yml`, the exported names in
`__init__.py`. Every one of those changes without anyone opening this page, and
a support promise that has quietly stopped being true is worse than none --
somebody is depending on it.

So the page is not prose to be kept in step by hand. Each table below is
compared against its actual source, in both directions, and the one behavioural
claim it makes about generated files is measured rather than asserted.
"""

from __future__ import annotations

import hashlib
import re
import time
import zipfile
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo
import yaml

from tests.helpers import REPO_ROOT, TIMESTAMPED_PART, repo_checkout_available

# Applied per class, not to the module: the table comparisons read repository
# files and cannot run against an installed wheel, but the determinism tests are
# about the wheel and are exactly what the release smoke test should exercise.
needs_checkout = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="these tests audit repository files, which a wheel install does not carry",
)

STABILITY = REPO_ROOT / "docs" / "stability.md"
PYPROJECT = REPO_ROOT / "pyproject.toml"
CI = REPO_ROOT / ".github" / "workflows" / "ci.yml"
RELEASE = REPO_ROOT / ".github" / "workflows" / "release.yml"


def _markdown_table(heading: str) -> list[list[str]]:
    """Rows of the first Markdown table under `heading` in docs/stability.md.

    Returns the body rows only, each already split into stripped cells.
    """
    text = STABILITY.read_text(encoding="utf-8")
    start = text.find(f"\n## {heading}\n")
    assert start != -1, f"docs/stability.md has no '## {heading}' section"
    section = text[start + 1 :]
    end = section.find("\n## ")
    if end != -1:
        section = section[:end]

    rows: list[list[str]] = []
    for line in section.splitlines():
        stripped = line.strip()
        if not stripped.startswith("|"):
            if rows:
                break  # the table ended; ignore any later one in the section
            continue
        cells = [cell.strip() for cell in stripped.strip("|").split("|")]
        if all(set(cell) <= {"-", ":"} for cell in cells):
            continue  # the ---|--- separator
        rows.append(cells)
    # A parser that matches nothing would make every comparison below vacuous.
    assert len(rows) >= 2, f"parsed {len(rows)} rows from the '{heading}' table"
    return rows[1:]  # drop the header row


def _classifier_pythons() -> set[str]:
    """Python versions claimed by the trove classifiers in pyproject.toml."""
    found = set(
        re.findall(
            r'"Programming Language :: Python :: (3\.\d+)"',
            PYPROJECT.read_text(encoding="utf-8"),
        )
    )
    assert found, "no versioned Python classifiers found in pyproject.toml"
    return found


def _ci_pythons() -> set[str]:
    """Interpreter versions any CI job runs the test suite against.

    Reads `include:` entries as well as the matrix axes. A version reachable
    only through an `include` is just as tested as one on an axis, and reading
    only the axes would report it as untested -- the direction that quietly
    understates coverage rather than overstating it, and therefore the one that
    would sit unnoticed.
    """
    config = yaml.safe_load(CI.read_text(encoding="utf-8"))
    found: set[str] = set()
    for job in config["jobs"].values():
        matrix = job.get("strategy", {}).get("matrix", {})
        found.update(str(version) for version in matrix.get("python-version", ()))
        for entry in matrix.get("include", ()):
            if "python-version" in entry:
                found.add(str(entry["python-version"]))
    assert found, "no python-version matrix found in ci.yml"
    return found


def _release_wheel_targets() -> set[tuple[str, str]]:
    """(os, architecture) pairs the release workflow builds wheels for."""
    config = yaml.safe_load(RELEASE.read_text(encoding="utf-8"))
    found: set[tuple[str, str]] = set()
    for name, job in config["jobs"].items():
        matrix = job.get("strategy", {}).get("matrix", {})
        for target in matrix.get("target", ()):
            found.add((name, str(target)))
    assert found, "no wheel target matrix found in release.yml"
    return found


def _built_wheel_artifacts() -> set[str]:
    """Artifact names the wheel-building jobs upload, e.g. `wheels-linux-aarch64`.

    Derived from the same `(job, target)` pairs the jobs interpolate into their
    upload step, so a new build target appears here without anyone maintaining a
    list.
    """
    return {f"wheels-{job}-{target}" for job, target in _release_wheel_targets()}


def _smoke_tested_artifacts() -> set[str]:
    """Artifact names the release workflow's smoke-test matrix actually installs."""
    config = yaml.safe_load(RELEASE.read_text(encoding="utf-8"))
    matrix = config["jobs"]["smoke-test"]["strategy"]["matrix"]
    found = {str(entry["wheel-artifact"]) for entry in matrix["include"]}
    assert found, "the smoke-test matrix installs no wheel artifacts"
    return found


@needs_checkout
class TestSupportedPythonTable:
    """The Python table states the classifiers and the CI matrix, not a memory of them."""

    def test_the_table_lists_exactly_the_classified_versions(self) -> None:
        """Adding or removing a classifier without the page is caught."""
        documented = {row[0] for row in _markdown_table("Supported Python versions")}
        assert documented == _classifier_pythons(), (
            "docs/stability.md and pyproject.toml's classifiers disagree about which "
            f"Python versions are supported: page has {sorted(documented)}, classifiers "
            f"have {sorted(_classifier_pythons())}"
        )

    def test_the_ci_column_matches_the_workflow_matrix(self) -> None:
        """The "Run in CI" column is the ci.yml matrix, in both directions."""
        claimed = {row[0] for row in _markdown_table("Supported Python versions") if row[2] == "yes"}
        assert claimed == _ci_pythons(), (
            "docs/stability.md claims CI runs "
            f"{sorted(claimed)} but ci.yml's matrix is {sorted(_ci_pythons())}"
        )

    def test_every_ci_version_is_a_supported_one(self) -> None:
        """CI does not test an interpreter the project does not claim to support.

        The opposite direction of the test above, and the one that catches a
        version added to the matrix and nowhere else -- which reads as a widened
        promise when it is only a widened test.
        """
        unsupported = _ci_pythons() - _classifier_pythons()
        assert not unsupported, (
            f"ci.yml runs {sorted(unsupported)}, which pyproject.toml does not classify as "
            f"supported. Either add the classifier and the docs row, or stop testing it."
        )

    def test_the_declared_floor_is_the_oldest_supported_version(self) -> None:
        """`requires-python` agrees with the oldest row on the page."""
        match = re.search(
            r'^requires-python\s*=\s*">=\s*(3\.\d+)"', PYPROJECT.read_text(encoding="utf-8"), re.M
        )
        assert match, "could not read requires-python from pyproject.toml"
        oldest = min(_classifier_pythons(), key=lambda version: int(version.split(".")[1]))
        assert match.group(1) == oldest, (
            f"requires-python says >={match.group(1)} but the oldest supported version "
            f"documented and classified is {oldest}"
        )


@needs_checkout
class TestSupportedPlatformTable:
    """The platform table states the release workflow's wheel matrix."""

    def test_the_table_lists_exactly_the_built_wheels(self) -> None:
        """A new wheel target, or a dropped one, must reach the page."""
        rows = _markdown_table("Supported platforms")
        documented = {(row[0].lower(), row[1].split()[0].lower()) for row in rows}
        built = {(job, target.lower()) for job, target in _release_wheel_targets()}
        assert documented == built, (
            f"docs/stability.md documents wheels for {sorted(documented)} but release.yml "
            f"builds {sorted(built)}"
        )

    def test_every_built_wheel_is_smoke_tested(self) -> None:
        """No wheel is published without being installed and run first.

        The gap this closes was open from the first release until 1.1.0: two
        cross-compiled targets were built, published, and never executed,
        because no runner of their architecture existed when the pipeline was
        written. Nothing compared the two matrices, so nothing said so.
        """
        untested = sorted(_built_wheel_artifacts() - _smoke_tested_artifacts())
        assert not untested, (
            f"release.yml builds and publishes {untested} without any smoke-test leg "
            f"installing them. Add a matrix entry on a runner of that architecture, or stop "
            f"publishing the wheel."
        )

    def test_the_smoke_test_matrix_installs_only_wheels_that_are_built(self) -> None:
        """The other direction: no leg names an artifact that is never produced.

        A mistyped `wheel-artifact` fails the download at release time, after
        every wheel has already built -- the most expensive moment to find a
        typo. This finds it in the ordinary test run instead.
        """
        phantom = sorted(_smoke_tested_artifacts() - _built_wheel_artifacts())
        assert not phantom, (
            f"the smoke-test matrix downloads {phantom}, which no build job uploads. "
            f"Artifacts actually built: {sorted(_built_wheel_artifacts())}"
        )

    def test_the_documented_smoke_test_column_matches_the_workflow(self) -> None:
        """The "Smoke-tested before publish" column is read from release.yml.

        Same rule as the Python table's CI column: a claim about what was tested
        is worth nothing if the page is the only thing asserting it.
        """
        rows = _markdown_table("Supported platforms")
        claimed = {
            f"wheels-{row[0].lower()}-{row[1].split()[0].lower()}" for row in rows if row[3] == "yes"
        }
        assert claimed == _smoke_tested_artifacts(), (
            f"docs/stability.md claims these wheels are smoke-tested: {sorted(claimed)}; "
            f"release.yml actually smoke-tests: {sorted(_smoke_tested_artifacts())}"
        )


@needs_checkout
class TestPublicSurfaceTable:
    """The surface table names what the package actually exports."""

    def test_every_exported_name_is_documented(self) -> None:
        """A new public name is a new promise and must be written down."""
        text = STABILITY.read_text(encoding="utf-8")
        start = text.find("\n## The public surface\n")
        assert start != -1, "docs/stability.md has no public-surface section"
        section = text[start : text.find("\n## What counts as a breaking change")]

        exported = set(xlsxturbo.__all__)
        # Documented as a family rather than one row each, which is what the
        # errors page is for; the hierarchy itself is tested there.
        exceptions = {name for name in exported if name.endswith("Error")}
        undocumented = sorted(
            name
            for name in exported - exceptions
            if f"`{name}`" not in section and f"`{name}()`" not in section
        )
        assert not undocumented, (
            f"xlsxturbo exports {undocumented}, which the public-surface section of "
            f"docs/stability.md does not mention. Either document the promise or make "
            f"the name private."
        )
        assert exceptions, "no exception classes exported -- the exclusion above is wrong"
        assert "`XlsxTurboError`" in section, "the exception family is not mentioned"


class TestGeneratedFileDeterminism:
    """The one behavioural claim on the page, measured rather than asserted.

    The page tells readers that hashing the output to detect change reports a
    change every time, and that `docProps/core.xml` is the only reason. Both
    halves are checked here: a run that produced a byte-identical file would
    mean the caveat is wrong, and a second differing part would mean the named
    exception is incomplete.
    """

    @staticmethod
    def _write_twice(tmp_path: Path) -> tuple[Path, Path]:
        """Two exports of identical data, separated by a clock second."""
        frame = pd.DataFrame({"n": [1, 2, 3], "s": ["x", "y", "z"]})
        first = tmp_path / "first.xlsx"
        second = tmp_path / "second.xlsx"
        xlsxturbo.df_to_xlsx(frame, str(first), table_style="Medium9", autofit=True)
        # The differing part has one-second resolution, so two writes inside the
        # same second are identical for a reason that has nothing to do with
        # determinism. Without this wait the test passes by accident.
        time.sleep(1.1)
        xlsxturbo.df_to_xlsx(frame, str(second), table_style="Medium9", autofit=True)
        return first, second

    def test_only_the_timestamped_part_differs_between_runs(self, tmp_path: Path) -> None:
        """Every archive member except `docProps/core.xml` is byte-identical."""
        first, second = self._write_twice(tmp_path)
        digests = []
        for path in (first, second):
            with zipfile.ZipFile(path) as archive:
                digests.append(
                    {name: hashlib.sha256(archive.read(name)).hexdigest() for name in archive.namelist()}
                )

        assert list(digests[0]) == list(digests[1]), "the two archives hold different parts"
        differing = {name for name in digests[0] if digests[0][name] != digests[1][name]}
        assert differing == {TIMESTAMPED_PART}, (
            f"docs/stability.md names {TIMESTAMPED_PART} as the only part that differs "
            f"between two runs, but these differ: {sorted(differing)}"
        )

    def test_the_whole_file_is_not_reproducible(self, tmp_path: Path) -> None:
        """The caveat on the page is real, not a precaution.

        Its own control: if this ever passes, the warning telling people not to
        hash the output has become false and should be deleted rather than left
        as folklore.
        """
        first, second = self._write_twice(tmp_path)
        assert first.read_bytes() != second.read_bytes(), (
            "two exports produced byte-identical files, so docs/stability.md's warning "
            "that hashing the output always reports a change is now wrong"
        )
