"""Tests that public-facing policy documents keep up with the released version.

Every correctness gate in this repository points inward -- at options, types,
errors, workflows. None of them reads `SECURITY.md` or the issue templates, so
those drifted for three minor versions and five releases without a single red
build: the security policy still said "xlsxturbo is pre-1.0", named `0.18.x` as
the supported line, and promised what would happen "once 1.0 ships", while the
package was at 1.1.0. An external reviewer found it, which is the expensive way.

These files are read by people deciding whether to trust the project, and a
stale support table is worse than a missing one -- it states a version is
supported when it is not.

The checks are deliberately narrow. They compare against the *declared* version
and reject phrases that a released project cannot truthfully say. They do not
try to review prose.
"""

from __future__ import annotations

import re

import pytest
import yaml

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="these tests audit repository files, which a wheel install does not carry",
)

SECURITY = REPO_ROOT / "SECURITY.md"
ISSUE_TEMPLATES = REPO_ROOT / ".github" / "ISSUE_TEMPLATE"
PYPROJECT = REPO_ROOT / "pyproject.toml"

# Phrases a 1.x project cannot truthfully make. Matched case-insensitively as
# whole phrases rather than by keyword, so ordinary prose about the 1.0 release
# ("the stability policy published in 1.0.0") does not trip them.
PRE_1_0_PHRASES = (
    "is pre-1.0",
    "once 1.0 ships",
    "before 1.0 ships",
    "until 1.0",
)


def _declared_version() -> str:
    """The version in pyproject.toml, which is the release of record."""
    match = re.search(
        r'^version\s*=\s*"(\d+\.\d+\.\d+)"', PYPROJECT.read_text(encoding="utf-8"), re.MULTILINE
    )
    assert match, "could not read a version from pyproject.toml"
    return match.group(1)


class TestSecurityPolicyMatchesTheRelease:
    """`SECURITY.md` describes the version that is actually shipping."""

    def test_no_pre_1_0_language_survives(self) -> None:
        """The policy does not describe the project as unreleased."""
        text = SECURITY.read_text(encoding="utf-8").lower()
        found = [phrase for phrase in PRE_1_0_PHRASES if phrase in text]
        assert not found, (
            f"SECURITY.md still says {found}, which stopped being true at 1.0.0. This file "
            f"is read by people deciding whether to trust the project."
        )

    def test_the_current_minor_line_is_listed_as_supported(self) -> None:
        """The shipping line appears in the support table.

        Matched as `<major>.<minor>.x` rather than the exact version, because a
        support table names lines and a patch release must not force an edit
        here -- a check that demands pointless edits is a check people learn to
        silence.
        """
        major, minor, _ = _declared_version().split(".")
        line = f"{major}.{minor}.x"
        assert line in SECURITY.read_text(encoding="utf-8"), (
            f"SECURITY.md does not mention {line}, the line pyproject.toml declares. Its "
            f"support table is describing versions that are no longer current."
        )

    def test_no_superseded_major_is_still_claimed_as_supported(self) -> None:
        """No pre-1.0 line is listed as supported.

        The failure this catches is specific: a table row whose version column
        is `0.x` and whose support column says yes. Mentioning old versions as
        *unsupported* is exactly what the table is for, so the check reads the
        support column rather than searching for version strings.
        """
        offenders: list[str] = []
        for row in SECURITY.read_text(encoding="utf-8").splitlines():
            if not row.strip().startswith("|"):
                continue
            cells = [cell.strip() for cell in row.strip().strip("|").split("|")]
            if len(cells) < 2:
                continue
            version, support = cells[0], cells[1].lower()
            if re.match(r"^0\.", version) and support.startswith("yes"):
                offenders.append(row.strip())
        assert not offenders, f"SECURITY.md still claims support for a 0.x line: {offenders}"


def _xlsxturbo_version_fields() -> list[tuple[str, str, str]]:
    """`(template, field id, placeholder)` for each xlsxturbo-version field.

    Parsed from the issue-form YAML rather than matched line by line. A line
    regex cannot tell an xlsxturbo version from a Python one, and it reads only
    the first line of a block placeholder -- both of which cost real coverage
    here: the Python-version field inflated the population enough that the
    control below could not fail, and a multi-line `Measurements` placeholder
    naming `0.17.2` and `0.18.0` went unseen entirely.
    """
    fields: list[tuple[str, str, str]] = []
    for template in sorted(ISSUE_TEMPLATES.glob("*.yml")):
        document = yaml.safe_load(template.read_text(encoding="utf-8"))
        if not isinstance(document, dict):
            continue
        for field in document.get("body", ()):
            identifier = str(field.get("id", ""))
            placeholder = (field.get("attributes") or {}).get("placeholder")
            if not placeholder or "version" not in identifier or "python" in identifier:
                continue
            fields.append((template.name, identifier, str(placeholder)))
    return fields


class TestIssueTemplatesSuggestACurrentVersion:
    """A version placeholder is an example, and a stale one misleads."""

    def test_no_version_field_suggests_a_pre_1_0_release(self) -> None:
        """The xlsxturbo-version fields name a release that still exists."""
        offenders = [
            f"{template}:{identifier} suggests {placeholder}"
            for template, identifier, placeholder in _xlsxturbo_version_fields()
            if placeholder.startswith("0.")
        ]
        assert not offenders, (
            f"issue templates suggest pre-1.0 versions, which tells a reporter the project "
            f"is on a line it abandoned: {offenders}"
        )

    def test_no_placeholder_anywhere_names_a_pre_1_0_release(self) -> None:
        """Wider net: no example text mentions a 0.x version.

        Separate from the field check because the miss it caught was not in a
        version field at all -- the `Measurements` placeholder showed a worked
        example comparing `0.17.2` against `0.18.0`, which reads as the versions
        worth benchmarking. Scans every placeholder as whole text, so a block
        scalar is covered.
        """
        offenders: list[str] = []
        for template in sorted(ISSUE_TEMPLATES.glob("*.yml")):
            document = yaml.safe_load(template.read_text(encoding="utf-8"))
            if not isinstance(document, dict):
                continue
            for field in document.get("body", ()):
                placeholder = (field.get("attributes") or {}).get("placeholder")
                if not placeholder:
                    continue
                stale = sorted(set(re.findall(r"\b0\.\d+\.\d+\b", str(placeholder))))
                if stale:
                    offenders.append(f"{template.name}:{field.get('id')} mentions {stale}")
        assert not offenders, f"pre-1.0 versions still shown as examples: {offenders}"

    def test_the_scan_actually_reads_version_fields(self) -> None:
        """Control: the checks above see the fields they claim to check.

        Compares the templates that *ask* for an xlsxturbo version against the
        templates the parser can actually read one from. Renaming the field, or
        breaking the YAML, empties one side and fails here rather than passing
        on nothing.
        """
        found = {template for template, _, _ in _xlsxturbo_version_fields()}
        asking = {
            template.name
            for template in sorted(ISSUE_TEMPLATES.glob("*.yml"))
            if re.search(
                r"^\s*label:.*\bversion\b",
                template.read_text(encoding="utf-8"),
                re.IGNORECASE | re.MULTILINE,
            )
        }
        assert asking, "no issue template asks for a version at all"
        assert found == asking, (
            f"templates asking for a version: {sorted(asking)}; templates the scan can read "
            f"an xlsxturbo version from: {sorted(found)}. The difference goes unchecked."
        )
