"""Guards on the MkDocs site configuration.

Two failure modes are worth catching mechanically rather than at review time:

- A page added under ``docs/`` and never linked from the nav is invisible. MkDocs'
  ``--strict`` catches that in CI, but only once the docs workflow runs; this
  catches it in the ordinary test run.
- A file that is deliberately untracked -- an internal planning memo, review
  notes -- is still published by ``mkdocs build``, because MkDocs publishes the
  whole of ``docs/`` and knows nothing about git. The deploy itself is safe (the
  workflow builds from a clean checkout, which has no untracked files), but a
  local build or a stray ``mkdocs gh-deploy`` would leak them. That was verified
  by removing one ``exclude_docs`` entry: the memo appeared in ``site/`` and
  ``--strict`` still exited 0, so strict mode alone does not cover it.

Every test here asserts the population it examined is non-empty. A structural
check that matches nothing reports the same "no problems" as a clean one.
"""

from __future__ import annotations

from typing import Any

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

# Skip before importing yaml, not after. This module audits repository files, so
# outside a checkout it has nothing to test -- and `yaml` is a test-only
# dependency the release smoke-test environment does not install either, so an
# import above this guard raises ImportError and makes the skip unreachable.
# That is exactly what happened on the first v0.19.0 release attempt: three
# platforms reported `ModuleNotFoundError: No module named 'yaml'`, which hid
# the fact that ten more tests in another module were failing for the same
# underlying reason.
if not repo_checkout_available():  # pragma: no cover - only true outside a checkout
    pytest.skip(
        "docs-site tests audit repository files, which a wheel install does not carry",
        allow_module_level=True,
    )

import yaml  # deliberately below the checkout guard above -- do not hoist

MKDOCS_YML = REPO_ROOT / "mkdocs.yml"
DOCS_DIR = REPO_ROOT / "docs"
GITIGNORE = REPO_ROOT / ".gitignore"


def _load_config() -> dict[str, Any]:
    """Parse mkdocs.yml, tolerating the theme's custom YAML tags."""
    text = MKDOCS_YML.read_text(encoding="utf-8")
    loaded = yaml.safe_load(text)
    assert isinstance(loaded, dict), "mkdocs.yml did not parse to a mapping"
    return loaded


def _nav_targets(nav: object) -> list[str]:
    """Flatten the nav tree to the list of document paths it references."""
    found: list[str] = []
    if isinstance(nav, str):
        found.append(nav)
    elif isinstance(nav, list):
        for item in nav:
            found.extend(_nav_targets(item))
    elif isinstance(nav, dict):
        for value in nav.values():  # pyright: ignore[reportUnknownVariableType]
            found.extend(_nav_targets(value))
    return found


def _excluded() -> list[str]:
    """Entries of mkdocs.yml's `exclude_docs` block, one per line."""
    raw = _load_config().get("exclude_docs", "")
    assert isinstance(raw, str), "exclude_docs should be a block scalar"
    return [line.strip() for line in raw.splitlines() if line.strip()]


def _gitignored_docs_paths() -> list[str]:
    """`.gitignore` entries under docs/, as paths relative to docs/."""
    entries: list[str] = []
    for line in GITIGNORE.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped.startswith("docs/"):
            entries.append(stripped[len("docs/"):])
    return entries


def _docs_pages_on_disk() -> list[str]:
    """Every Markdown file under docs/, relative to docs/, POSIX-separated.

    Deliberately the filesystem rather than `git ls-files`: MkDocs publishes what
    is on disk, and it is indifferent to whether git knows about it. An earlier
    version of this helper asked git, which made the orphan check below examine
    almost nothing while every page was still uncommitted -- it passed for the
    same reason an empty audit passes.
    """
    return sorted(
        p.relative_to(DOCS_DIR).as_posix() for p in DOCS_DIR.rglob("*.md")
    )


def _is_excluded(page: str, excluded: list[str]) -> bool:
    """Whether `exclude_docs` covers this page, by name or by directory."""
    for entry in excluded:
        if entry.endswith("/"):
            if page.startswith(entry):
                return True
        elif page == entry:
            return True
    return False


class TestDocsSite:
    """The site config, the docs tree, and .gitignore must agree."""

    def test_config_parses_and_declares_a_nav(self) -> None:
        """mkdocs.yml is valid YAML with the keys the workflow relies on."""
        config = _load_config()
        for key in ("site_name", "site_url", "theme", "nav", "exclude_docs"):
            assert key in config, f"mkdocs.yml is missing `{key}`"

    def test_every_nav_target_exists(self) -> None:
        """A nav entry pointing at a missing file breaks the build."""
        targets = _nav_targets(_load_config()["nav"])
        assert len(targets) >= 15, f"nav references only {len(targets)} page(s)"
        missing = [t for t in targets if not (DOCS_DIR / t).is_file()]
        assert not missing, f"nav references missing file(s): {missing}"

    def test_nav_has_no_duplicate_targets(self) -> None:
        """The same page listed twice is a copy-paste slip, not a feature."""
        targets = _nav_targets(_load_config()["nav"])
        duplicates = sorted({t for t in targets if targets.count(t) > 1})
        assert not duplicates, f"page(s) appear more than once in the nav: {duplicates}"

    def test_every_page_on_disk_is_in_the_nav_or_excluded(self) -> None:
        """Every page under docs/ must be reachable or deliberately excluded.

        Accounting for all of them, not just the committed ones, is what makes
        this check meaningful: a page is published because it is on disk.
        """
        pages = _docs_pages_on_disk()
        assert len(pages) >= 15, (
            f"only {len(pages)} page(s) found under docs/ -- the site is smaller "
            "than the nav claims, or this is not a full checkout"
        )
        targets = set(_nav_targets(_load_config()["nav"]))
        excluded = _excluded()
        orphans = [
            p for p in pages if p not in targets and not _is_excluded(p, excluded)
        ]
        assert not orphans, (
            f"{len(orphans)} page(s) are neither in the nav nor excluded, so they "
            f"would be published unreachable: {orphans}"
        )

    def test_gitignored_docs_are_excluded_from_the_site(self) -> None:
        """A file too private to commit must not be published by a local build.

        This is the leak guard. MkDocs publishes everything under docs/, so an
        untracked file present in a working copy is published unless named in
        `exclude_docs`.
        """
        ignored = _gitignored_docs_paths()
        assert ignored, (
            ".gitignore lists nothing under docs/ -- either the entries moved or "
            "this check is no longer looking in the right place"
        )
        excluded = _excluded()
        unguarded = [entry for entry in ignored if entry not in excluded]
        assert not unguarded, (
            f"{len(unguarded)} gitignored docs path(s) are absent from mkdocs.yml's "
            f"exclude_docs, so `mkdocs build` would publish them: {unguarded}"
        )

    def test_exclude_docs_entries_are_not_stale(self) -> None:
        """Every exclusion must still refer to something real.

        An entry is legitimate if the path exists (a tracked file kept off the
        site, such as the roadmap) or if it is gitignored (a private file that
        may be absent on this machine). One that is neither is dead weight, and
        dead weight in a safety list is how the list stops being read.
        """
        excluded = _excluded()
        assert excluded, "exclude_docs is empty; the leak guard is not configured"
        ignored = set(_gitignored_docs_paths())
        stale = [
            entry
            for entry in excluded
            if entry not in ignored and not (DOCS_DIR / entry.rstrip("/")).exists()
        ]
        assert not stale, (
            f"exclude_docs names {len(stale)} path(s) that neither exist nor are "
            f"gitignored: {stale}"
        )


API_REFERENCE = DOCS_DIR / "api-reference.md"

# The DataFrame entry points, and the section heading each is documented under.
# Both return a shape, not `None`: `df_to_xlsx` gives `(rows, cols)` and
# `dfs_to_xlsx` a list of one such pair per sheet, the header row counted when
# `header=True`. The page said `None` for both -- a caller reading it would have
# thrown away the only report the library makes of what it wrote.
DATAFRAME_ENTRY_POINTS = ("df_to_xlsx", "dfs_to_xlsx")

_COUNT_WORDS = {
    2: "Two", 3: "Three", 4: "Four", 5: "Five", 6: "Six",
    7: "Seven", 8: "Eight", 9: "Nine", 10: "Ten",
}


def _api_reference_sections() -> dict[str, str]:
    """`api-reference.md` split into its `##` sections, keyed by heading text."""
    sections: dict[str, str] = {}
    heading, body = "", []
    for line in API_REFERENCE.read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            if heading:
                sections[heading] = "\n".join(body)
            heading, body = line[3:].strip(), []
        else:
            body.append(line)
    if heading:
        sections[heading] = "\n".join(body)
    return sections


def _section_for(name: str) -> str:
    """The body of the section whose heading mentions `name`."""
    matches = [
        body for heading, body in _api_reference_sections().items() if name in heading
    ]
    assert len(matches) == 1, (
        f"expected exactly one `##` section in api-reference.md mentioning "
        f"`{name}`, found {len(matches)}"
    )
    return matches[0]


def _returns_bullet(name: str) -> str:
    """The `- **returns**` bullet from that function's section."""
    lines = [
        line.strip()
        for line in _section_for(name).splitlines()
        if line.strip().startswith("- **returns**")
    ]
    assert len(lines) == 1, (
        f"expected exactly one `- **returns**` bullet in api-reference.md's "
        f"`{name}` section, found {len(lines)}"
    )
    return lines[0]


def _exported_error_names() -> set[str]:
    """Every exception class the package exports, from `xlsxturbo.__all__`.

    Derived rather than listed, because a hardcoded copy here would agree with a
    stale page forever -- which is the failure this class exists to catch.
    """
    import xlsxturbo

    return {name for name in xlsxturbo.__all__ if name.endswith("Error")}


class TestApiReferenceMatchesTheLibrary:
    """`docs/api-reference.md` calls itself authoritative; nothing read it.

    It said both DataFrame entry points return `None` (they return a shape),
    and it drew an exception tree of "Six classes" that omitted `OptionError`
    -- the one class most callers should be catching, since it is what makes
    `except OptionError` mean "anything wrong with what I passed".

    `tests/test_stability_policy.py` deliberately exempts exception names from
    its public-surface check, so this page was the one part of the documented
    surface with no mechanical reader at all.

    Every assertion below is derived from the package or from the page's own
    structure, and each has a control asserting it examined something.
    """

    def test_the_page_parses_into_the_sections_this_class_reads(self) -> None:
        """Control: a parser matching nothing passes every check below."""
        sections = _api_reference_sections()
        assert len(sections) >= 5, (
            f"api-reference.md parsed into {len(sections)} `##` section(s); the "
            "checks below read named sections, so a page that stopped parsing "
            "would pass them all while reading nothing"
        )
        assert "Exceptions" in sections, (
            f"api-reference.md has no `## Exceptions` section; found "
            f"{sorted(sections)}"
        )
        for name in DATAFRAME_ENTRY_POINTS:
            assert _returns_bullet(name), f"no returns bullet found for `{name}`"

    def test_the_dataframe_entry_points_do_not_claim_to_return_none(self) -> None:
        """Both return a `(rows, cols)` shape, and the page must say so."""
        wrong = [
            f"{name}: {_returns_bullet(name)}"
            for name in DATAFRAME_ENTRY_POINTS
            if "`None`" in _returns_bullet(name)
        ]
        assert not wrong, (
            f"api-reference.md documents {len(wrong)} DataFrame entry point(s) as "
            f"returning `None`. They return the rows and columns written -- "
            f"`tuple[int, int]` and `list[tuple[int, int]]`: {wrong}"
        )

    def test_the_exception_section_names_every_exported_error_class(self) -> None:
        """A class missing from the tree is a class nobody knows to catch."""
        exported = _exported_error_names()
        assert len(exported) >= 5, (
            f"only {len(exported)} exception name(s) derived from xlsxturbo.__all__ "
            f"({sorted(exported)}); the comparison below would be near-vacuous"
        )
        section = _section_for("Exceptions")
        missing = sorted(name for name in exported if name not in section)
        assert not missing, (
            f"api-reference.md's Exceptions section does not name {missing}, which "
            f"xlsxturbo.__all__ exports. Every exported class is part of the "
            f"stability promise and has to be catchable from the documentation."
        )

    def test_the_stated_class_count_matches_what_is_exported(self) -> None:
        """The page opens with a count, and a count is a claim.

        It read "Six classes" against seven exported ones for two releases. A
        number written in prose beside a list is the part of a page that rots
        first, because adding to the list does not touch it.
        """
        expected = _COUNT_WORDS[len(_exported_error_names())]
        section = _section_for("Exceptions")
        stated = [
            line.strip()
            for line in section.splitlines()
            if line.strip().endswith("classes, all exported from `xlsxturbo`:")
        ]
        assert len(stated) == 1, (
            f"expected one '<N> classes, all exported from `xlsxturbo`:' line in "
            f"api-reference.md's Exceptions section, found {stated}"
        )
        assert stated[0].startswith(expected), (
            f"api-reference.md says {stated[0]!r} but xlsxturbo exports "
            f"{len(_exported_error_names())} exception classes "
            f"({sorted(_exported_error_names())}), so it should read "
            f"{expected!r}"
        )
