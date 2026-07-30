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

from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
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
