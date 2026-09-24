"""AGENTS.md is an index into dev-docs/, and the two must agree in both directions.

Sections move out of AGENTS.md verbatim to keep the instruction files agents load
on every session under their size budget, and AGENTS.md keeps a one-line entry
per section with a link. Nothing else checks that pairing: a dev-docs heading
renamed or added without its entry, or an entry whose link stopped resolving,
would leave the next agent with a trap it cannot find.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(), reason="audits repository files, absent from an installed wheel"
)

AGENTS = REPO_ROOT / "AGENTS.md"
DEV_DOCS = REPO_ROOT / "dev-docs"
LINK = re.compile(r"\]\((dev-docs/[^)#\s]+\.md)(?:#([^)\s]+))?\)")
SECTION = re.compile(r"^#{2,3} (.+)$", re.MULTILINE)


def github_slug(heading: str) -> str:
    """The anchor GitHub renders for a Markdown heading.

    Lowercase, drop every character that is not a word character, a space or a
    hyphen, then turn each space into a hyphen -- so ``Upstream defects belong
    upstream — file them`` becomes ``upstream-defects-belong-upstream--file-them``,
    with the doubled hyphen where the em dash stood between two spaces.

    Args:
        heading: The heading text without its leading hashes.

    Returns:
        The anchor slug.
    """
    kept = "".join(ch for ch in heading.lower() if ch.isalnum() or ch in " -_")
    return kept.replace(" ", "-")


def links() -> list[tuple[str, str | None]]:
    """Every ``dev-docs/`` link in AGENTS.md, as ``(path, anchor or None)``.

    Returns:
        The links in file order.
    """
    return [(m.group(1), m.group(2)) for m in LINK.finditer(AGENTS.read_text(encoding="utf-8"))]


def sections(path: Path) -> list[str]:
    """The slugs of a dev-docs file's ``##`` and ``###`` headings.

    Args:
        path: A file under ``dev-docs/``.

    Returns:
        The slugs, in file order.
    """
    return [github_slug(h) for h in SECTION.findall(path.read_text(encoding="utf-8"))]


def test_the_index_has_links_to_check() -> None:
    """The emptiness control: a regex that matched nothing would pass every other test."""
    assert len(links()) >= 10
    assert len(list(DEV_DOCS.glob("*.md"))) >= 5


def test_the_slug_matches_githubs_for_the_shapes_in_use() -> None:
    """The anchors in use are spelled the way GitHub renders them."""
    assert github_slug("Upstream defects belong upstream — file them") == (
        "upstream-defects-belong-upstream--file-them"
    )
    assert github_slug("A new CPython needs no code change — and the `abi3` promise has one hole") == (
        "a-new-cpython-needs-no-code-change--and-the-abi3-promise-has-one-hole"
    )
    assert github_slug("Coverage, and why the obvious command lies") == (
        "coverage-and-why-the-obvious-command-lies"
    )


def test_every_link_resolves() -> None:
    """Each linked file exists, and each anchor names one of its headings."""
    broken = []
    for path, anchor in links():
        target = REPO_ROOT / path
        if not target.is_file():
            broken.append(f"{path}: no such file")
        elif anchor is not None and anchor not in sections(target):
            broken.append(f"{path}#{anchor}: no such heading")
    assert not broken, broken


def test_every_dev_docs_section_has_an_index_entry() -> None:
    """Each section is linked from AGENTS.md.

    A section is reached by a link to its anchor, or by a plain link to its file
    when that file holds only one section.
    """
    linked = links()
    anchors = {(path, anchor) for path, anchor in linked if anchor is not None}
    plain = {path for path, anchor in linked if anchor is None}
    unreached = []
    for doc in sorted(DEV_DOCS.glob("*.md")):
        rel = doc.relative_to(REPO_ROOT).as_posix()
        slugs = sections(doc)
        if not slugs:
            unreached.append(f"{rel}: has no ## or ### section")
            continue
        for slug in slugs:
            if (rel, slug) in anchors or (len(slugs) == 1 and rel in plain):
                continue
            unreached.append(f"{rel}#{slug}")
    assert not unreached, f"dev-docs sections with no AGENTS.md entry: {unreached}"
