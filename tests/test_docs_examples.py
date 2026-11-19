"""Every ``python`` example in the published docs runs.

A page's blocks run in order in one namespace, in a scratch directory, the way a
reader would paste them one after the other. A block that raises fails the page's
test, unless the page itself says it raises: a trailing ``# SomeError: message``
comment, or a ``>>>`` line followed by ``xlsxturbo.SomeError: message``. Those are
checked the other way round: the documented exception has to be the one raised.

This exists because a wrong example passes every other gate. ``docs/errors.md`` once
showed ``column_widths={"A": 12}`` inside ``except ValueError`` as a handler that
"still catches exactly what it did"; the call raises a ``TypeError`` subclass, so the
handler shown did not catch it, and the page stayed that way until someone ran it.

``docs/migrating.md`` is not here: ``tests/test_migration_docs.py`` runs it and also
compares what its examples write.

What this cannot see: prose, option lists and tables. It runs code, it does not read.
"""

from __future__ import annotations

import base64
import re
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import REPO_ROOT, TINY_PNG_B64, repo_checkout_available

pytestmark = pytest.mark.skipif(not repo_checkout_available(), reason="reads docs/*.md from a source checkout")

DOCS = REPO_ROOT / "docs"

# Pages with their own, stricter test.
OWN_TEST = {"migrating.md"}

# Blocks that cannot run here, each found by a line it contains, with the reason.
# A line that matches no block, or more than one, fails `test_the_tables_are_current`.
NOT_RUN = {
    "from fastapi import Response": "needs FastAPI and a running app",
    'return {"error": f"bad export options: {exc}"}': "the body of a handler: `return` outside a function",
}

# Blocks that raise by design without saying so in a comment: the example shows a
# handler that inspects the error and re-raises it.
RAISES_BY_DESIGN = {
    'xlsxturbo.df_to_xlsx(df, "/mnt/reports/out.xlsx")': xlsxturbo.FileError,
}

# Files the examples read. `logo.png` is deliberately not among them: the errors
# page shows what a missing image raises.
CSV_INPUTS = ("input.csv", "data.csv", "big_data.csv", "us_data.csv", "eu_data.csv", "sales.csv")
IMAGE_INPUTS = ("images/widget_a.png", "images/widget_b.png")

# How many characters of a documented message have to open the real one. Pages
# shorten long messages and quote the POSIX wording of an OS error, so the whole
# text cannot be compared; the start names the option and the key, which is the part
# a reader acts on.
MESSAGE_PREFIX = 24

DOCUMENTED_ERROR = re.compile(r"^# (\w+(?:Error|Warning)): (.+)$", re.MULTILINE)
DOCTEST_ERROR = re.compile(r"^xlsxturbo\.(\w+): (.+)$", re.MULTILINE)


def python_blocks(text: str) -> list[str]:
    """The fenced ``python`` blocks of a Markdown page, in order.

    Args:
        text: The page source.

    Returns:
        Each block's code.
    """
    return re.findall(r"^```python[^\n]*\n(.*?)^```", text, re.DOTALL | re.MULTILINE)


def pages() -> list[Path]:
    """The published pages that hold at least one ``python`` block."""
    return [
        page
        for page in sorted(DOCS.glob("*.md"))
        if page.name not in OWN_TEST and python_blocks(page.read_text(encoding="utf-8"))
    ]


def seeded_namespace() -> dict[str, Any]:
    """The names examples use without defining them.

    A page introduces ``df`` or ``path`` once in prose and then shows fragments, so
    these stand in for "the frame you already have". `test_the_tables_are_current`
    fails when a name here is no longer used by any page.

    Returns:
        A fresh namespace for one page.
    """
    frame = pd.DataFrame({"name": ["A", "B"], "amount": [1.5, 2.5]})
    sales = pd.DataFrame({"Region": ["North", "South"], "Volume": [10, 20]})
    costs = pd.DataFrame({"Region": ["North", "South"], "Cost": [3.0, 4.0]})
    return {
        "__name__": "docs_examples",
        "xlsxturbo": xlsxturbo,
        "df_to_xlsx": xlsxturbo.df_to_xlsx,
        "pd": pd,
        "pl": pl,
        "df": frame,
        "path": "example.xlsx",
        "q1": frame,
        "q2": frame,
        "summary": frame,
        "detail": frame,
        "inputs": frame,
        "results": frame,
        "legend": frame,
        "sales": sales,
        "costs": costs,
        "jobs": [SimpleNamespace(frame=frame, path="job1.xlsx"), SimpleNamespace(frame=frame, path="job2.xlsx")],
    }


# Seeded names that every page may rely on; the others must each be used somewhere.
ALWAYS_SEEDED = {"__name__", "xlsxturbo", "pd", "pl"}


def matching(table: dict[str, Any], code: str) -> list[str]:
    """The keys of ``table`` that occur in ``code``."""
    return [line for line in table if line in code]


def run_block(code: str, namespace: dict[str, Any], where: str) -> None:
    """Run one block, holding it to what the page says about it.

    Args:
        code: The block.
        namespace: The page's namespace, shared by its blocks.
        where: Page and block number, for the failure message.
    """
    if matching(NOT_RUN, code):
        return

    doctest_lines = [line[4:] for line in code.splitlines() if line.startswith(">>> ")]
    if doctest_lines:
        documented = DOCTEST_ERROR.search(code)
        assert documented, f"{where}: a `>>>` example without the error it shows:\n{code}"
        expected = getattr(xlsxturbo, documented.group(1))
        with pytest.raises(expected) as caught:
            exec(compile("\n".join(doctest_lines), where, "exec"), namespace)  # noqa: S102
        assert type(caught.value) is expected, f"{where}: raised {type(caught.value).__name__}"
        assert str(caught.value) == documented.group(2), f"{where}: the page shows another message"
        return

    by_design = [RAISES_BY_DESIGN[line] for line in matching(RAISES_BY_DESIGN, code)]
    documented = DOCUMENTED_ERROR.search(code)
    if by_design or documented:
        expected = by_design[0] if by_design else getattr(xlsxturbo, documented.group(1))  # type: ignore[union-attr]
        with pytest.raises(expected) as caught:
            exec(compile(code, where, "exec"), namespace)  # noqa: S102
        assert type(caught.value) is expected, f"{where}: raised {type(caught.value).__name__}"
        if documented:
            shown = documented.group(2)[:MESSAGE_PREFIX]
            assert str(caught.value).startswith(shown), f"{where}: the page shows {shown!r}, got {caught.value}"
        return

    try:
        exec(compile(code, where, "exec"), namespace)  # noqa: S102
    except Exception as error:
        pytest.fail(f"{where} raised {type(error).__name__}: {error}\n{code}")


@pytest.mark.parametrize("page", pages(), ids=lambda page: page.name)
def test_the_examples_on_the_page_run(page: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """A page's examples run in order, and its documented errors are the real ones.

    Args:
        page: The docs page.
        tmp_path: The scratch directory the examples write into.
        monkeypatch: Used to make that directory the working directory.
    """
    monkeypatch.chdir(tmp_path)
    for name in CSV_INPUTS:
        (tmp_path / name).write_text("name,when,amount\nA,01-02-2024,1.5\nB,03-04-2024,2.5\n", encoding="utf-8")
    for name in IMAGE_INPUTS:
        (tmp_path / name).parent.mkdir(exist_ok=True)
        (tmp_path / name).write_bytes(base64.b64decode(TINY_PNG_B64))

    namespace = seeded_namespace()
    for index, code in enumerate(python_blocks(page.read_text(encoding="utf-8"))):
        run_block(code, namespace, f"docs/{page.name} block {index}")


def test_the_tables_are_current() -> None:
    """The lists above still describe the pages, and the pages were found at all."""
    found = pages()
    blocks = [code for page in found for code in python_blocks(page.read_text(encoding="utf-8"))]
    # 66 blocks on 18 pages when this was written. A regex that stops matching the
    # fences would otherwise leave a green suite that runs nothing.
    assert len(found) >= 15, [page.name for page in found]
    assert len(blocks) >= 60, len(blocks)

    for table in (NOT_RUN, RAISES_BY_DESIGN):
        for line in table:
            hits = sum(line in code for code in blocks)
            assert hits == 1, f"{line!r} is in {hits} blocks, expected exactly one"

    everything = "\n".join(blocks)
    for name in set(seeded_namespace()) - ALWAYS_SEEDED:
        assert re.search(rf"\b{re.escape(name)}\b", everything), f"no example uses the seeded name {name!r}"
    for name in (*CSV_INPUTS, *IMAGE_INPUTS):
        assert name in everything, f"no example reads {name}"
