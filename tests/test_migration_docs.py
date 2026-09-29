"""The migration guide's examples run, and each "after" matches its "before".

``docs/migrating.md`` promises that every example is executed and compared. This is
that check: the page's ``python`` blocks run in order in one namespace, in a scratch
directory, and every ``before_<name>.xlsx`` they write is compared by cell value with
``after_<name>.xlsx``. A migration example that stops producing the same workbook as
the code it replaces fails here instead of misleading a reader.

The "before" examples use XlsxWriter directly and through ``polars.write_excel``, which
is why ``xlsxwriter`` is in ``requirements-test.txt``.
"""

from __future__ import annotations

import math
import os
import re
from pathlib import Path
from typing import Any

import pytest

from tests.helpers import HAS_OPENPYXL, REPO_ROOT, load_workbook, repo_checkout_available

pytestmark = [
    pytest.mark.skipif(not repo_checkout_available(), reason="reads docs/migrating.md from a source checkout"),
    pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification"),
]

PAGE = REPO_ROOT / "docs" / "migrating.md"

# The pairs the page is written to contain. Listed rather than discovered, so an
# example that silently stops writing its file is a failure, not a smaller loop.
EXPECTED_PAIRS = {"plain", "sheets", "index", "formatted", "polars", "openpyxl"}


def python_blocks(text: str) -> list[str]:
    """The fenced ``python`` blocks of a Markdown page, in order.

    Args:
        text: The page source.

    Returns:
        Each block's code.
    """
    return re.findall(r"^```python\n(.*?)^```$", text, re.DOTALL | re.MULTILINE)


def sheet_values(path: Path) -> dict[str, list[list[Any]]]:
    """Every sheet's cell values, keyed by sheet name.

    Args:
        path: The workbook.

    Returns:
        Sheet name to rows of values.
    """
    book = load_workbook(path)
    return {ws.title: [list(row) for row in ws.iter_rows(values_only=True)] for ws in book}


def same_value(after: Any, before: Any) -> bool:
    """Whether two cell values match, allowing for openpyxl's float rounding.

    openpyxl, the default ``to_excel`` engine, writes floats to 16 significant
    digits, so ``0.1 + 0.2`` is stored as ``0.3``; xlsxturbo stores the exact
    double. The page documents that difference, and it is the only one allowed.

    Args:
        after: The value xlsxturbo wrote.
        before: The value the replaced code wrote.

    Returns:
        True when the values are equal, or are floats within 1e-15 of each other.
    """
    if isinstance(after, float) and isinstance(before, float):
        return math.isclose(after, before, rel_tol=1e-15)
    return type(after) is type(before) and after == before


@pytest.fixture(scope="module")
def ran_page(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Execute the page's examples once, in a scratch directory.

    Args:
        tmp_path_factory: pytest's temporary directory factory.

    Returns:
        The directory the examples wrote into.
    """
    workdir = tmp_path_factory.mktemp("migrating")
    blocks = python_blocks(PAGE.read_text(encoding="utf-8"))
    assert len(blocks) >= 2 * len(EXPECTED_PAIRS), f"found only {len(blocks)} python blocks"
    namespace: dict[str, Any] = {"__name__": "migrating_docs"}
    previous = Path.cwd()
    os.chdir(workdir)
    try:
        for index, code in enumerate(blocks):
            try:
                exec(compile(code, f"docs/migrating.md block {index}", "exec"), namespace)  # noqa: S102
            except Exception as error:  # pragma: no cover - reported with the block number
                pytest.fail(f"block {index} of docs/migrating.md raised {error!r}:\n{code}")
    finally:
        os.chdir(previous)
    return workdir


def test_every_expected_pair_was_written(ran_page: Path) -> None:
    """Each named example wrote both its files, and no unlisted pair exists."""
    before = {p.stem.removeprefix("before_") for p in ran_page.glob("before_*.xlsx")}
    after = {p.stem.removeprefix("after_") for p in ran_page.glob("after_*.xlsx")}
    assert before == EXPECTED_PAIRS
    assert after == EXPECTED_PAIRS


@pytest.mark.parametrize("name", sorted(EXPECTED_PAIRS))
def test_after_matches_before(ran_page: Path, name: str) -> None:
    """The xlsxturbo version writes the same cell values as the code it replaces.

    Args:
        ran_page: The directory the examples wrote into.
        name: The example pair.
    """
    after = sheet_values(ran_page / f"after_{name}.xlsx")
    before = sheet_values(ran_page / f"before_{name}.xlsx")
    assert list(after) == list(before), "sheet names or order differ"
    for sheet, rows in before.items():
        assert len(after[sheet]) == len(rows), f"{sheet}: row count differs"
        for row_number, (got, expected) in enumerate(zip(after[sheet], rows, strict=True), start=1):
            assert len(got) == len(expected), f"{sheet} row {row_number}: {got} != {expected}"
            assert all(same_value(a, b) for a, b in zip(got, expected, strict=True)), (
                f"{sheet} row {row_number}: {got} != {expected}"
            )


def display_format(code: str) -> str:
    """A number format reduced to what it displays for a number or date.

    Two spellings differ without changing what a reader sees, and both occur here:
    Excel's date codes are case-insensitive (pandas writes ``YYYY-MM-DD``,
    xlsxturbo ``yyyy-mm-dd``), and polars appends a ``;@`` section that applies
    only to text values, which a date column does not hold. Nothing else is
    normalised.

    Args:
        code: The number format code.

    Returns:
        The code, lower-cased, without a trailing text section.
    """
    return code.lower().removesuffix(";@")


def sheet_formatting(path: Path) -> dict[str, Any]:
    """The formatting a migration has to carry over, for the first sheet.

    Args:
        path: The workbook.

    Returns:
        Number formats of every data cell, header bold and fill per column, the
        freeze pane, and the tables' styles.
    """
    ws = load_workbook(path).worksheets[0]
    rows = list(ws.iter_rows())
    return {
        "number_formats": [[display_format(cell.number_format) for cell in row] for row in rows[1:]],
        "header_bold": [bool(cell.font.b) for cell in rows[0]],
        "header_fill": [cell.fill.fgColor.rgb if cell.fill.fill_type else None for cell in rows[0]],
        "freeze_panes": ws.freeze_panes,
        "tables": sorted(str(t.tableStyleInfo.name) if t.tableStyleInfo else "" for t in ws.tables.values()),
    }


@pytest.mark.parametrize("name", ["formatted", "polars"])
def test_formatting_matches_before(ran_page: Path, name: str) -> None:
    """The formatting examples carry over the formats, not only the values.

    Args:
        ran_page: The directory the examples wrote into.
        name: The example pair.
    """
    after = sheet_formatting(ran_page / f"after_{name}.xlsx")
    before = sheet_formatting(ran_page / f"before_{name}.xlsx")
    for key, expected in before.items():
        assert after[key] == expected, f"{key}: {after[key]} != {expected}"
