"""Table column headers: what the writer escapes, and what is still refused.

rust_xlsxwriter 0.102.0 wrote a table column's name into ``<tableColumn name="...">``
without the escaping Excel uses in that attribute, so a tab, a carriage return, another
control character or a literal ``_x000D_`` gave a workbook Excel refused to open, and
xlsxturbo refused those headers. 0.103.0 escapes them (upstream #216 and #217), and the
fix was opened in Excel on its main branch at ``a94fc21cf``: every table column name equal
to its header cell. ``TestHeadersTheWriterNowEscapes`` pins the escaped forms the release
writes. A name over 255 UTF-16 units stays refused: Excel cuts it, and the crate now raises
its own error, which the guard turns into a ``ConfigurationError`` naming the column. The
measurements are in ``dev-docs/upstream-rust-xlsxwriter.md``.

openpyxl cannot stand in for Excel on this: the refused files are about what Excel accepts,
so the tests pin the refusal and the raw attribute text, not openpyxl's reading of it.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

REFUSED = [
    ("x" * 256, "256 characters"),
    ("\U0001f600" * 128, "256 characters"),
]

# Header -> the raw text of its <tableColumn name="..."> attribute in rust_xlsxwriter
# 0.103.0. A tab and a CR are character references, which XML attribute normalisation
# leaves alone; other control characters and a literal _xHHHH_ use Excel's own escape.
ESCAPED = [
    ("a\tb", "a&#x9;b"),
    ("a\rb", "a&#xD;b"),
    ("a\r\nb", "a&#xD;&#xA;b"),
    ("a\x01b", "a_x0001_b"),
    ("a\x00b", "a_x0000_b"),
    ("a\ufffeb", "a_xFFFE_b"),
    ("_x000D_", "_x005f_x000D_"),
    ("total_x0041_", "total_x005f_x0041_"),
]

ACCEPTED = ["a\nb", "  a  ", "_X000D_", "_x000D", "x" * 255, "\U0001f600" * 127]


def _frame(header: str) -> pd.DataFrame:
    """Two columns, two rows; the first column carries the header under test."""
    return pd.DataFrame({header: [1, 2], "b": [3, 4]})


def _table_xml(path: Path) -> str:
    """The first table part of the workbook."""
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/tables/table1.xml").decode("utf-8")


class TestUnwritableTableHeaderIsRefused:
    """The refusal names the sheet, the column and the reason."""

    @pytest.mark.parametrize(("header", "reason"), REFUSED)
    def test_a_table_with_that_header_raises(self, tmp_path: Path, header: str, reason: str) -> None:
        """Each header Excel refuses in a table raises ConfigurationError."""
        with pytest.raises(xlsxturbo.ConfigurationError) as caught:
            xlsxturbo.df_to_xlsx(_frame(header), str(tmp_path / "t.xlsx"), sheet_name="Data", table_style="Medium2")
        message = str(caught.value)
        assert "sheet 'Data'" in message
        assert "cannot be the header of a table" in message
        assert reason in message
        assert "table_style" in message

    def test_a_long_name_is_cut_in_the_message(self, tmp_path: Path) -> None:
        """A 256-character name is shown as its first 40 characters and a ``...`` marker."""
        with pytest.raises(xlsxturbo.ConfigurationError) as caught:
            xlsxturbo.df_to_xlsx(_frame("x" * 256), str(tmp_path / "t.xlsx"), table_style="Medium2")
        message = str(caught.value)
        assert f'column "{"x" * 40}"... cannot be the header of a table' in message
        assert "x" * 41 not in message
        assert len(message) < 300

    def test_a_later_sheet_is_checked_and_named(self, tmp_path: Path) -> None:
        """In a workbook of several sheets the message names the sheet at fault."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="sheet 'Bad'"):
            xlsxturbo.dfs_to_xlsx(
                [(_frame("a"), "Good"), (_frame("x" * 256), "Bad")],
                str(tmp_path / "t.xlsx"),
                table_style="Medium2",
            )


class TestTheSameHeaderIsFineWhereNoTableIsWritten:
    """The check belongs to the table, not to the column name."""

    @pytest.mark.parametrize(("header", "_reason"), REFUSED)
    def test_without_a_table_style_it_is_written(self, tmp_path: Path, header: str, _reason: str) -> None:
        """No table, no refusal."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(header), str(path))
        assert path.exists()

    def test_with_header_false_no_table_is_made(self, tmp_path: Path) -> None:
        """A sheet without a header row gets no table, so the name is never a table header."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame("x" * 256), str(path), table_style="Medium2", header=False)
        with zipfile.ZipFile(path) as archive:
            assert not [n for n in archive.namelist() if n.startswith("xl/tables/")]

    def test_constant_memory_makes_no_table(self, tmp_path: Path) -> None:
        """constant_memory skips the table, with its usual warning, and does not refuse."""
        path = tmp_path / "t.xlsx"
        with pytest.warns(RuntimeWarning, match="table_style"):
            xlsxturbo.df_to_xlsx(_frame("x" * 256), str(path), table_style="Medium2", constant_memory=True)
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
        assert "xl/worksheets/sheet1.xml" in names
        assert not [n for n in names if n.startswith("xl/tables/")]


#: Every name ``table_style`` accepts, in the families its error message advertises.
TABLE_STYLES = (
    [f"Light{n}" for n in range(1, 22)]
    + [f"Medium{n}" for n in range(1, 29)]
    + [f"Dark{n}" for n in range(1, 12)]
)


class TestEveryTableStyleReachesTheXml:
    """Each style name selects its own style, not merely some style."""

    @pytest.mark.parametrize("style", TABLE_STYLES)
    def test_every_table_style_name_reaches_the_xml(self, tmp_path: Path, style: str) -> None:
        """The table part names ``TableStyle`` plus the name, so no name maps to a neighbour."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame("a"), str(path), table_style=style)
        assert f'<tableStyleInfo name="TableStyle{style}" ' in _table_xml(path)

    def test_the_list_is_the_whole_family(self) -> None:
        """Control: 21 + 28 + 11 names, no duplicate."""
        assert len(set(TABLE_STYLES)) == 60

    def test_style_none_writes_a_table_without_a_style_name(self, tmp_path: Path) -> None:
        """``"None"`` is a table with no style, not a missing table."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame("a"), str(path), table_style="None")
        xml = _table_xml(path)
        assert "<tableStyleInfo " in xml
        assert "TableStyle" not in xml.split("<tableStyleInfo ", 1)[1].split("/>", 1)[0]


class TestHeadersTheWriterNowEscapes:
    """What 0.102.0 wrote raw, 0.103.0 escapes; the guard no longer refuses these headers."""

    @pytest.mark.parametrize(("header", "attribute"), ESCAPED)
    def test_the_table_is_written_with_the_escaped_name(self, tmp_path: Path, header: str, attribute: str) -> None:
        """The table column name is the release's escaped form of the header."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(header), str(path), table_style="Medium2")
        assert f'<tableColumn id="1" name="{attribute}"' in _table_xml(path)


class TestHeadersExcelOpensStillMakeATable:
    """The guard refuses only what was measured to fail."""

    @pytest.mark.parametrize("header", ACCEPTED)
    def test_the_table_is_written(self, tmp_path: Path, header: str) -> None:
        """Each of these opened in Excel as a table header and is still written."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(header), str(path), table_style="Medium2")
        assert '<tableColumn id="1"' in _table_xml(path)

    def test_a_line_feed_is_written_as_a_character_reference(self, tmp_path: Path) -> None:
        """The form Excel opens; if the writer changes it, the guard needs a second look."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame("a\nb"), str(path), table_style="Medium2")
        assert 'name="a&#xA;b"' in _table_xml(path)
