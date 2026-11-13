"""A column name the table writer cannot express is refused, not written into a damaged file.

rust_xlsxwriter 0.102.0 writes a table column's name into ``<tableColumn name="...">``
without the ``_xHHHH_`` escaping Excel uses in that attribute. A tab, a carriage return,
another control character, a literal ``_x000D_`` or a name over 255 characters therefore
gave a workbook Excel refused to open, with no error from the writer. Each case here was
opened in Excel when the guard was written; the measurement is in
``dev-docs/upstream-rust-xlsxwriter.md``.

openpyxl cannot stand in for Excel on this: the refused files are about what Excel accepts,
so the tests pin the refusal and the three ways around it, not the damaged output.
"""

from __future__ import annotations

import warnings
import zipfile
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

REFUSED = [
    ("a\tb", "a tab"),
    ("a\rb", "a carriage return"),
    ("a\r\nb", "a carriage return"),
    ("a\x01b", "U+0001"),
    ("a\x00b", "U+0000"),
    ("a￾b", "U+FFFE"),
    ("_x000D_", "_xHHHH_"),
    ("total_x0041_", "_xHHHH_"),
    ("x" * 256, "256 characters"),
    ("\U0001f600" * 128, "256 characters"),
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

    def test_the_message_shows_the_control_character_escaped(self, tmp_path: Path) -> None:
        """A tab is shown in its escaped form so the column can be found."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=r'column "Net\\tAmount"'):
            xlsxturbo.df_to_xlsx(_frame("Net\tAmount"), str(tmp_path / "t.xlsx"), table_style="Medium2")

    def test_a_long_name_is_cut_in_the_message(self, tmp_path: Path) -> None:
        """A 256-character name does not fill the message."""
        with pytest.raises(xlsxturbo.ConfigurationError) as caught:
            xlsxturbo.df_to_xlsx(_frame("x" * 256), str(tmp_path / "t.xlsx"), table_style="Medium2")
        assert len(str(caught.value)) < 300

    def test_a_later_sheet_is_checked_and_named(self, tmp_path: Path) -> None:
        """In a workbook of several sheets the message names the sheet at fault."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="sheet 'Bad'"):
            xlsxturbo.dfs_to_xlsx(
                [(_frame("a"), "Good"), (_frame("a\tb"), "Bad")],
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
        xlsxturbo.df_to_xlsx(_frame("a\tb"), str(path), table_style="Medium2", header=False)
        with zipfile.ZipFile(path) as archive:
            assert not [n for n in archive.namelist() if n.startswith("xl/tables/")]

    def test_constant_memory_makes_no_table(self, tmp_path: Path) -> None:
        """constant_memory skips the table, with its usual warning, and does not refuse."""
        path = tmp_path / "t.xlsx"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", RuntimeWarning)
            xlsxturbo.df_to_xlsx(_frame("a\tb"), str(path), table_style="Medium2", constant_memory=True)
        assert path.exists()


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
