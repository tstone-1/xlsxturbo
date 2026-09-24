"""Dates keep their date number format when a cell or column also has a format.

Until 1.6.0 a caller's format replaced the date format instead of adding to it:
``column_formats={"d": {"bold": True}}`` on a date column wrote every date as a
bare serial (``46289`` for 2026-09-24) with the ``General`` number format, and
nothing warned. An explicit ``num_format`` has always been kept, and still is.
"""

from __future__ import annotations

import datetime as dt
from typing import Any

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")

DATE = dt.date(2026, 9, 24)
DATETIME = dt.datetime(2026, 9, 24, 13, 45, 30)
DATE_FORMAT = "yyyy-mm-dd"
DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"

FRAMES: list[Any] = [
    pytest.param(pl.DataFrame({"d": [DATE]}), DATE_FORMAT, id="polars_date"),
    pytest.param(pl.DataFrame({"d": [DATETIME]}), DATETIME_FORMAT, id="polars_datetime"),
    pytest.param(pd.DataFrame({"d": pd.to_datetime([DATETIME])}), DATETIME_FORMAT, id="pandas_timestamp"),
    pytest.param(pd.DataFrame({"d": [DATE]}), DATE_FORMAT, id="pandas_object_date"),
]


def _cell(path: str, ref: str) -> tuple[Any, str, bool]:
    """Read one cell back: its value, number format and whether it is bold.

    Args:
        path: The workbook to read.
        ref: The cell reference, e.g. ``"A2"``.

    Returns:
        The value openpyxl reads, the cell's number format, and its bold flag.
    """
    wb = load_workbook(path)
    cell = active_ws(wb)[ref]
    result = (cell.value, cell.number_format, bool(cell.font.b))
    wb.close()
    return result


class TestColumnFormatsOnDateColumns:
    """``column_formats`` on a date or datetime column."""

    @pytest.mark.parametrize(("df", "expected_format"), FRAMES)
    def test_no_format_is_the_control(self, tmp_xlsx: str, df: Any, expected_format: str) -> None:
        """Without a column format the default date format applies.

        Args:
            tmp_xlsx: Output path fixture.
            df: A one-column frame holding a date or datetime.
            expected_format: The default number format for that type.
        """
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        value, number_format, _ = _cell(tmp_xlsx, "A2")
        assert isinstance(value, dt.datetime)
        assert number_format == expected_format

    @pytest.mark.parametrize(("df", "expected_format"), FRAMES)
    @pytest.mark.parametrize(
        "column_format",
        [{"bold": True}, {"bold": True, "bg_color": "#FFFF00"}, {"bold": True, "font_name": "Arial"}],
        ids=["bold", "bold_fill", "bold_font"],
    )
    def test_a_format_without_num_format_keeps_the_date_readable(
        self, tmp_xlsx: str, df: Any, expected_format: str, column_format: dict[str, Any]
    ) -> None:
        """The caller's styling is applied and the date format is kept beside it.

        Args:
            tmp_xlsx: Output path fixture.
            df: A one-column frame holding a date or datetime.
            expected_format: The default number format for that type.
            column_format: A format that sets no ``num_format``.
        """
        formats: Any = {"d": column_format}
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats=formats)
        value, number_format, bold = _cell(tmp_xlsx, "A2")
        assert isinstance(value, dt.datetime), f"read back {value!r}, a serial rather than a date"
        assert number_format == expected_format
        assert bold

    @pytest.mark.parametrize(("df", "expected_format"), FRAMES)
    def test_an_explicit_num_format_still_wins(self, tmp_xlsx: str, df: Any, expected_format: str) -> None:
        """The caller's number format is used rather than the default.

        Args:
            tmp_xlsx: Output path fixture.
            df: A one-column frame holding a date or datetime.
            expected_format: Unused; the default this case must not produce.
        """
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={"d": {"bold": True, "num_format": "dd.mm.yyyy"}})
        value, number_format, bold = _cell(tmp_xlsx, "A2")
        assert isinstance(value, dt.datetime)
        assert number_format == "dd.mm.yyyy" != expected_format
        assert bold

    def test_non_date_cells_in_the_same_column_keep_the_callers_format(self, tmp_xlsx: str) -> None:
        """Only date values get the date number format; a number stays a number.

        A pandas object column can mix a date with other values, and the merged
        date variant must not leak onto them.

        Args:
            tmp_xlsx: Output path fixture.
        """
        df = pd.DataFrame({"d": [DATE, 42]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={"d": {"bold": True}})
        assert _cell(tmp_xlsx, "A2")[1] == DATE_FORMAT
        value, number_format, bold = _cell(tmp_xlsx, "A3")
        assert (value, number_format, bold) == (42, "General", True)

    def test_dfs_to_xlsx_per_sheet_column_formats(self, tmp_xlsx: str) -> None:
        """The multi-sheet path shares the writer and gets the same behaviour.

        Args:
            tmp_xlsx: Output path fixture.
        """
        df = pl.DataFrame({"d": [DATE]})
        xlsxturbo.dfs_to_xlsx([(df, "S1", {"column_formats": {"d": {"bold": True}}})], tmp_xlsx)
        value, number_format, bold = _cell(tmp_xlsx, "A2")
        assert isinstance(value, dt.datetime)
        assert (number_format, bold) == (DATE_FORMAT, True)


class TestCellsWithDates:
    """``cells`` entries whose value is a date or datetime."""

    @pytest.mark.parametrize(
        ("value", "expected_format"),
        [(DATE, DATE_FORMAT), (DATETIME, DATETIME_FORMAT)],
        ids=["date", "datetime"],
    )
    @pytest.mark.parametrize(
        "options",
        [
            {},
            {"format": {"bold": True}},
            {"font_name": "Arial"},
            {"wrap_text": True},
            {"format": {"bold": True}, "font_name": "Arial"},
        ],
        ids=["plain", "format", "font_shorthand", "wrap_shorthand", "format_and_shorthand"],
    )
    def test_the_date_format_survives_any_styling(
        self, tmp_xlsx: str, value: dt.date, expected_format: str, options: dict[str, Any]
    ) -> None:
        """Every way of styling a cell keeps a date readable.

        Args:
            tmp_xlsx: Output path fixture.
            value: The date or datetime written.
            expected_format: The default number format for that type.
            options: The styling keys beside ``value``.
        """
        cell: Any = {"value": value, **options}
        xlsxturbo.df_to_xlsx(pl.DataFrame({"a": [1]}), tmp_xlsx, cells={"C1": cell})
        read, number_format, _ = _cell(tmp_xlsx, "C1")
        assert isinstance(read, dt.datetime), f"read back {read!r}, a serial rather than a date"
        assert number_format == expected_format

    def test_an_explicit_num_format_still_wins(self, tmp_xlsx: str) -> None:
        """A shorthand ``num_format`` overrides the date default, as before.

        Args:
            tmp_xlsx: Output path fixture.
        """
        cell: Any = {"value": DATE, "format": {"bold": True}, "num_format": "d mmm yyyy"}
        xlsxturbo.df_to_xlsx(pl.DataFrame({"a": [1]}), tmp_xlsx, cells={"C1": cell})
        _, number_format, bold = _cell(tmp_xlsx, "C1")
        assert (number_format, bold) == ("d mmm yyyy", True)
