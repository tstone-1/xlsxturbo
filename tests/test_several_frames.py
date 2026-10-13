"""Tests for several frames on one sheet: ``dfs_to_xlsx`` entries that name the same sheet.

Each entry writes its frame at its own ``startrow``. The rules a second frame
brings with it are the subject here: frames may not overlap, the first entry for
a sheet decides the options that describe the whole sheet, an option addressed
by cell reference is applied once and not once per frame, and under
``constant_memory`` the frames are listed from top to bottom.
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

import pandas as pd
import pytest
import xlsxturbo
from xlsxturbo.types import SheetOptions

from tests.helpers import HAS_OPENPYXL, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


def _summary() -> pd.DataFrame:
    """The upper frame: two columns, two rows."""
    return pd.DataFrame({"Item": ["rows checked", "rows flagged"], "Value": ["120", "7"]})


def _reasons() -> pd.DataFrame:
    """The lower frame: two columns, two rows."""
    return pd.DataFrame({"Reason": ["missing", "outdated"], "Count": [4, 3]})


def _rows(ws: Any, last_row: int, last_col: int = 2) -> list[tuple[Any, ...]]:
    """The values of rows 1..last_row, columns 1..last_col."""
    return [tuple(c.value for c in row) for row in ws.iter_rows(min_row=1, max_row=last_row, max_col=last_col)]


def _tab_rgb(ws: Any) -> str | None:
    """The sheet's tab colour as openpyxl reports it, or None when there is none."""
    color = ws.sheet_properties.tabColor
    return None if color is None else color.rgb


def _two_frames(**second: Any) -> list[Any]:
    """The two frames on one sheet, the second from row 6 unless told otherwise."""
    return [(_summary(), "Summary"), (_reasons(), "Summary", {"startrow": 5, **second})]


class TestPlacement:
    """Where the frames land, and what the call returns."""

    def test_two_frames_below_each_other(self, tmp_xlsx: str) -> None:
        """The case from the report: one sheet, the second frame three rows further down."""
        stats = xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx)
        assert stats == [(3, 2), (3, 2)]
        wb = load_workbook(tmp_xlsx)
        assert wb.sheetnames == ["Summary"]
        assert _rows(wb["Summary"], 8) == [
            ("Item", "Value"),
            ("rows checked", "120"),
            ("rows flagged", "7"),
            (None, None),
            (None, None),
            ("Reason", "Count"),
            ("missing", 4),
            ("outdated", 3),
        ]

    def test_polars_frames(self, tmp_xlsx: str) -> None:
        """Polars frames are placed the same way."""
        pl = pytest.importorskip("polars")
        sheets = [(pl.from_pandas(_summary()), "Summary"), (pl.from_pandas(_reasons()), "Summary", {"startrow": 5})]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["A1"].value, ws["A6"].value, ws["B8"].value) == ("Item", "Reason", 3)

    def test_one_frame_per_sheet_is_unchanged(self, tmp_xlsx: str) -> None:
        """Control: entries with different names still get a sheet each."""
        xlsxturbo.dfs_to_xlsx([(_summary(), "A"), (_reasons(), "B")], tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        assert wb.sheetnames == ["A", "B"]
        assert (wb["A"]["A1"].value, wb["B"]["A1"].value) == ("Item", "Reason")

    def test_frames_of_a_sheet_need_not_be_adjacent_in_the_list(self, tmp_xlsx: str) -> None:
        """Another sheet between two entries of one sheet; sheets keep first-appearance order."""
        sheets = [
            (_summary(), "Summary"),
            (_reasons(), "Detail"),
            (_reasons(), "Summary", {"startrow": 5}),
            (_summary(), "Summary", {"startrow": 10}),
        ]
        assert xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx) == [(3, 2)] * 4
        wb = load_workbook(tmp_xlsx)
        assert wb.sheetnames == ["Summary", "Detail"]
        ws = wb["Summary"]
        assert (ws["A1"].value, ws["A6"].value, ws["A11"].value) == ("Item", "Reason", "Item")
        assert wb["Detail"]["A1"].value == "Reason"

    def test_the_upper_frame_may_be_listed_second(self, tmp_xlsx: str) -> None:
        """List order is free in normal mode; the first entry starts lower here."""
        sheets = [(_reasons(), "Summary", {"startrow": 5}), (_summary(), "Summary")]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["A1"].value, ws["A6"].value) == ("Item", "Reason")

    def test_names_that_differ_only_by_case_are_one_sheet(self, tmp_xlsx: str) -> None:
        """Excel compares sheet names without case; the sheet keeps the first spelling."""
        xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), (_reasons(), "SUMMARY", {"startrow": 5})], tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        assert wb.sheetnames == ["Summary"]
        assert wb["Summary"]["A6"].value == "Reason"

    def test_a_second_frame_without_header(self, tmp_xlsx: str) -> None:
        """header=False applies to the frame that sets it and not to the other."""
        xlsxturbo.dfs_to_xlsx(_two_frames(header=False), tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["A1"].value, ws["A6"].value, ws["A7"].value) == ("Item", "missing", "outdated")


class TestOverlap:
    """Frames that share a row are refused before anything is written."""

    def test_overlap_is_refused_by_name(self, tmp_xlsx: str) -> None:
        """The second frame starts on the last row of the first."""
        Path(tmp_xlsx).unlink()
        sheets = [(_summary(), "Summary"), (_reasons(), "Summary", {"startrow": 2})]
        message = r"sheets\[0\] and sheets\[1\] overlap on sheet 'Summary': rows 1-3 and rows 3-5"
        with pytest.raises(xlsxturbo.FrameOverlapError, match=message) as caught:
            xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)
        assert isinstance(caught.value, xlsxturbo.WorkbookValidationError)
        assert not Path(tmp_xlsx).exists()

    def test_two_frames_at_the_default_row_overlap(self, tmp_xlsx: str) -> None:
        """Naming a sheet twice without a startrow is an overlap, said in those words."""
        message = "overlap on sheet 'Summary': rows 1-3 and rows 1-3"
        with pytest.raises(xlsxturbo.WorkbookValidationError, match=message):
            xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), (_reasons(), "Summary")], tmp_xlsx)

    @pytest.mark.parametrize("handler", [xlsxturbo.FileError, OSError, ValueError, xlsxturbo.XlsxTurboError])
    def test_handlers_written_for_a_repeated_sheet_name_still_catch_it(
        self, tmp_xlsx: str, handler: type[Exception]
    ) -> None:
        """A repeated sheet name raised FileError until 1.10.0, and each of its bases still catches it.

        The same call, with no startrow, is what an existing program makes. It
        is an overlap now, and the class it raises keeps every base the old one
        had. ``errno`` is None: no file is involved.
        """
        with pytest.raises(handler, match="overlap") as caught:
            xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), (_reasons(), "Summary")], tmp_xlsx)
        assert type(caught.value) is xlsxturbo.FrameOverlapError
        assert caught.value.errno is None  # type: ignore[attr-defined]

    def test_frames_that_touch_are_accepted(self, tmp_xlsx: str) -> None:
        """The second frame starts on the row after the first: no gap, no overlap."""
        xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), (_reasons(), "Summary", {"startrow": 3})], tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["A3"].value, ws["A4"].value) == ("rows flagged", "Reason")

    def test_a_frame_that_writes_nothing_overlaps_nothing(self, tmp_xlsx: str) -> None:
        """No rows and no header: the frame occupies no row, wherever it is put."""
        empty = _reasons().iloc[0:0]
        xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), (empty, "Summary", {"header": False})], tmp_xlsx)
        assert load_workbook(tmp_xlsx)["Summary"]["A1"].value == "Item"

    def test_an_empty_frame_with_a_header_occupies_its_header_row(self, tmp_xlsx: str) -> None:
        """No rows but a header: one row, which can overlap."""
        sheets = [(_summary(), "Summary"), (_reasons().iloc[0:0], "Summary", {"startrow": 1})]
        with pytest.raises(xlsxturbo.WorkbookValidationError, match="rows 1-3 and rows 2-2"):
            xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)


class TestWholeSheetOptions:
    """The first entry for a sheet decides the options that describe the sheet."""

    @pytest.mark.parametrize(
        ("option", "value"),
        [
            ("autofit", True),
            ("freeze_panes", True),
            ("freeze_panes", False),
            ("tab_color", "red"),
            ("tab_color", None),
            ("column_widths", {0: 30}),
            ("column_widths", {}),
            ("row_heights", {0: 30}),
        ],
    )
    def test_a_later_entry_may_not_set_one(self, tmp_xlsx: str, option: str, value: Any) -> None:
        """Set on the second entry, each is refused, naming the option and both entries."""
        message = rf"sheets\[1\] \('Summary'\): '{option}' describes the whole sheet, and sheets\[0\]"
        with pytest.raises(xlsxturbo.ConfigurationError, match=message):
            xlsxturbo.dfs_to_xlsx(_two_frames(**{option: value}), tmp_xlsx)

    def test_the_first_entry_sets_them(self, tmp_xlsx: str) -> None:
        """Tab colour, freeze panes, widths and row heights from the first entry's dict."""
        first: SheetOptions = {
            "tab_color": "#4472C4",
            "freeze_panes": True,
            "column_widths": {1: 33},
            "row_heights": {5: 40},
        }
        xlsxturbo.dfs_to_xlsx([(_summary(), "Summary", first), (_reasons(), "Summary", {"startrow": 5})], tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert _tab_rgb(ws) == "FF4472C4"
        assert ws.freeze_panes == "A2"
        assert ws.column_dimensions["B"].width == pytest.approx(33, abs=1)
        assert ws.row_dimensions[6].height == pytest.approx(40, abs=0.5)

    def test_the_keywords_apply_once_to_the_sheet(self, tmp_xlsx: str) -> None:
        """A workbook-wide freeze follows the first entry's header, not the last frame's."""
        xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx, freeze_panes=True, tab_color="red")
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert ws.freeze_panes == "A2"
        assert _tab_rgb(ws) == "FFFF0000"

    def test_autofit_sees_every_frame(self, tmp_xlsx: str) -> None:
        """Column A is fitted to the longest value on the sheet, which is in the second frame."""
        long = pd.DataFrame({"Reason": ["a reason that is far longer than anything above"], "Count": [1]})
        xlsxturbo.dfs_to_xlsx([(_summary(), "S"), (long, "S", {"startrow": 5})], tmp_xlsx, autofit=True)
        fitted = load_workbook(tmp_xlsx)["S"].column_dimensions["A"].width
        xlsxturbo.dfs_to_xlsx([(_summary(), "S")], tmp_xlsx, autofit=True)
        alone = load_workbook(tmp_xlsx)["S"].column_dimensions["A"].width
        assert alone < 20 < 40 < fitted

    def test_the_capped_autofit_takes_the_widest_frame_per_column(self, tmp_xlsx: str) -> None:
        """With an '_all' cap the estimate is the maximum over the frames, in either order.

        The third column exists in the second frame only and is fitted as well.
        """
        wide = pd.DataFrame({"Reason": ["a reason of thirty characters."], "Count": [1], "Third column": ["x"]})
        widths = []
        for sheets in (
            [(_summary(), "S"), (wide, "S", {"startrow": 5})],
            [(wide, "S", {"startrow": 5}), (_summary(), "S")],
        ):
            xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx, autofit=True, column_widths={"_all": 60})
            dims = load_workbook(tmp_xlsx)["S"].column_dimensions
            widths.append((dims["A"].width, dims["C"].width))
        assert widths[0] == widths[1]
        assert widths[0][0] == pytest.approx(31.71, abs=0.01)
        assert widths[0][1] == pytest.approx(13.71, abs=0.01)


class TestFrameBoundOptions:
    """Tables, formats, formulas and rules belong to the frame that asks for them."""

    def test_two_tables_on_one_sheet(self, tmp_xlsx: str) -> None:
        """Each frame becomes its own table, over its own range."""
        sheets = [
            (_summary(), "Summary", {"table_style": "Medium2", "table_name": "Figures"}),
            (_reasons(), "Summary", {"startrow": 5, "table_style": "Light1", "table_name": "Reasons"}),
        ]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)
        tables = load_workbook(tmp_xlsx)["Summary"].tables
        assert dict(tables.items()) == {"Figures": "A1:B3", "Reasons": "A6:B8"}

    def test_a_workbook_wide_table_style_gives_each_frame_a_table(self, tmp_xlsx: str) -> None:
        """The keyword is a default per frame; the tables are auto-named in list order."""
        xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx, table_style="Medium2")
        tables = load_workbook(tmp_xlsx)["Summary"].tables
        assert dict(tables.items()) == {"Table1": "A1:B3", "Table2": "A6:B8"}

    def test_one_table_name_for_two_frames_is_refused(self, tmp_xlsx: str) -> None:
        """Table names are unique in the workbook, also on one sheet."""
        with pytest.raises(xlsxturbo.WorkbookValidationError, match="Duplicate table name 'T'"):
            xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx, table_style="Medium2", table_name="T")

    def test_formula_columns_use_the_rows_of_their_frame(self, tmp_xlsx: str) -> None:
        """{row} is the sheet row of the second frame's data."""
        xlsxturbo.dfs_to_xlsx(_two_frames(formula_columns={"Double": "=B{row}*2"}), tmp_xlsx)
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["C1"].value, ws["C6"].value, ws["C7"].value, ws["C8"].value) == (None, "Double", "=B7*2", "=B8*2")

    def test_column_formats_and_rules_stay_with_their_frame(self, tmp_xlsx: str) -> None:
        """A format, a conditional format and a validation given for the second frame only."""
        second = {
            "column_formats": {"Count": {"num_format": "0.00"}},
            "conditional_formats": {"Count": {"type": "data_bar"}},
            "validations": {"Reason": {"type": "text_length", "min": 1, "max": 50}},
        }
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            xlsxturbo.dfs_to_xlsx(_two_frames(**second), tmp_xlsx)
            ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["B2"].number_format, ws["B7"].number_format) == ("General", "0.00")
        assert [str(r.sqref) for r in ws.conditional_formatting] == ["B7:B8"]
        assert [str(v.sqref) for v in ws.data_validations.dataValidation] == ["A7:A8"]

    def test_a_workbook_wide_format_applies_to_each_frame_that_has_the_column(self, tmp_xlsx: str) -> None:
        """A keyword column format is a default for every frame, like for every sheet."""
        other = pd.DataFrame({"Item": ["x"], "Value": [1.5]})
        sheets = [(pd.DataFrame({"Item": ["y"], "Value": [2.5]}), "S"), (other, "S", {"startrow": 4})]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx, column_formats={"Value": {"num_format": "0.00"}})
        ws = load_workbook(tmp_xlsx)["S"]
        assert (ws["B2"].number_format, ws["B6"].number_format) == ("0.00", "0.00")


class TestAutofilter:
    """A sheet holds one plain-range filter."""

    def test_one_frame_may_carry_the_filter(self, tmp_xlsx: str) -> None:
        """The filter covers the frame that asks for it."""
        xlsxturbo.dfs_to_xlsx(_two_frames(autofilter=True), tmp_xlsx)
        assert load_workbook(tmp_xlsx)["Summary"].auto_filter.ref == "A6:B8"

    def test_two_filters_are_refused(self, tmp_xlsx: str) -> None:
        """The workbook-wide keyword asks for a filter on both frames."""
        message = r"sheet 'Summary': sheets\[0\] and sheets\[1\] both have autofilter=True"
        with pytest.raises(xlsxturbo.ConfigurationError, match=message):
            xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx, autofilter=True)

    def test_the_keyword_with_one_frame_switched_off(self, tmp_xlsx: str) -> None:
        """What the message recommends works: the first frame keeps the filter."""
        xlsxturbo.dfs_to_xlsx(_two_frames(autofilter=False), tmp_xlsx, autofilter=True)
        assert load_workbook(tmp_xlsx)["Summary"].auto_filter.ref == "A1:B3"


class TestCellAddressedOptions:
    """An option addressed by cell reference is applied once, and after every frame."""

    def test_workbook_wide_options_are_applied_once(self, tmp_xlsx: str) -> None:
        """A keyword merged range and comment with two frames on the sheet.

        Applied once per frame, the merged range would be refused by the writer
        as overlapping itself.
        """
        xlsxturbo.dfs_to_xlsx(
            _two_frames(),
            tmp_xlsx,
            merged_ranges=[("D1:E1", "Title")],
            comments={"D3": "a note"},
            hyperlinks=[("D4", "https://example.com", "link")],
        )
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert [str(r) for r in ws.merged_cells.ranges] == ["D1:E1"]
        assert ws["D3"].comment is not None
        assert ws["D4"].value == "link"

    def test_two_sheets_still_get_a_workbook_wide_option_each(self, tmp_xlsx: str) -> None:
        """Control: with one frame per sheet the keyword applies to both sheets."""
        xlsxturbo.dfs_to_xlsx([(_summary(), "A"), (_reasons(), "B")], tmp_xlsx, cells={"D1": "note"})
        wb = load_workbook(tmp_xlsx)
        assert (wb["A"]["D1"].value, wb["B"]["D1"].value) == ("note", "note")

    def test_a_later_entry_applies_its_own(self, tmp_xlsx: str) -> None:
        """A title cell given with the second frame is written."""
        xlsxturbo.dfs_to_xlsx(_two_frames(cells={"A5": "Reasons"}), tmp_xlsx, cells={"D1": "note"})
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert (ws["A5"].value, ws["D1"].value) == ("Reasons", "note")

    def test_cells_overwrite_the_data_of_a_later_frame(self, tmp_xlsx: str) -> None:
        """A cell from the first entry and one from the keyword land on the second frame's data."""
        sheets = [(_summary(), "Summary", {"cells": {"A7": "corrected"}}), (_reasons(), "Summary", {"startrow": 5})]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)
        assert load_workbook(tmp_xlsx)["Summary"]["A7"].value == "corrected"
        xlsxturbo.dfs_to_xlsx(_two_frames(), tmp_xlsx, cells={"B8": 99})
        assert load_workbook(tmp_xlsx)["Summary"]["B8"].value == 99


class TestConstantMemory:
    """Rows are written once and in order, so the frames are listed from top to bottom."""

    def test_frames_from_top_to_bottom(self, tmp_xlsx: str) -> None:
        """Two frames in order are written in full, with the sheet options of the first."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            stats = xlsxturbo.dfs_to_xlsx(
                _two_frames(), tmp_xlsx, constant_memory=True, freeze_panes=True, column_widths={0: 30}
            )
        assert stats == [(3, 2), (3, 2)]
        ws = load_workbook(tmp_xlsx)["Summary"]
        assert _rows(ws, 8)[5:] == [("Reason", "Count"), ("missing", 4), ("outdated", 3)]
        assert _rows(ws, 3) == [("Item", "Value"), ("rows checked", "120"), ("rows flagged", "7")]
        assert ws.freeze_panes == "A2"
        assert ws.column_dimensions["A"].width == pytest.approx(30, abs=1)

    def test_frames_listed_bottom_first_are_refused(self, tmp_xlsx: str) -> None:
        """The lower frame first: the upper one cannot be written after it."""
        sheets = [(_reasons(), "Summary", {"startrow": 5}), (_summary(), "Summary")]
        message = r"constant_memory=True writes rows in order, so sheets\[1\] \(rows 1-3\) cannot follow sheets\[0\]"
        with pytest.raises(xlsxturbo.ConfigurationError, match=message):
            xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx, constant_memory=True)

    def test_the_same_list_is_fine_in_normal_mode(self, tmp_xlsx: str) -> None:
        """Control: the order rule belongs to constant-memory mode only."""
        sheets = [(_reasons(), "Summary", {"startrow": 5}), (_summary(), "Summary")]
        assert xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx) == [(3, 2), (3, 2)]


class TestErrors:
    """A failure in one frame names its sheet."""

    def test_a_write_error_in_the_second_frame_names_the_sheet(self, tmp_xlsx: str) -> None:
        """A column format for a column the second frame does not have."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"sheet 'Summary'.*column_formats.*Missing"):
            xlsxturbo.dfs_to_xlsx(_two_frames(column_formats={"Missing": {"bold": True}}), tmp_xlsx)

    def test_an_unsupported_frame_is_refused_before_anything_is_written(self, tmp_xlsx: str) -> None:
        """The second entry is not a DataFrame."""
        Path(tmp_xlsx).unlink()
        with pytest.raises(xlsxturbo.InputDataError, match="sheet 'Summary'"):
            xlsxturbo.dfs_to_xlsx([(_summary(), "Summary"), ([1, 2], "Summary", {"startrow": 5})], tmp_xlsx)
        assert not Path(tmp_xlsx).exists()
