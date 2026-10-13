"""Tests for the sheet-level options: tab_color, freeze_panes as (row, col), autofilter.

All three are properties of the sheet rather than cell writes, which is why they
also apply under ``constant_memory``. Each test reads the produced file back.
"""

from __future__ import annotations

import warnings
import zipfile
from typing import Any

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


def _frame() -> pd.DataFrame:
    """Two columns, three rows."""
    return pd.DataFrame({"Name": ["Alice", "Bob", "Carol"], "Score": [10, 50, 90]})


def _parts(path: str) -> list[str]:
    """The names of the parts in the workbook package."""
    with zipfile.ZipFile(path) as package:
        return package.namelist()


def _tab_rgb(ws: Any) -> str | None:
    """The sheet's tab colour as openpyxl reports it, or None when there is none."""
    color = ws.sheet_properties.tabColor
    return None if color is None else color.rgb


class TestTabColor:
    """tab_color colours the sheet tab."""

    @pytest.mark.parametrize(("value", "rgb"), [("#4472C4", "FF4472C4"), ("red", "FFFF0000"), ("#70ad47", "FF70AD47")])
    def test_hex_and_named_colours(self, tmp_xlsx: str, value: str, rgb: str) -> None:
        """A hex value and a named colour both reach the sheet properties."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, tab_color=value)
        assert _tab_rgb(active_ws(load_workbook(tmp_xlsx))) == rgb

    def test_no_colour_by_default(self, tmp_xlsx: str) -> None:
        """Control: without the option the tab has no colour."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx)
        assert _tab_rgb(active_ws(load_workbook(tmp_xlsx))) is None

    def test_per_sheet_colours(self, tmp_xlsx: str) -> None:
        """Each sheet takes its own colour; a sheet without one stays uncoloured."""
        xlsxturbo.dfs_to_xlsx(
            [
                (_frame(), "Inputs", {"tab_color": "#4472C4"}),
                (_frame(), "Results", {"tab_color": "#70AD47"}),
                (_frame(), "Legend"),
            ],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        assert [_tab_rgb(wb[name]) for name in ("Inputs", "Results", "Legend")] == ["FF4472C4", "FF70AD47", None]

    def test_per_sheet_none_removes_the_workbook_colour(self, tmp_xlsx: str) -> None:
        """A per-sheet value overrides the keyword, and None turns it off."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "A"), (_frame(), "B", {"tab_color": "blue"}), (_frame(), "C", {"tab_color": None})],
            tmp_xlsx,
            tab_color="red",
        )
        wb = load_workbook(tmp_xlsx)
        assert [_tab_rgb(wb[name]) for name in "ABC"] == ["FFFF0000", "FF0000FF", None]

    def test_applies_under_constant_memory(self, tmp_xlsx: str) -> None:
        """The colour is written in constant-memory mode, without a warning."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, tab_color="#4472C4", constant_memory=True)
        assert _tab_rgb(active_ws(load_workbook(tmp_xlsx))) == "FF4472C4"

    @pytest.mark.parametrize("value", ["#12345", "#GGGGGG", "turquoise", ""])
    def test_an_unusable_colour_is_refused(self, tmp_xlsx: str, value: str) -> None:
        """A colour the parser cannot read raises; the writer would ignore it silently."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="tab_color"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, tab_color=value)
        with pytest.raises(xlsxturbo.ConfigurationError, match="sheet option 'tab_color'"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"tab_color": value})], tmp_xlsx)

    def test_a_non_string_is_a_type_error(self, tmp_xlsx: str) -> None:
        """A number is refused as a ConfigurationTypeError naming the option."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="tab_color must be a color string, got int"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, tab_color=0xFF0000)  # type: ignore[arg-type]


class TestFreezePanesAt:
    """freeze_panes=(row, col) freezes at a chosen cell."""

    @pytest.mark.parametrize(
        ("value", "top_left"),
        [((1, 1), "B2"), ((4, 1), "B5"), ((0, 2), "C1"), ((3, 0), "A4"), ([1, 1], "B2")],
    )
    def test_the_pair_is_the_top_left_scrolling_cell(self, tmp_xlsx: str, value: Any, top_left: str) -> None:
        """(row, col) is zero-based, as XlsxWriter's freeze_panes(row, col) takes it."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=value)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == top_left

    def test_the_issue_example(self, tmp_xlsx: str) -> None:
        """Three title rows, the header and the key column stay in view."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=3, freeze_panes=(4, 1))
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == "B5"

    def test_zero_zero_freezes_nothing(self, tmp_xlsx: str) -> None:
        """(0, 0) is a pane with nothing above or left of it: no freeze."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=(0, 0))
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes is None

    def test_true_still_means_below_the_header(self, tmp_xlsx: str) -> None:
        """The bool form is unchanged and follows startrow."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=True, startrow=2)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == "A4"

    def test_the_pair_applies_without_a_header(self, tmp_xlsx: str) -> None:
        """A pair is taken as given with header=False, where True freezes nothing."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=(1, 1), header=False)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == "B2"
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=True, header=False)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes is None

    def test_per_sheet_pair_overrides_the_keyword(self, tmp_xlsx: str) -> None:
        """A per-sheet pair, a per-sheet False and the workbook default each apply."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "A"), (_frame(), "B", {"freeze_panes": (1, 2)}), (_frame(), "C", {"freeze_panes": False})],
            tmp_xlsx,
            freeze_panes=True,
        )
        wb = load_workbook(tmp_xlsx)
        assert [wb[name].freeze_panes for name in "ABC"] == ["A2", "C2", None]

    @pytest.mark.parametrize("value", [True, (2, 1)])
    def test_applies_under_constant_memory(self, tmp_xlsx: str, value: Any) -> None:
        """Both forms are written in constant-memory mode, without a warning."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=value, constant_memory=True)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == ("A2" if value is True else "B3")

    @pytest.mark.parametrize(
        ("value", "message"),
        [
            ((1,), "got 1 elements"),
            ((1, 2, 3), "got 3 elements"),
            ((-1, 0), "freeze_panes row must be between 0 and 1048575, got -1"),
            ((1_048_576, 0), "freeze_panes row must be between 0 and 1048575"),
            ((0, 16_384), "freeze_panes column must be between 0 and 16383, got 16384"),
        ],
    )
    def test_a_pair_outside_the_grid_is_refused(self, tmp_xlsx: str, value: Any, message: str) -> None:
        """A wrong length or a position off the grid raises, naming option and value."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=message):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=value)

    @pytest.mark.parametrize("value", [(True, 0), (0, False), ("1", 0), (1.0, 0), "A2", 1])
    def test_a_wrong_type_is_a_type_error(self, tmp_xlsx: str, value: Any) -> None:
        """A bool inside the pair, a string and a bare int are all refused."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="freeze_panes"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, freeze_panes=value)
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="sheet option 'freeze_panes'"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"freeze_panes": value})], tmp_xlsx)


class TestAutofilter:
    """autofilter=True puts filter buttons on a plain range, with no table."""

    def test_the_filter_spans_header_and_data(self, tmp_xlsx: str) -> None:
        """The range is the header row and every data row, and no table part exists."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A1:B4"
        assert not [name for name in _parts(tmp_xlsx) if "table" in name]

    def test_no_filter_by_default(self, tmp_xlsx: str) -> None:
        """Control: without the option the sheet has no filter."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref is None

    def test_a_table_still_writes_a_table_part(self, tmp_xlsx: str) -> None:
        """Control for the part check above: table_style does write xl/tables/."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, table_style="None")
        assert [name for name in _parts(tmp_xlsx) if "table" in name] == ["xl/tables/table1.xml"]

    def test_the_filter_follows_startrow(self, tmp_xlsx: str) -> None:
        """With startrow the filter starts on the header row, not on row 1."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True, startrow=3)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A4:B7"

    def test_the_filter_covers_formula_columns(self, tmp_xlsx: str) -> None:
        """Unlike a table, the plain filter includes the formula columns."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True, formula_columns={"Double": "=B{row}*2"})
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A1:C4"

    def test_a_frame_without_rows_gets_a_header_only_filter(self, tmp_xlsx: str) -> None:
        """No data rows: the filter covers the header row alone."""
        xlsxturbo.df_to_xlsx(_frame().iloc[0:0], tmp_xlsx, autofilter=True)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A1:B1"

    def test_a_frame_without_columns_gets_no_filter(self, tmp_xlsx: str) -> None:
        """No columns: there is nothing to put a button on, and nothing is raised."""
        xlsxturbo.df_to_xlsx(pd.DataFrame(index=range(3)), tmp_xlsx, autofilter=True)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref is None

    def test_per_sheet(self, tmp_xlsx: str) -> None:
        """The per-sheet option switches the filter on and off."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "A"), (_frame(), "B", {"autofilter": False})],
            tmp_xlsx,
            autofilter=True,
        )
        wb = load_workbook(tmp_xlsx)
        assert [wb[name].auto_filter.ref for name in "AB"] == ["A1:B4", None]

    def test_applies_under_constant_memory(self, tmp_xlsx: str) -> None:
        """The filter is written in constant-memory mode, without a warning."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True, constant_memory=True)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A1:B4"

    def test_polars_frame(self, tmp_xlsx: str) -> None:
        """The filter range is the same for a polars frame."""
        pl = pytest.importorskip("polars")
        xlsxturbo.df_to_xlsx(pl.from_pandas(_frame()), tmp_xlsx, autofilter=True)
        assert active_ws(load_workbook(tmp_xlsx)).auto_filter.ref == "A1:B4"

    def test_with_a_table_style_it_is_refused(self, tmp_xlsx: str) -> None:
        """A table has its own filter, so the pair is a configuration error."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="autofilter=True cannot be combined with table_style"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True, table_style="Medium2")

    def test_the_conflict_names_the_sheet(self, tmp_xlsx: str) -> None:
        """A workbook-wide filter meeting a per-sheet table is refused by sheet name."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="sheet 'B': autofilter=True cannot be combined"):
            xlsxturbo.dfs_to_xlsx(
                [(_frame(), "A"), (_frame(), "B", {"table_style": "Light1"})],
                tmp_xlsx,
                autofilter=True,
            )

    def test_without_a_header_it_is_refused(self, tmp_xlsx: str) -> None:
        """Without a header the first data row would carry the buttons."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="autofilter=True needs the header row"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, autofilter=True, header=False)

    def test_a_non_bool_is_a_type_error(self, tmp_xlsx: str) -> None:
        """A per-sheet value that is not a bool is refused, naming the option."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="sheet option 'autofilter' must be a bool"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"autofilter": "yes"})], tmp_xlsx)  # type: ignore[typeddict-item]


def _width(path: str, letter: str = "A") -> float:
    """The stored width of one column."""
    return active_ws(load_workbook(path)).column_dimensions[letter].width


class TestFilterButtonWidth:
    """With autofit and an '_all' cap, a filtered header gets room for its button.

    That path estimates widths itself and does not call the writer's autofit,
    which is where the allowance otherwise comes from. The header here is 18
    characters and longer than every value, so it decides the column width.
    """

    HEADER = "a_long_header_name"
    BUTTON = 16 / 7

    def _write(self, path: str, cap: float = 60, **options: Any) -> float:
        """Write one column under an autofit cap and return its width."""
        df = pd.DataFrame({self.HEADER: [1, 2]})
        xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": cap}, **options)
        return _width(path)

    def test_an_autofilter_header_is_wider_by_the_button(self, tmp_xlsx: str) -> None:
        """The same sheet with a filter is one button wider than without."""
        plain = self._write(tmp_xlsx)
        filtered = self._write(tmp_xlsx, autofilter=True)
        assert filtered - plain == pytest.approx(self.BUTTON, abs=0.01)

    def test_a_table_header_is_wider_by_the_button(self, tmp_xlsx: str) -> None:
        """A table's header carries the same button and gets the same room."""
        plain = self._write(tmp_xlsx)
        table = self._write(tmp_xlsx, table_style="Medium2")
        assert table - plain == pytest.approx(self.BUTTON, abs=0.01)

    def test_the_cap_still_wins(self, tmp_xlsx: str) -> None:
        """The allowance never lifts a column over the '_all' cap.

        Compared with the capped sheet without a filter, not with the number 12:
        the writer stores a width of 12 as 12.71.
        """
        capped = self._write(tmp_xlsx, cap=12)
        assert self._write(tmp_xlsx, cap=12, autofilter=True) == capped
        assert capped < self._write(tmp_xlsx)

    def test_a_column_decided_by_its_data_is_unchanged(self, tmp_xlsx: str) -> None:
        """Where a value is much longer than the header, the button adds nothing."""
        df = pd.DataFrame({"id": ["a value far longer than its header"]})
        widths = []
        for autofilter in (False, True):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, column_widths={"_all": 60}, autofilter=autofilter)
            widths.append(_width(tmp_xlsx))
        assert widths[0] == widths[1]

    def test_a_table_that_is_not_created_adds_nothing(self, tmp_xlsx: str) -> None:
        """A frame without rows gets no table, so no button and no extra width."""
        df = pd.DataFrame({self.HEADER: []})
        widths = []
        for style in (None, "Medium2"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, column_widths={"_all": 60}, table_style=style)
            widths.append(_width(tmp_xlsx))
        assert widths[0] == widths[1]
