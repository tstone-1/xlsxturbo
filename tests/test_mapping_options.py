"""The dict- and list-valued options take any mapping or sequence, narrowly typed or not.

Two halves of one promise. The annotations are ``Mapping`` and ``Sequence``
because ``dict`` and ``list`` are invariant: a caller's ``dict[str, str]`` is
refused by a type checker for a parameter declared ``dict[str, str | int | ...]``,
although every value in it is allowed. And since an annotation may not promise
more than the function takes, the extension accepts a mapping that is not a
``dict`` and a sequence that is not a ``list``.

The typing half is checked by pyright, not by pytest: this module is inside the
project's pyright ``include``, so every call in ``TestNarrowAnnotations`` is a
type-check assertion. Reverting one annotation in ``xlsxturbo.pyi`` to ``dict``
turns the lint gate red here. The variables carry explicit annotations on
purpose -- an inferred literal would be widened to fit and prove nothing.
"""

from __future__ import annotations

from collections import OrderedDict, UserDict
from types import MappingProxyType
from typing import Any

import pandas as pd
import pytest
import xlsxturbo
from xlsxturbo.options import ExportOptions
from xlsxturbo.types import CommentOptions, ConditionalFormat, RichTextFormat, SheetOptions

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


def _frame() -> pd.DataFrame:
    """Two columns, three rows."""
    return pd.DataFrame({"Name": ["Alice", "Bob", "Carol"], "Score": [10, 50, 90]})


class TestNarrowAnnotations:
    """A dict annotated with one member of the value union is accepted."""

    def test_cells(self, tmp_xlsx: str) -> None:
        """The case from the report: every cell is text."""
        notes: dict[str, str] = {"D1": "checked", "D2": "open"}
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, cells=notes)
        options: SheetOptions = {"cells": notes}
        xlsxturbo.dfs_to_xlsx([(_frame(), "S", options)], tmp_xlsx)
        bundle = ExportOptions(cells=notes)
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, **bundle.as_kwargs())
        assert active_ws(load_workbook(tmp_xlsx))["D2"].value == "open"

    def test_comments_images_checkboxes_textboxes(self, tmp_xlsx: str) -> None:
        """The simple form and the options form of each are both narrow dicts."""
        comments: dict[str, str] = {"A1": "a note"}
        comment_options: dict[str, CommentOptions] = {"A1": {"text": "a note"}}
        checkboxes: dict[str, bool] = {"D1": True}
        textboxes: dict[str, str] = {"E1": "a textbox"}
        images: dict[str, str] = {}
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, comments=comments, checkboxes=checkboxes, textboxes=textboxes)
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, comments=comment_options, images=images)

    def test_column_widths_by_key_type(self, tmp_xlsx: str) -> None:
        """Index keys, the '_all' key and both together, each as its own dict type.

        A mapping is invariant in its key type as well, which covariance in the
        value does not help: ``Mapping[int | str, float]`` alone refuses a
        ``dict[int, float]``.
        """
        by_index: dict[int, float] = {0: 12.5}
        whole_numbers: dict[int, int] = {0: 12}
        cap: dict[str, int] = {"_all": 30}
        both: dict[int | str, int | float] = {0: 12, "_all": 30}
        for widths in (by_index, whole_numbers, cap, both):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, column_widths=widths)
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, column_widths={0: 12, "_all": 30})

    def test_row_heights_and_conditional_formats(self, tmp_xlsx: str) -> None:
        """Whole-number heights, and a rule dict typed for one of its two value forms."""
        heights: dict[int, int] = {0: 24}
        single: dict[str, ConditionalFormat] = {"Score": {"type": "data_bar"}}
        several: dict[str, list[ConditionalFormat]] = {"Score": [{"type": "data_bar"}]}
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, row_heights=heights, conditional_formats=single)
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, conditional_formats=several)


class TestNarrowSequenceAnnotations:
    """A list annotated with one member of the element union is accepted."""

    def test_merged_ranges_and_hyperlinks(self, tmp_xlsx: str) -> None:
        """Lists of two-element tuples, where the parameter also allows three."""
        merged: list[tuple[str, str]] = [("D1:E1", "Title")]
        links: list[tuple[str, str]] = [("F1", "https://example.com")]
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, merged_ranges=merged, hyperlinks=links)
        options: SheetOptions = {"merged_ranges": merged, "hyperlinks": links}
        xlsxturbo.dfs_to_xlsx([(_frame(), "S", options)], tmp_xlsx)
        ExportOptions(merged_ranges=merged, hyperlinks=links)
        assert active_ws(load_workbook(tmp_xlsx))["D1"].value == "Title"

    def test_rich_text_segments(self, tmp_xlsx: str) -> None:
        """Segment lists that are all plain strings, or all formatted tuples."""
        plain: dict[str, list[str]] = {"D1": ["one ", "two"]}
        formatted: dict[str, list[tuple[str, RichTextFormat]]] = {"D1": [("bold", {"bold": True}), (" x", {})]}
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, rich_text=plain)
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, rich_text=formatted)

    def test_sheets(self, tmp_xlsx: str) -> None:
        """A list of (frame, name) pairs, where the parameter also allows options."""
        sheets: list[tuple[pd.DataFrame, str]] = [(_frame(), "A"), (_frame(), "B")]
        assert xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx) == [(4, 2), (4, 2)]


class TestSequencesAtRuntime:
    """A sequence that is not a list is taken like one."""

    def test_tuples_in_place_of_lists(self, tmp_xlsx: str) -> None:
        """merged_ranges, hyperlinks, rich_text segments and a rule list as tuples."""
        xlsxturbo.df_to_xlsx(
            _frame(),
            tmp_xlsx,
            merged_ranges=(("D1:E1", "Title"),),
            hyperlinks=(("F1", "https://example.com", "link"),),
            rich_text={"G1": ("one ", ("two", {"bold": True}))},
            conditional_formats={"Score": ({"type": "data_bar"},)},
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert (ws["D1"].value, ws["F1"].value, ws["G1"].value) == ("Title", "link", "one two")
        assert "D1:E1" in {str(r) for r in ws.merged_cells.ranges}
        assert len(ws.conditional_formatting) == 1

    def test_per_sheet_tuples(self, tmp_xlsx: str) -> None:
        """The per-sheet path takes a tuple for a list-valued option, and sheets may be a tuple."""
        options: SheetOptions = {"merged_ranges": (("D1:E1", "Title"),), "hyperlinks": (("F1", "https://example.com"),)}
        xlsxturbo.dfs_to_xlsx(((_frame(), "S", options),), tmp_xlsx)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert ws["D1"].value == "Title"
        assert ws["F1"].hyperlink is not None

    def test_a_list_still_works(self, tmp_xlsx: str) -> None:
        """Control: the list form writes the same cells as the tuple form."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, merged_ranges=[("D1:E1", "Title")], rich_text={"G1": ["one ", "two"]})
        ws = active_ws(load_workbook(tmp_xlsx))
        assert (ws["D1"].value, ws["G1"].value) == ("Title", "one two")

    @pytest.mark.parametrize("value", ["D1:E1", b"D1:E1", {"D1:E1": "t"}, 5])
    def test_what_is_not_a_sequence_is_still_refused(self, tmp_xlsx: str, value: Any) -> None:
        """A string, bytes, a dict and a number raise what they raised before."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="expected list for 'merged_ranges'"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, merged_ranges=value)
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="sheet option 'merged_ranges' must be a list"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"merged_ranges": value})], tmp_xlsx)

    def test_a_bare_string_is_not_a_segment_list(self, tmp_xlsx: str) -> None:
        """``Sequence[str]`` lets a string past a type checker; the runtime refuses it.

        Read as a sequence it would be one segment per character.
        """
        message = r"rich_text\['D1'\]: expected list of segments, got str"
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match=message):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, rich_text={"D1": "plain"})


class TestMappingsAtRuntime:
    """A mapping that is not a dict is taken like one."""

    @pytest.mark.parametrize("wrap", [MappingProxyType, UserDict, OrderedDict])
    def test_cells_from_any_mapping(self, tmp_xlsx: str, wrap: Any) -> None:
        """A read-only proxy, a UserDict and an OrderedDict all write their cells."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, cells=wrap({"D1": "x", "D2": 7}))
        ws = active_ws(load_workbook(tmp_xlsx))
        assert (ws["D1"].value, ws["D2"].value) == ("x", 7)

    def test_per_sheet_option_from_a_mapping(self, tmp_xlsx: str) -> None:
        """The per-sheet path takes a mapping for a dict-valued option too."""
        options: SheetOptions = {"cells": MappingProxyType({"D1": "x"})}
        xlsxturbo.dfs_to_xlsx([(_frame(), "S", options)], tmp_xlsx)
        assert active_ws(load_workbook(tmp_xlsx))["D1"].value == "x"

    def test_every_dict_option_takes_a_mapping(self, tmp_xlsx: str) -> None:
        """Each dict-valued keyword accepts a MappingProxyType of a valid value."""
        values: dict[str, dict[Any, Any]] = {
            "column_widths": {0: 30},
            "row_heights": {0: 22},
            "header_format": {"bold": True},
            "column_formats": {"Score": {"bold": True}},
            "conditional_formats": {"Score": {"type": "data_bar"}},
            "formula_columns": {"Double": "=B{row}*2"},
            "comments": {"I1": "a note"},
            "validations": {"Name": {"type": "text_length", "min": 1, "max": 50}},
            "rich_text": {"J1": [("Bold", {"bold": True}), " plain"]},
            "images": {},
            "checkboxes": {"K1": True},
            "textboxes": {"L1": "a textbox"},
            "charts": {"M2": {"type": "bar", "data_range": "Sheet1!$B$2:$B$3"}},
            "sparklines": {"N2": {"range": "Sheet1!A2:B2"}},
            "cells": {"E1": "note"},
            "defined_names": {"MyName": "Sheet1!$A$1"},
        }
        for option, value in values.items():
            kwargs: dict[str, Any] = {option: MappingProxyType(value)}
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, **kwargs)

    def test_the_mapping_order_is_kept(self, tmp_xlsx: str) -> None:
        """Formula columns appear in the mapping's order, as they do for a dict."""
        formulas = MappingProxyType({"Zeta": "=B{row}", "Alpha": "=B{row}*2"})
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, formula_columns=formulas)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert (ws["C1"].value, ws["D1"].value) == ("Zeta", "Alpha")

    @pytest.mark.parametrize("value", [[("D1", "x")], "D1", 5, {"D1"}])
    def test_what_is_not_a_mapping_is_still_refused(self, tmp_xlsx: str, value: Any) -> None:
        """A list of pairs, a string, a number and a set raise what they raised before."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="expected dict for 'cells'"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, cells=value)
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="sheet option 'cells' must be a dict"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"cells": value})], tmp_xlsx)

    def test_defined_names_inner_types_are_classified(self, tmp_xlsx: str) -> None:
        """A non-string name or reference is a ConfigurationTypeError naming the entry."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match=r"defined_names\['Total'\]: expected a string"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, defined_names={"Total": 5})  # type: ignore[dict-item]
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="defined_names: name: expected a string"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, defined_names={5: "=Sheet1!$A$1"})  # type: ignore[dict-item]
