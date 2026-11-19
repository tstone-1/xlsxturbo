"""Tests for ``pivot_tables``: native Excel pivot tables over the frame they are given with.

rust_xlsxwriter writes the pivot table definition and its cache definition, but no
cached records: Excel builds the table from the source when the file is opened. So
these tests read the definitions back from the package XML; what Excel makes of
them is checked with ``scripts/excel_open_check.ps1``, not here.

The source of a pivot table is always its frame's written range, header row
included. A location is ``'A3'`` (the frame's own sheet) or ``'Sheet!A3'`` (another
sheet, created empty when the workbook has none of that name).
"""

from __future__ import annotations

import posixpath
import warnings
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import pandas as pd
import pytest
import xlsxturbo

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _frame() -> pd.DataFrame:
    """Four columns, four rows: the crate's own pivot table example data."""
    return pd.DataFrame(
        {
            "Region": ["East", "West", "East", "West"],
            "Item": ["Apple", "Pear", "Pear", "Apple"],
            "Volume": [9000, 5000, 3000, 7000],
            "Month": ["July", "July", "August", "August"],
        }
    )


@dataclass
class Pivot:
    """One pivot table as the package describes it."""

    sheet: str  # The sheet the pivot table is on
    location: str  # <location ref>; its first cell is the requested anchor
    source_sheet: str
    source_ref: str
    cache_fields: list[str]
    row_fields: list[int] = field(default_factory=list)
    col_fields: list[int] = field(default_factory=list)
    page_fields: list[int] = field(default_factory=list)
    data_fields: list[dict[str, str]] = field(default_factory=list)
    style: str | None = None
    attrs: dict[str, str] = field(default_factory=dict)


def _xml(package: zipfile.ZipFile, part: str) -> ElementTree.Element:
    """One part of the package, parsed."""
    return ElementTree.fromstring(package.read(part))  # noqa: S314 - our own output, not untrusted input


def _rels(package: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Relationship id -> absolute target part, for one part of the package."""
    folder, name = posixpath.split(part)
    rels_part = posixpath.join(folder, "_rels", name + ".rels")
    if rels_part not in package.namelist():
        return {}
    return {
        r.attrib["Id"]: posixpath.normpath(posixpath.join(folder, r.attrib["Target"]))
        for r in _xml(package, rels_part).iter(f"{PKG_REL}Relationship")
    }


def _sheet_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """Sheet name -> worksheet part, in workbook order."""
    rels = _rels(package, "xl/workbook.xml")
    return {s.attrib["name"]: rels[s.attrib[f"{REL}id"]] for s in _xml(package, "xl/workbook.xml").iter(f"{MAIN}sheet")}


def _pivot(package: zipfile.ZipFile, sheet: str, part: str) -> Pivot:
    """The pivot table definition at ``part`` and its cache definition."""
    table = _xml(package, part)
    cache_part = next(t for t in _rels(package, part).values() if "/pivotCache/" in t)
    cache = _xml(package, cache_part)
    source = cache.find(f"{MAIN}cacheSource/{MAIN}worksheetSource")
    location = table.find(f"{MAIN}location")
    style = table.find(f"{MAIN}pivotTableStyleInfo")
    assert source is not None, cache_part
    assert location is not None, part
    return Pivot(
        sheet=sheet,
        location=location.attrib["ref"],
        source_sheet=source.attrib["sheet"],
        source_ref=source.attrib["ref"],
        cache_fields=[f.attrib["name"] for f in cache.iter(f"{MAIN}cacheField")],
        row_fields=[int(f.attrib["x"]) for f in table.iterfind(f"{MAIN}rowFields/{MAIN}field")],
        col_fields=[int(f.attrib["x"]) for f in table.iterfind(f"{MAIN}colFields/{MAIN}field")],
        page_fields=[int(f.attrib["fld"]) for f in table.iterfind(f"{MAIN}pageFields/{MAIN}pageField")],
        data_fields=[dict(f.attrib) for f in table.iterfind(f"{MAIN}dataFields/{MAIN}dataField")],
        style=None if style is None else style.attrib.get("name"),
        attrs=dict(table.attrib),
    )


def _pivots(path: str) -> list[Pivot]:
    """Every pivot table in the workbook, in sheet order."""
    with zipfile.ZipFile(path) as package:
        return [
            _pivot(package, sheet, target)
            for sheet, sheet_part in _sheet_parts(package).items()
            for target in _rels(package, sheet_part).values()
            if "/pivotTables/" in target
        ]


def _only(path: str) -> Pivot:
    """The workbook's one pivot table."""
    pivots = _pivots(path)
    assert len(pivots) == 1, pivots
    return pivots[0]


def _sheet_names(path: str) -> list[str]:
    """The workbook's sheet names, in order."""
    with zipfile.ZipFile(path) as package:
        return list(_sheet_parts(package))


def _anchor(pivot: Pivot) -> str:
    """The first cell of the pivot table's location."""
    return pivot.location.split(":")[0]


def _write(path: str, pivot_tables: Any, **kwargs: Any) -> None:
    """Write ``_frame()`` with ``pivot_tables``, typed ``Any`` for the cases that are invalid on purpose."""
    xlsxturbo.df_to_xlsx(kwargs.pop("frame", _frame()), path, pivot_tables=pivot_tables, **kwargs)


class TestPivotTableSource:
    """The source is the frame's written range, whatever the frame's position."""

    def test_source_is_the_frame_with_its_header(self, tmp_xlsx: str) -> None:
        """Header row plus four data rows, four columns, named by the frame."""
        _write(tmp_xlsx, {"F3": {"rows": ["Region"]}}, sheet_name="Data")
        pivot = _only(tmp_xlsx)
        assert (pivot.source_sheet, pivot.source_ref) == ("Data", "A1:D5")
        assert pivot.cache_fields == ["Region", "Item", "Volume", "Month"]

    def test_source_follows_startrow(self, tmp_xlsx: str) -> None:
        """A frame moved down by startrow moves the source with it."""
        _write(tmp_xlsx, {"Pivot!A3": {}}, sheet_name="Data", startrow=2)
        assert _only(tmp_xlsx).source_ref == "A3:D7"

    def test_source_of_each_frame_on_a_shared_sheet(self, tmp_xlsx: str) -> None:
        """Two frames on one sheet each get a pivot table over their own range."""
        xlsxturbo.dfs_to_xlsx(
            [
                (_frame(), "Data", {"pivot_tables": {"Pivot!A3": {"rows": ["Region"]}}}),
                (_frame().iloc[:2], "Data", {"startrow": 8, "pivot_tables": {"Pivot!F3": {"rows": ["Item"]}}}),
            ],
            tmp_xlsx,
        )
        refs = sorted((_anchor(p), p.source_ref) for p in _pivots(tmp_xlsx))
        assert refs == [("A3", "A1:D5"), ("F3", "A9:D11")]

    def test_sheet_names_needing_quotes(self, tmp_xlsx: str) -> None:
        """A quoted target sheet and a source sheet with a space both work."""
        _write(tmp_xlsx, {"'My Pivot'!A3": {}}, sheet_name="Sales 2026")
        pivot = _only(tmp_xlsx)
        assert (pivot.sheet, pivot.source_sheet) == ("My Pivot", "Sales 2026")


class TestPivotTableFields:
    """rows, columns, filters and values reach the definition as column indices."""

    def test_row_column_and_filter_fields(self, tmp_xlsx: str) -> None:
        """Each field list names the frame column at the expected index."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Region"], "columns": ["Month"], "filters": ["Item"]}})
        pivot = _only(tmp_xlsx)
        assert (pivot.row_fields, pivot.col_fields, pivot.page_fields) == ([0], [3], [1])

    def test_field_order_is_kept(self, tmp_xlsx: str) -> None:
        """The first row field is the outermost, whatever its column position."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Month", "Region"]}})
        assert _only(tmp_xlsx).row_fields == [3, 0]

    def test_a_bare_value_name_sums_the_column(self, tmp_xlsx: str) -> None:
        """A plain column name in values is a sum, Excel's default."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Region"], "values": ["Volume"]}})
        (data,) = _only(tmp_xlsx).data_fields
        assert data["fld"] == "2"
        assert "subtotal" not in data

    @pytest.mark.parametrize(
        ("function", "subtotal"),
        [
            ("sum", None),
            ("count", "count"),
            ("average", "average"),
            ("max", "max"),
            ("min", "min"),
            ("product", "product"),
            ("count_numbers", "countNums"),
            ("std_dev", "stdDev"),
            ("std_dev_p", "stdDevp"),
            ("var", "var"),
            ("var_p", "varp"),
        ],
    )
    def test_value_function(self, tmp_xlsx: str, function: Any, subtotal: str | None) -> None:
        """Every documented function reaches the dataField's subtotal attribute."""
        _write(tmp_xlsx, {"Pivot!A3": {"values": [{"field": "Volume", "function": function}]}})
        (data,) = _only(tmp_xlsx).data_fields
        assert data.get("subtotal") == subtotal

    def test_value_name_and_number_format(self, tmp_xlsx: str) -> None:
        """A custom caption and a number format reach the dataField."""
        _write(tmp_xlsx, {"Pivot!A3": {"values": [{"field": "Volume", "name": "Total", "num_format": "#,##0.00"}]}})
        (data,) = _only(tmp_xlsx).data_fields
        assert data["name"] == "Total"
        assert int(data["numFmtId"]) > 0

    def test_one_column_as_row_and_value_field(self, tmp_xlsx: str) -> None:
        """Counting the items of the row field is a normal pivot table, accepted by Excel."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Item"], "values": [{"field": "Item", "function": "count"}]}})
        pivot = _only(tmp_xlsx)
        assert (pivot.row_fields, pivot.data_fields[0]["fld"]) == ([1], "1")


class TestPivotTableAppearance:
    """style, layout and the grand totals."""

    def test_default_style_is_light16(self, tmp_xlsx: str) -> None:
        """Without a style the pivot table gets Excel's default."""
        _write(tmp_xlsx, {"Pivot!A3": {}})
        assert _only(tmp_xlsx).style == "PivotStyleLight16"

    @pytest.mark.parametrize("style", ["Light1", "Medium9", "Dark28"])
    def test_style(self, tmp_xlsx: str, style: str) -> None:
        """A style name is the gallery name without the 'PivotStyle' prefix."""
        _write(tmp_xlsx, {"Pivot!A3": {"style": style}})
        assert _only(tmp_xlsx).style == f"PivotStyle{style}"

    def test_style_none_turns_it_off(self, tmp_xlsx: str) -> None:
        """An explicit None writes the style info without a name, as table_style=None does."""
        _write(tmp_xlsx, {"Pivot!A3": {"style": None}})
        assert _only(tmp_xlsx).style is None

    @pytest.mark.parametrize("style", ["Light0", "Light29", "Medium", "PivotStyleMedium9", "TableStyleMedium9", ""])
    def test_unknown_style_is_refused(self, tmp_xlsx: str, style: str) -> None:
        """A name outside the gallery is refused, not written for Excel to drop."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="style"):
            _write(tmp_xlsx, {"Pivot!A3": {"style": style}})

    @pytest.mark.parametrize(
        ("layout", "compact", "outline"),
        [("compact", None, "1"), ("outline", "0", "1"), ("tabular", "0", None)],
    )
    def test_layout(self, tmp_xlsx: str, layout: str, compact: str | None, outline: str | None) -> None:
        """Each layout sets the compact and outline attributes the crate documents."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Region"], "layout": layout}})
        attrs = _only(tmp_xlsx).attrs
        assert (attrs.get("compact"), attrs.get("outline")) == (compact, outline)

    @pytest.mark.parametrize(
        ("option", "xml_attribute"),
        # Excel's UI and XML names are swapped: the UI's row grand totals are the
        # "Grand Total" column, which the XML calls colGrandTotals.
        [("row_grand_totals", "colGrandTotals"), ("column_grand_totals", "rowGrandTotals")],
    )
    def test_grand_totals_off(self, tmp_xlsx: str, option: str, xml_attribute: str) -> None:
        """Each grand total switches off alone, under Excel's UI name."""
        _write(tmp_xlsx, {"Pivot!A3": {option: False}})
        attrs = _only(tmp_xlsx).attrs
        assert {k: v for k, v in attrs.items() if k.endswith("GrandTotals")} == {xml_attribute: "0"}


class TestPivotTablePlacement:
    """Any sheet, any cell outside the frame's data."""

    def test_on_the_frames_own_sheet(self, tmp_xlsx: str) -> None:
        """A bare cell is on the frame's sheet, and no sheet is added."""
        _write(tmp_xlsx, {"F3": {}}, sheet_name="Data")
        pivot = _only(tmp_xlsx)
        assert (pivot.sheet, _anchor(pivot)) == ("Data", "F3")
        assert _sheet_names(tmp_xlsx) == ["Data"]

    def test_a_missing_sheet_is_created_after_the_others(self, tmp_xlsx: str) -> None:
        """A sheet no entry names is created, after every entry's sheet."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "Data", {"pivot_tables": {"Summary!B4": {}}}), (_frame(), "More")],
            tmp_xlsx,
        )
        pivot = _only(tmp_xlsx)
        assert (pivot.sheet, _anchor(pivot)) == ("Summary", "B4")
        assert _sheet_names(tmp_xlsx) == ["Data", "More", "Summary"]

    def test_an_existing_sheet_is_used(self, tmp_xlsx: str) -> None:
        """A later entry's sheet takes the pivot table; no new sheet appears."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "Data", {"pivot_tables": {"More!H2": {}}}), (_frame(), "More")],
            tmp_xlsx,
        )
        pivot = _only(tmp_xlsx)
        assert (pivot.sheet, pivot.source_sheet) == ("More", "Data")
        assert _sheet_names(tmp_xlsx) == ["Data", "More"]

    def test_an_existing_sheet_is_matched_without_case(self, tmp_xlsx: str) -> None:
        """Excel's sheet names are case-insensitive, so 'more' is the sheet 'More'."""
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "Data", {"pivot_tables": {"more!H2": {}}}), (_frame(), "More")],
            tmp_xlsx,
        )
        assert _only(tmp_xlsx).sheet == "More"
        assert _sheet_names(tmp_xlsx) == ["Data", "More"]

    def test_two_frames_share_one_new_sheet(self, tmp_xlsx: str) -> None:
        """The first mention creates the sheet; the second uses it."""
        xlsxturbo.dfs_to_xlsx(
            [
                (_frame(), "A", {"pivot_tables": {"Summary!A3": {}}}),
                (_frame(), "B", {"pivot_tables": {"Summary!H3": {}}}),
            ],
            tmp_xlsx,
        )
        assert sorted((p.source_sheet, _anchor(p)) for p in _pivots(tmp_xlsx)) == [("A", "A3"), ("B", "H3")]
        assert _sheet_names(tmp_xlsx) == ["A", "B", "Summary"]

    def test_several_pivot_tables_from_one_frame(self, tmp_xlsx: str) -> None:
        """Each key is one pivot table, written in key order."""
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Region"]}, "Pivot!E3": {"rows": ["Item"]}})
        assert [_anchor(p) for p in _pivots(tmp_xlsx)] == ["A3", "E3"]

    @pytest.mark.parametrize("location", ["A1", "B3", "D5", "Data!C2", "data!C2"])
    def test_inside_the_frame_is_refused(self, tmp_xlsx: str, location: str) -> None:
        """The frame's header and data cells are A1:D5; an anchor there is refused."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=location.split("!")[-1]):
            _write(tmp_xlsx, {location: {}}, sheet_name="Data")

    def test_inside_another_frame_on_the_target_sheet_is_refused(self, tmp_xlsx: str) -> None:
        """The target sheet's own frame is protected too, though it is written later."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="B2"):
            xlsxturbo.dfs_to_xlsx(
                [(_frame(), "Data", {"pivot_tables": {"More!B2": {}}}), (_frame(), "More")],
                tmp_xlsx,
            )

    def test_just_outside_the_frame_is_accepted(self, tmp_xlsx: str) -> None:
        """The first free column and the first free row are both allowed."""
        _write(tmp_xlsx, {"E1": {}, "A6": {}})
        assert sorted(_anchor(p) for p in _pivots(tmp_xlsx)) == ["A6", "E1"]

    @pytest.mark.parametrize("location", ["", "A0", "XFE1", "A1048577", "A1:B2", "Pivot!", "!A3", "Bad[Name]!A3"])
    def test_bad_location_is_refused(self, tmp_xlsx: str, location: str) -> None:
        """A location that is not one cell on a valid sheet is refused, naming the option."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="pivot_tables"):
            _write(tmp_xlsx, {location: {}})

    def test_the_same_location_twice_is_refused(self, tmp_xlsx: str) -> None:
        """Two pivot tables at one cell, even spelled with another sheet case, are refused."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="A3"):
            xlsxturbo.dfs_to_xlsx(
                [
                    (_frame(), "A", {"pivot_tables": {"Summary!A3": {}}}),
                    (_frame(), "B", {"pivot_tables": {"summary!A3": {}}}),
                ],
                tmp_xlsx,
            )


class TestPivotTableRefusals:
    """What would give a broken or empty pivot table is refused before anything is written."""

    @pytest.mark.parametrize("key", ["rows", "columns", "filters", "values"])
    def test_unknown_column_is_refused_with_the_column_list(self, tmp_xlsx: str, key: str) -> None:
        """The message names the missing column and the columns that exist."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"Nope.*Region"):
            _write(tmp_xlsx, {"Pivot!A3": {key: ["Nope"]}})

    def test_column_names_are_matched_exactly(self, tmp_xlsx: str) -> None:
        """'region' is not 'Region'; the crate's own lookup is exact."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="region"):
            _write(tmp_xlsx, {"Pivot!A3": {"rows": ["region"]}})

    def test_unknown_key_is_refused(self, tmp_xlsx: str) -> None:
        """A misspelt key is refused rather than ignored."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="row_fields"):
            _write(tmp_xlsx, {"Pivot!A3": {"row_fields": ["Region"]}})

    def test_unknown_value_field_key_is_refused(self, tmp_xlsx: str) -> None:
        """A misspelt key inside a value field is refused too."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="func"):
            _write(tmp_xlsx, {"Pivot!A3": {"values": [{"field": "Volume", "func": "sum"}]}})

    def test_value_field_without_field_is_refused(self, tmp_xlsx: str) -> None:
        """A value field dict must name its column."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="field"):
            _write(tmp_xlsx, {"Pivot!A3": {"values": [{"function": "sum"}]}})

    def test_unknown_function_is_refused(self, tmp_xlsx: str) -> None:
        """Excel has no median in a pivot table, so it is refused by name."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="median"):
            _write(tmp_xlsx, {"Pivot!A3": {"values": [{"field": "Volume", "function": "median"}]}})

    def test_unknown_layout_is_refused(self, tmp_xlsx: str) -> None:
        """Only compact, outline and tabular exist."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="flat"):
            _write(tmp_xlsx, {"Pivot!A3": {"layout": "flat"}})

    def test_a_string_where_a_list_belongs_is_refused(self, tmp_xlsx: str) -> None:
        """A str would read as one-letter column names; refusing it beats "column 'R' not found"."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="rows"):
            _write(tmp_xlsx, {"Pivot!A3": {"rows": "Region"}})

    def test_repeated_value_caption_is_refused(self, tmp_xlsx: str) -> None:
        """Two value fields with one caption are refused, as the crate does."""
        values = [{"field": "Volume", "name": "Total"}, {"field": "Volume", "function": "count", "name": "Total"}]
        with pytest.raises(xlsxturbo.ConfigurationError, match="Total"):
            _write(tmp_xlsx, {"Pivot!A3": {"values": values}})

    def test_without_a_header_row_is_refused(self, tmp_xlsx: str) -> None:
        """Without a header row the source has no field names."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="header"):
            _write(tmp_xlsx, {"Pivot!A3": {}}, header=False)

    def test_an_empty_frame_is_refused(self, tmp_xlsx: str) -> None:
        """A source of only a header row is refused, as the crate does."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="row"):
            _write(tmp_xlsx, {"Pivot!A3": {}}, frame=_frame().iloc[:0])

    def test_duplicate_column_names_are_refused(self, tmp_xlsx: str) -> None:
        """Repeated field names give Excel's repair prompt, so they are refused by name."""
        frame = _frame()
        frame.columns = ["Region", "Item", "Volume", "Item"]
        with pytest.raises(xlsxturbo.ConfigurationError, match="Item"):
            _write(tmp_xlsx, {"Pivot!A3": {}}, frame=frame)

    def test_column_names_equal_without_case_are_refused(self, tmp_xlsx: str) -> None:
        """The crate refuses field names equal without case, as measured on the upstream branch."""
        frame = _frame()
        frame.columns = ["Region", "Item", "Volume", "ITEM"]
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"(?i)item"):
            _write(tmp_xlsx, {"Pivot!A3": {}}, frame=frame)

    def test_an_empty_column_name_is_refused(self, tmp_xlsx: str) -> None:
        """A blank field name gives Excel's repair prompt, so it is refused."""
        frame = _frame()
        frame.columns = ["Region", "", "Volume", "Month"]
        with pytest.raises(xlsxturbo.ConfigurationError, match="empty"):
            _write(tmp_xlsx, {"Pivot!A3": {}}, frame=frame)

    def test_a_refusal_leaves_the_target_untouched(self, tmp_xlsx: str) -> None:
        """The checks run before the save, so the empty file the fixture made stays empty."""
        with pytest.raises(xlsxturbo.ConfigurationError):
            _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Nope"]}})
        assert Path(tmp_xlsx).stat().st_size == 0


class TestPivotTableConstantMemory:
    """The crate cannot read a constant-memory header row, so xlsxturbo passes the names."""

    def test_written_with_the_frames_column_names(self, tmp_xlsx: str) -> None:
        """The pivot table is written, with no skip warning, and names every column."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            pivot_tables = {"Pivot!A3": {"rows": ["Region"], "values": ["Volume"]}}
            _write(tmp_xlsx, pivot_tables, sheet_name="Data", constant_memory=True)
        pivot = _only(tmp_xlsx)
        assert pivot.cache_fields == ["Region", "Item", "Volume", "Month"]
        assert (pivot.source_sheet, pivot.source_ref) == ("Data", "A1:D5")


class TestPivotTableWithPolars:
    """A polars frame gives the same definition as the pandas one."""

    def test_polars_frame(self, tmp_xlsx: str) -> None:
        """Column names come from the polars schema."""
        pl = pytest.importorskip("polars")
        _write(tmp_xlsx, {"Pivot!A3": {"rows": ["Region"]}}, frame=pl.from_pandas(_frame()))
        assert _only(tmp_xlsx).cache_fields == ["Region", "Item", "Volume", "Month"]
