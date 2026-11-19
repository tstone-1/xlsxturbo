"""Strings that rust_xlsxwriter writes into XML attributes and formulas, escaped as Excel needs.

Three upstream defects that xlsxturbo has no guard for, each fixed in rust_xlsxwriter 0.103.0:

- #216: a control character in a string the crate writes as an XML attribute (sheet name,
  font name, number format, validation title, textbox alt text) gave XML that is not
  well-formed. 0.103.0 writes it as ``_x0001_``, as it always did in cell text.
- #217: a sheet or table named with literal ``_xHHHH_`` text was written raw, so Excel read
  the name as an escaped character. 0.103.0 escapes the underscore (``_x005f_x000D_``).
- Upstream commit ``0c034f85b``: a sheet named ``TRUE`` or ``FALSE`` (any case) was not
  quoted in the autofilter's defined name, so ``TRUE!$A$1:$B$3`` reached Excel and the
  workbook did not open. 0.103.0 quotes it.

The tests read the package XML; whether Excel opens each file is checked with
``scripts/excel_open_check.ps1``, not here. The measurements and the issue programs are in
``dev-docs/upstream-rust-xlsxwriter.md``.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import pandas as pd
import pytest
import xlsxturbo

CONTROL = "\x01"


def _frame() -> pd.DataFrame:
    """Two columns, two rows."""
    return pd.DataFrame({"a": [1, 2], "b": [3, 4]})


def _xml_parts(path: Path) -> dict[str, str]:
    """Every XML part of the package, decoded."""
    with zipfile.ZipFile(path) as package:
        return {n: package.read(n).decode("utf-8") for n in package.namelist() if n.endswith((".xml", ".rels", ".vml"))}


def _defined_names(path: Path) -> list[str]:
    """The text of every ``<definedName>`` in the workbook part."""
    return re.findall(r"<definedName[^>]*>([^<]*)</definedName>", _xml_parts(path)["xl/workbook.xml"])


class TestControlCharacterInAnAttribute:
    """Upstream #216: every part stays well-formed XML."""

    @pytest.mark.parametrize(
        "options",
        [
            {"sheet_name": f"a{CONTROL}b"},
            {"header_format": {"font_name": f"Ari{CONTROL}al"}},
            {"column_formats": {"b": {"num_format": f"0{CONTROL}"}}},
            {"validations": {"b": {"type": "whole_number", "min": 0, "input_title": f"t{CONTROL}"}}},
            {"textboxes": {"D2": {"text": "x", "alt_text": f"t{CONTROL}"}}},
        ],
        ids=["sheet_name", "font_name", "num_format", "validation_title", "textbox_alt_text"],
    )
    def test_every_part_is_well_formed(self, tmp_path: Path, options: dict[str, Any]) -> None:
        """0.102.0 wrote U+0001 raw into the attribute; every one of these parts failed to parse."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), **options)
        for part, text in _xml_parts(path).items():
            ElementTree.fromstring(text)  # noqa: S314 - our own output, not untrusted input
            assert CONTROL not in text, part

    def test_control_the_same_character_in_a_cell(self, tmp_path: Path) -> None:
        """Control: in cell text the character was always escaped, so this passed on 0.102.0 too."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [f"x{CONTROL}y"]}), str(path))
        assert any("x_x0001_y" in text for text in _xml_parts(path).values())


class TestNullInASheetName:
    """U+0000 is not escaped but refused: Excel allows no null character in a sheet name."""

    def test_it_raises_configuration_error(self, tmp_path: Path) -> None:
        """0.102.0 wrote it raw and gave a broken file; 0.103.0 refuses the name."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="null character"):
            xlsxturbo.df_to_xlsx(_frame(), str(tmp_path / "t.xlsx"), sheet_name="a\x00b")


class TestLiteralEscapeShapeInAName:
    """Upstream #217: a name that is literally ``_x000D_`` keeps that text in Excel."""

    def test_sheet_name_attribute(self, tmp_path: Path) -> None:
        """The sheet's name attribute escapes the underscore, as Excel writes it."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), sheet_name="_x000D_")
        assert 'name="_x005f_x000D_"' in _xml_parts(path)["xl/workbook.xml"]

    def test_sheet_name_in_the_autofilter_defined_name(self, tmp_path: Path) -> None:
        """The filter's defined name refers to the sheet by its escaped name."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), sheet_name="_x000D_", autofilter=True)
        (filter_range,) = _defined_names(path)
        assert re.search(r"(?<!_x005F)(?<!_x005f)_x000D_", filter_range) is None, filter_range
        assert filter_range.endswith("!$A$1:$B$3")

    def test_table_name_attribute(self, tmp_path: Path) -> None:
        """A table named ``_x000D_`` keeps that name; 0.102.0 wrote it raw."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), table_style="Medium2", table_name="_x000D_")
        table = _xml_parts(path)["xl/tables/table1.xml"]
        assert 'name="_x005f_x000D_"' in table
        assert 'displayName="_x005f_x000D_"' in table

    def test_control_a_name_without_the_shape(self, tmp_path: Path) -> None:
        """Control: ``_X000D_`` (upper-case X) is not an escape and is written as it is."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), sheet_name="_X000D_")
        assert 'name="_X000D_"' in _xml_parts(path)["xl/workbook.xml"]


class TestBooleanSheetNameIsQuoted:
    """Upstream ``0c034f85b``: Excel reads an unquoted TRUE or FALSE as a boolean, not a sheet."""

    @pytest.mark.parametrize("name", ["TRUE", "FALSE", "true", "False"])
    def test_autofilter_defined_name_quotes_it(self, tmp_path: Path, name: str) -> None:
        """The filter range names the sheet in quotes."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), sheet_name=name, autofilter=True)
        assert _defined_names(path) == [f"'{name}'!$A$1:$B$3"]

    @pytest.mark.parametrize("name", ["TRUE1", "FALSEY", "TRUE_", "Data"])
    def test_control_other_names_stay_unquoted(self, tmp_path: Path, name: str) -> None:
        """Control: a name that only starts like a boolean needs no quotes, then or now."""
        path = tmp_path / "t.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), str(path), sheet_name=name, autofilter=True)
        assert _defined_names(path) == [f"{name}!$A$1:$B$3"]
