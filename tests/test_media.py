"""Tests for media features: images, checkboxes, textboxes, and charts."""

from __future__ import annotations

import base64
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import TINY_PNG_B64, active_ws, load_workbook

if TYPE_CHECKING:
    from xlsxturbo.types import (
        ChartOptions,
        CheckboxOptions,
        CommentOptions,
        ImageOptions,
        SparklineOptions,
        TextboxOptions,
    )

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

# A drawing measures in EMU; one pixel at 96 dpi is 9525 EMU.
EMU_PER_PX = 9525
# The default column width (8.43 characters) and row height (15 points), in pixels.
# A drawing anchor counts whole cells plus an offset, so the anchor converts to a
# pixel position only across columns and rows left at these defaults.
DEFAULT_COL_PX = 64
DEFAULT_ROW_PX = 20


def _drawing(path: str, number: int = 1) -> ElementTree.Element:
    """Return ``xl/drawings/drawing<number>.xml`` of a written workbook, parsed."""
    with zipfile.ZipFile(path) as zf:
        return ElementTree.fromstring(zf.read(f"xl/drawings/drawing{number}.xml"))  # noqa: S314 - our own output


def _sheet_drawing(path: str, sheet: int) -> ElementTree.Element:
    """Return the drawing that worksheet ``sheet`` (1-based) points at, parsed.

    Followed through the sheet's relationships rather than assumed from the part
    number, so a test about which sheet got which shape cannot pass on a
    numbering coincidence.
    """
    with zipfile.ZipFile(path) as zf:
        rels = ElementTree.fromstring(zf.read(f"xl/worksheets/_rels/sheet{sheet}.xml.rels"))  # noqa: S314
        targets = [
            rel.attrib["Target"]
            for rel in rels
            if rel.attrib.get("Type", "").endswith("/drawing")
        ]
        assert len(targets) == 1, f"sheet{sheet} has drawings {targets}"
        name = "xl/drawings/" + targets[0].rsplit("/", 1)[-1]
        return ElementTree.fromstring(zf.read(name))  # noqa: S314 - our own output


def _anchor_px(drawing: ElementTree.Element) -> tuple[int, int, int, int]:
    """Pixel position of the first anchor's two corners: ``(x_from, y_from, x_to, y_to)``.

    Valid only where every column and row the anchor spans has the default size.
    """
    anchor = drawing.find("xdr:twoCellAnchor", NS)
    assert anchor is not None, "the drawing has no two-cell anchor"

    def corner(tag: str) -> tuple[int, int]:
        def number(field: str) -> int:
            node = anchor.find(f"xdr:{tag}/xdr:{field}", NS)
            assert node is not None, f"no {tag}/{field}"
            assert node.text is not None, f"empty {tag}/{field}"
            return int(node.text)

        x = number("col") * DEFAULT_COL_PX + number("colOff") // EMU_PER_PX
        y = number("row") * DEFAULT_ROW_PX + number("rowOff") // EMU_PER_PX
        return x, y

    x_from, y_from = corner("from")
    x_to, y_to = corner("to")
    return x_from, y_from, x_to, y_to


def _shape_extent_px(drawing: ElementTree.Element) -> tuple[int, int]:
    """Width and height of the first shape or picture, from its ``a:ext``, in pixels."""
    ext = drawing.find(".//a:xfrm/a:ext", NS)
    assert ext is not None, "the drawing has no a:ext"
    return int(ext.attrib["cx"]) // EMU_PER_PX, int(ext.attrib["cy"]) // EMU_PER_PX


def _write_png(factory: Callable[..., str]) -> str:
    """Write the one-pixel PNG to a fresh temporary path and return the path."""
    img_path = factory(".png")
    Path(img_path).write_bytes(base64.b64decode(TINY_PNG_B64))
    return img_path


class TestImages:
    """Tests for images feature (v0.10.0)."""

    def test_image_simple_path(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """Image with simple path."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = tmp_xlsx_factory()
        png_data = base64.b64decode(TINY_PNG_B64)
        img_path = tmp_xlsx_factory(".png")
        with Path(img_path).open("wb") as f:
            f.write(png_data)

        xlsxturbo.df_to_xlsx(df, path, images={"D1": img_path})
        # The image must actually land in the package, not just produce a file.
        with zipfile.ZipFile(path) as zf:
            media = [n for n in zf.namelist() if n.startswith("xl/media/")]
            assert media, "no embedded image found in xl/media/"
            assert any(n.endswith(".png") for n in media)

    def test_image_with_options(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """Image scale_width and scale_height reach the picture's size.

        The two scales differ, so a width applied as the height (or one scale
        applied to both) fails too. The image is one pixel square.
        """
        df = pd.DataFrame({"A": [1]})
        path = tmp_xlsx_factory()
        img_path = _write_png(tmp_xlsx_factory)

        xlsxturbo.df_to_xlsx(
            df,
            path,
            images={"B5": {"path": img_path, "scale_width": 2.0, "scale_height": 3.0}},
        )
        with zipfile.ZipFile(path) as zf:
            media = [n for n in zf.namelist() if n.startswith("xl/media/")]
            assert media, "no embedded image found in xl/media/"
        assert _shape_extent_px(_drawing(path)) == (2, 3)

    def test_image_without_options_keeps_its_size(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """The control for the scale test: unscaled, the one-pixel image stays one pixel."""
        path = tmp_xlsx_factory()
        xlsxturbo.df_to_xlsx(pd.DataFrame({"A": [1]}), path, images={"B5": _write_png(tmp_xlsx_factory)})
        assert _shape_extent_px(_drawing(path)) == (1, 1)

    def test_image_alt_text_reaches_the_drawing(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """alt_text becomes the picture's ``descr``, which screen readers announce."""
        path = tmp_xlsx_factory()
        xlsxturbo.df_to_xlsx(
            pd.DataFrame({"A": [1]}),
            path,
            images={"B5": {"path": _write_png(tmp_xlsx_factory), "alt_text": "Quarterly logo"}},
        )
        props = _drawing(path).find(".//xdr:pic/xdr:nvPicPr/xdr:cNvPr", NS)
        assert props is not None
        assert props.attrib.get("descr") == "Quarterly logo"

    def test_image_unknown_key_raises(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """A typo in an image option is refused with the valid keys, not dropped."""
        bad: Any = {"path": _write_png(tmp_xlsx_factory), "scale": 2.0}
        with pytest.raises(
            xlsxturbo.ConfigurationError,
            match=r"images\['B5'\]: unknown option 'scale'\. Valid: scale_width, scale_height, alt_text",
        ):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"A": [1]}), tmp_xlsx_factory(), images={"B5": bad})


class TestCheckboxes:
    """Tests for checkboxes feature (v0.13.0)."""

    def test_checkbox_simple_bool(self, tmp_xlsx: str) -> None:
        """Checkboxes with plain bool values."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"D1": True, "D2": False})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Checkboxes render as boolean TRUE/FALSE in cells
        assert ws["D1"].value is True
        assert ws["D2"].value is False

    def test_checkbox_dict_form(self, tmp_xlsx: str) -> None:
        """Checkbox with dict specifying checked state."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": {"checked": True}})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["B2"].value is True

    def test_checkbox_with_format(self, tmp_xlsx: str) -> None:
        """Checkbox with optional cell format."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            checkboxes={"C3": {"checked": True, "format": {"bg_color": "#C6EFCE", "bold": True}}},
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["C3"].value is True
        assert ws["C3"].font.bold is True
        assert ws["C3"].fill.fgColor.rgb == "FFC6EFCE"

    def test_checkbox_missing_checked_key_raises(self, tmp_xlsx: str) -> None:
        """Dict form without 'checked' raises a clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="checked"):
            # Deliberately omits the required 'checked' key.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": {"format": {"bold": True}}})  # type: ignore[dict-item]

    def test_checkbox_format_not_dict_raises(self, tmp_xlsx: str) -> None:
        """'format' field present but not a dict raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                # Intentionally invalid: 'format' must be a dict.
                checkboxes={"B2": {"checked": True, "format": "not a dict"}},  # type: ignore[dict-item]
            )
        message = str(exc_info.value)
        assert "format" in message
        assert "dict" in message

    def test_checkbox_dict_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A typo'd/extra key in the checkbox dict form is rejected, not silently dropped."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="unknown option 'label'"):
            # Intentionally invalid: 'label' is not a recognized checkbox key.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": {"checked": True, "label": "x"}})  # type: ignore[typeddict-unknown-key]

    def test_checkbox_invalid_cell_ref_raises(self, tmp_xlsx: str) -> None:
        """Invalid cell reference raises."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"checkboxes\['not_a_ref'\]: Invalid row number in cell reference"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"not_a_ref": True})

    def test_checkbox_wrong_value_type_raises(self, tmp_xlsx: str) -> None:
        """Non-bool, non-dict value raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="checkboxes"):
            # Intentionally invalid: value must be a bool or dict.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": "not_a_bool"})  # type: ignore[dict-item]

    def test_checkbox_with_dfs_to_xlsx_per_sheet(self, tmp_xlsx: str) -> None:
        """Checkboxes work via per-sheet options in dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": [1, 2]})
        df2 = pd.DataFrame({"B": [3, 4]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "S1", {"checkboxes": {"D1": True}}),
                (df2, "S2", {"checkboxes": {"D1": False}}),
            ],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        assert wb["S1"]["D1"].value is True
        assert wb["S2"]["D1"].value is False

    def test_checkbox_combined_with_other_features(self, tmp_xlsx: str) -> None:
        """Checkboxes coexist with other features on the same sheet."""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Active": [True, False]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            checkboxes={"D1": True, "D2": False},
            comments={"A1": "Names"},
            validations={"Name": {"type": "text_length", "min": 1, "max": 100}},
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert (ws["D1"].value, ws["D2"].value) == (True, False)
        comment = ws["A1"].comment
        assert comment is not None
        assert "Names" in comment.text
        assert [dv.type for dv in ws.data_validations.dataValidation] == ["textLength"]


class TestTextboxes:
    """Tests for textboxes feature (v0.14.0)."""

    def test_textbox_simple_string(self, tmp_xlsx: str) -> None:
        """Textbox with bare string value writes file successfully."""
        df = pd.DataFrame({"A": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"C2": "Simple note"})
        with zipfile.ZipFile(tmp_xlsx) as zf:
            assert "xl/drawings/drawing1.xml" in zf.namelist()
            drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
            assert "Simple note" in drawing

    def test_textbox_dict_form(self, tmp_xlsx: str) -> None:
        """Textbox dict form with required 'text' key."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "Hello"}})
        texts = [node.text for node in _drawing(tmp_xlsx).iterfind(".//xdr:txBody//a:t", NS)]
        assert texts == ["Hello"]

    def test_textbox_with_all_options(self, tmp_xlsx: str) -> None:
        """Every textbox option reaches the drawing.

        Size and offsets are read from the shape's ``a:ext`` and its anchor,
        the font from the run properties, and the line colour from ``a:ln``.
        The control below shows that none of the asserted values is a default.
        """
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            textboxes={
                "B2": {
                    "text": "Annotated",
                    "width": 200,
                    "height": 100,
                    "x_offset": 10,
                    "y_offset": 5,
                    "font": {
                        "name": "Arial",
                        "size": 14,
                        "bold": True,
                        "italic": True,
                        "underline": True,
                        "color": "#FF0000",
                    },
                    "fill_color": "#F0F0F0",
                    "line_color": "#00B050",
                    "alt_text": "A textbox annotation",
                }
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
            assert "Annotated" in drawing
            # alt_text, fill color, and font color must reach the XML.
            assert "A textbox annotation" in drawing
            assert "F0F0F0" in drawing.upper()
            assert "FF0000" in drawing.upper()

        root = _drawing(tmp_xlsx)
        assert _shape_extent_px(root) == (200, 100)
        x_from, y_from, x_to, y_to = _anchor_px(root)
        # B2 is column 1, row 1; the offsets move the top-left corner inside it.
        assert (x_from, y_from) == (1 * DEFAULT_COL_PX + 10, 1 * DEFAULT_ROW_PX + 5)
        assert (x_to - x_from, y_to - y_from) == (200, 100)

        run = root.find(".//xdr:txBody//a:rPr", NS)
        assert run is not None
        assert run.attrib.get("sz") == "1400"
        assert run.attrib.get("b") == "1"
        assert run.attrib.get("i") == "1"
        assert run.attrib.get("u") == "sng"
        latin = run.find("a:latin", NS)
        assert latin is not None
        assert latin.attrib.get("typeface") == "Arial"

        line = root.find(".//xdr:spPr/a:ln/a:solidFill/a:srgbClr", NS)
        assert line is not None, "line_color did not reach the shape outline"
        assert line.attrib.get("val") == "00B050"

    def test_textbox_without_options_has_the_default_look(self, tmp_xlsx: str) -> None:
        """The control for the test above: none of what it asserts is a default.

        A plain textbox is 192 x 120 pixels at the cell's corner, with an 11 point
        run that is neither bold, italic nor underlined, and a theme-coloured
        outline. If any of these were the asserted value, the test above could
        not tell an option that arrived from one that was dropped.
        """
        xlsxturbo.df_to_xlsx(pd.DataFrame({"A": [1]}), tmp_xlsx, textboxes={"B2": "plain"})
        root = _drawing(tmp_xlsx)
        assert _shape_extent_px(root) == (192, 120)
        assert _anchor_px(root)[:2] == (1 * DEFAULT_COL_PX, 1 * DEFAULT_ROW_PX)
        run = root.find(".//xdr:txBody//a:rPr", NS)
        assert run is not None
        assert run.attrib.get("sz") == "1100"
        assert not {"b", "i", "u"} & set(run.attrib)
        assert run.find("a:latin", NS) is None
        assert root.find(".//xdr:spPr/a:ln/a:solidFill/a:srgbClr", NS) is None

    def test_textbox_font_partial(self, tmp_xlsx: str) -> None:
        """Partial font options work (only some keys set)."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "font": {"bold": True}}})
        run = _drawing(tmp_xlsx).find(".//xdr:txBody//a:rPr", NS)
        assert run is not None
        assert run.attrib.get("b") == "1"
        # Only the key that was set: no italic, no underline, the default size.
        assert "i" not in run.attrib
        assert "u" not in run.attrib
        assert run.attrib.get("sz") == "1100"

    def test_textbox_missing_text_raises(self, tmp_xlsx: str) -> None:
        """Dict form without 'text' raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"textboxes\['B2'\] dict missing 'text' key"):
            # Deliberately omits the required 'text' key.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"width": 100}})  # type: ignore[dict-item]

    def test_textbox_unknown_option_raises(self, tmp_xlsx: str) -> None:
        """Unknown top-level option raises with the list of valid keys."""
        df = pd.DataFrame({"A": [1]})
        # Intentionally invalid: 'bogus' is not a valid textbox option.
        textboxes: dict[str, str | TextboxOptions] = {
            "B2": {"text": "T", "bogus": 1}  # type: ignore[typeddict-unknown-key]
        }
        with pytest.raises(ValueError) as exc_info:  # noqa: PT011
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes=textboxes)
        message = str(exc_info.value)
        assert "bogus" in message
        assert "Valid" in message

    def test_textbox_unknown_font_option_raises(self, tmp_xlsx: str) -> None:
        """Unknown font sub-option raises."""
        df = pd.DataFrame({"A": [1]})
        # Intentionally invalid: 'weight' is not a valid font option.
        textboxes: dict[str, str | TextboxOptions] = {
            "B2": {"text": "T", "font": {"weight": "heavy"}}  # type: ignore[typeddict-unknown-key]
        }
        with pytest.raises(ValueError) as exc_info:  # noqa: PT011
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes=textboxes)
        message = str(exc_info.value)
        assert "weight" in message
        assert "font" in message.lower()

    def test_textbox_font_not_dict_raises(self, tmp_xlsx: str) -> None:
        """'font' present but not a dict raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        # Intentionally invalid: 'font' must be a dict.
        textboxes: dict[str, str | TextboxOptions] = {
            "B2": {"text": "T", "font": "bold"}  # type: ignore[dict-item]
        }
        with pytest.raises(ValueError) as exc_info:  # noqa: PT011
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes=textboxes)
        message = str(exc_info.value)
        assert "font" in message
        assert "dict" in message.lower()

    def test_textbox_offsets_only(self, tmp_xlsx: str) -> None:
        """Offsets alone move the textbox's corner and leave its size alone.

        The 25-pixel ``y_offset`` is larger than the 20-pixel row, so the anchor
        rolls over into row 3; the pixel position is what must come out right.
        """
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "x_offset": 15, "y_offset": 25}})
        root = _drawing(tmp_xlsx)
        assert _anchor_px(root)[:2] == (1 * DEFAULT_COL_PX + 15, 1 * DEFAULT_ROW_PX + 25)
        assert _shape_extent_px(root) == (192, 120)

    def test_textbox_wrong_value_type_raises(self, tmp_xlsx: str) -> None:
        """Non-string, non-dict value raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="textboxes"):
            # Intentionally invalid: value must be a string or dict.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": 123})  # type: ignore[dict-item]

    def test_textbox_invalid_color_raises(self, tmp_xlsx: str) -> None:
        """Invalid color string surfaces the parse error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"textboxes\['B2'\]: 'fill_color': Unknown color: not_a_color"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "fill_color": "not_a_color"}})

    def test_textbox_invalid_cell_ref_raises(self, tmp_xlsx: str) -> None:
        """Invalid cell reference raises."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"textboxes\['not_a_ref'\]: Invalid row number in cell reference"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"not_a_ref": "T"})

    def test_textbox_with_dfs_to_xlsx_per_sheet(self, tmp_xlsx: str) -> None:
        """Textboxes work via per-sheet options in dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "S1", {"textboxes": {"C2": "Sheet 1 note"}}),
                (df2, "S2", {"textboxes": {"C2": {"text": "Sheet 2", "font": {"bold": True}}}}),
            ],
            tmp_xlsx,
        )
        first, second = _sheet_drawing(tmp_xlsx, 1), _sheet_drawing(tmp_xlsx, 2)
        assert [t.text for t in first.iterfind(".//a:t", NS)] == ["Sheet 1 note"]
        assert [t.text for t in second.iterfind(".//a:t", NS)] == ["Sheet 2"]
        first_run, second_run = first.find(".//a:rPr", NS), second.find(".//a:rPr", NS)
        assert first_run is not None
        assert second_run is not None
        # Each sheet keeps its own font: the bold is S2's alone.
        assert "b" not in first_run.attrib
        assert second_run.attrib.get("b") == "1"

    def test_textbox_combined_with_other_features(self, tmp_xlsx: str) -> None:
        """Textboxes coexist with images, checkboxes, comments on the same sheet."""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [85, 92]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            textboxes={"D2": {"text": "Notes", "width": 150, "height": 60}},
            checkboxes={"E1": True},
            comments={"A1": "Names column"},
        )
        root = _drawing(tmp_xlsx)
        assert [t.text for t in root.iterfind(".//a:t", NS)] == ["Notes"]
        assert _shape_extent_px(root) == (150, 60)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert ws["E1"].value is True
        comment = ws["A1"].comment
        assert comment is not None
        assert "Names column" in comment.text


class TestCharts:
    """Tests for native Excel charts."""

    def test_single_series_chart(self, tmp_xlsx: str) -> None:
        """Charts create native chart XML with title and data table."""
        df = pd.DataFrame({"Month": ["Jan", "Feb", "Mar"], "Sales": [120, 145, 160]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            charts={
                "D2": {
                    "type": "bar",
                    "data_range": "Sheet1!$B$2:$B$4",
                    "categories_range": "Sheet1!$A$2:$A$4",
                    "title": "Monthly Sales",
                    "width": 720,
                    "height": 480,
                    "show_data_table": True,
                    "legend_position": "bottom",
                }
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            names = zf.namelist()
            assert "xl/charts/chart1.xml" in names
            chart_xml = zf.read("xl/charts/chart1.xml").decode("utf-8")
            assert "Monthly Sales" in chart_xml
            assert "<c:dTable>" in chart_xml
            assert "<c:barChart>" in chart_xml
            # The data_range and categories_range must actually land in the
            # chart XML, not just produce a well-formed but empty chart.
            assert "Sheet1!$B$2:$B$4" in chart_xml
            assert "Sheet1!$A$2:$A$4" in chart_xml

    @staticmethod
    def _chart(path: str) -> ElementTree.Element:
        """Return ``xl/charts/chart1.xml`` of a written workbook, parsed."""
        with zipfile.ZipFile(path) as zf:
            return ElementTree.fromstring(zf.read("xl/charts/chart1.xml"))  # noqa: S314 - our own output

    @staticmethod
    def _axis_title(chart: ElementTree.Element, axis: str) -> str | None:
        """The title text of the chart's ``c:catAx`` or ``c:valAx``, or None."""
        texts = [t.text or "" for t in chart.iterfind(f".//c:plotArea/c:{axis}/c:title//a:t", NS)]
        return "".join(texts) if texts else None

    def test_chart_geometry_legend_and_axis_names(self, tmp_xlsx: str) -> None:
        """width, height, both offsets, legend_position and both axis names reach the XML.

        A column chart, where ``x_axis_name`` titles the category axis; a bar
        chart turns the axes round. The size and the offsets are read from the
        drawing anchor, in pixels.
        """
        df = pd.DataFrame({"Month": ["Jan", "Feb", "Mar"], "Sales": [120, 145, 160]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            charts={
                "D2": {
                    "type": "column",
                    "data_range": "Sheet1!$B$2:$B$4",
                    "width": 720,
                    "height": 480,
                    "x_offset": 15,
                    "y_offset": 25,
                    "legend_position": "bottom",
                    "x_axis_name": "Month",
                    "y_axis_name": "Units sold",
                }
            },
        )
        x_from, y_from, x_to, y_to = _anchor_px(_drawing(tmp_xlsx))
        # D2 is column 3, row 1.
        assert (x_from, y_from) == (3 * DEFAULT_COL_PX + 15, 1 * DEFAULT_ROW_PX + 25)
        assert (x_to - x_from, y_to - y_from) == (720, 480)

        chart = self._chart(tmp_xlsx)
        legend_pos = chart.find(".//c:legend/c:legendPos", NS)
        assert legend_pos is not None
        assert legend_pos.attrib.get("val") == "b"
        assert self._axis_title(chart, "catAx") == "Month"
        assert self._axis_title(chart, "valAx") == "Units sold"

    def test_chart_defaults_are_not_what_the_test_above_asserts(self, tmp_xlsx: str) -> None:
        """The control: a chart with none of those options is 480 x 288 at the cell, legend right, no axis titles."""
        df = pd.DataFrame({"Month": ["Jan", "Feb", "Mar"], "Sales": [120, 145, 160]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, charts={"D2": {"type": "column", "data_range": "Sheet1!$B$2:$B$4"}}
        )
        x_from, y_from, x_to, y_to = _anchor_px(_drawing(tmp_xlsx))
        assert (x_from, y_from) == (3 * DEFAULT_COL_PX, 1 * DEFAULT_ROW_PX)
        assert (x_to - x_from, y_to - y_from) == (480, 288)
        chart = self._chart(tmp_xlsx)
        legend_pos = chart.find(".//c:legend/c:legendPos", NS)
        assert legend_pos is not None
        assert legend_pos.attrib.get("val") == "r"
        assert self._axis_title(chart, "catAx") is None
        assert self._axis_title(chart, "valAx") is None

    def test_chart_invalid_legend_position_raises(self, tmp_xlsx: str) -> None:
        """An unknown legend position is refused with the valid ones."""
        df = pd.DataFrame({"A": [1, 2]})
        charts: Any = {
            "D2": {"type": "bar", "data_range": "Sheet1!$A$2:$A$3", "legend_position": "middle"}
        }
        with pytest.raises(
            xlsxturbo.ConfigurationError,
            match=r"charts\['D2'\]: Unknown legend position 'middle'\. Valid: right, left, top, bottom, top_right",
        ):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_multi_series_chart(self, tmp_xlsx: str) -> None:
        """Charts support explicit multiple series."""
        df = pd.DataFrame(
            {
                "Month": ["Jan", "Feb", "Mar"],
                "Sales": [120, 145, 160],
                "Margin": [32, 41, 48],
            }
        )
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            charts={
                "E2": {
                    "type": "column",
                    "series": [
                        {"name": "Sales", "values_range": "Sheet1!$B$2:$B$4"},
                        {"name": "Margin", "values_range": "Sheet1!$C$2:$C$4"},
                    ],
                    "categories_range": "Sheet1!$A$2:$A$4",
                    "title": "Quarter Results",
                    "show_legend": False,
                }
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            chart_xml = zf.read("xl/charts/chart1.xml").decode("utf-8")
            assert chart_xml.count("<c:ser>") == 2
            assert "Quarter Results" in chart_xml
            assert "<c:legend>" not in chart_xml
            # Each series' values range and the shared categories range
            # must reach the XML.
            assert "Sheet1!$B$2:$B$4" in chart_xml
            assert "Sheet1!$C$2:$C$4" in chart_xml
            assert "Sheet1!$A$2:$A$4" in chart_xml

    def test_charts_with_dfs_to_xlsx_per_sheet(self, tmp_xlsx: str) -> None:
        """Charts work via per-sheet options in dfs_to_xlsx."""
        df1 = pd.DataFrame({"Month": ["Jan", "Feb"], "Sales": [10, 20]})
        df2 = pd.DataFrame({"Month": ["Jan", "Feb"], "Sales": [30, 40]})
        xlsxturbo.dfs_to_xlsx(
            [
                (
                    df1,
                    "North",
                    {
                        "charts": {
                            "D2": {
                                "type": "line",
                                "data_range": "North!$B$2:$B$3",
                                "categories_range": "North!$A$2:$A$3",
                            }
                        }
                    },
                ),
                (
                    df2,
                    "South",
                    {
                        "charts": {
                            "D2": {
                                "type": "line",
                                "data_range": "South!$B$2:$B$3",
                                "categories_range": "South!$A$2:$A$3",
                            }
                        }
                    },
                ),
            ],
            tmp_xlsx,
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            assert "xl/charts/chart1.xml" in zf.namelist()
            assert "xl/charts/chart2.xml" in zf.namelist()

    def test_chart_invalid_type_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid chart types raise a clear error."""
        df = pd.DataFrame({"A": [1, 2]})
        # Intentionally invalid: 'not_a_chart' is not a known chart type.
        charts: dict[str, ChartOptions] = {
            "D2": {"type": "not_a_chart", "data_range": "Sheet1!$A$2:$A$3"}  # type: ignore[typeddict-item]
        }
        with pytest.raises(ValueError, match="Unknown chart type"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    @pytest.mark.parametrize("style", [0, 49, 200, 300, -1])
    def test_chart_invalid_style_raises(self, style: int, tmp_xlsx: str) -> None:
        """A style outside 1-48 is rejected with the 1-48 message.

        Covers values a u8 would silently accept (0, 200) — Excel discards
        those and rust_xlsxwriter only complains on stderr, where Python
        cannot see it — as well as values beyond u8 (300) and negatives.
        """
        df = pd.DataFrame({"A": [1, 2]})
        charts: dict[str, ChartOptions] = {
            "D2": {"type": "bar", "data_range": "Sheet1!$A$2:$A$3", "style": style}
        }
        with pytest.raises(ValueError, match="'style' must be in the range 1-48"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_valid_style_is_applied(self, tmp_xlsx: str) -> None:
        """A style inside 1-48 reaches the chart XML."""
        df = pd.DataFrame({"A": [1, 2]})
        charts: dict[str, ChartOptions] = {
            "D2": {"type": "bar", "data_range": "Sheet1!$A$2:$A$3", "style": 42}
        }
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)
        with zipfile.ZipFile(tmp_xlsx) as zf:
            assert '<c:style val="42"/>' in zf.read("xl/charts/chart1.xml").decode("utf-8")

    def test_chart_series_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A typo in a series-item key is rejected, not silently dropped."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
        # Intentionally invalid: 'categorie_range' is a typo for 'categories_range'.
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "column",
                "series": [
                    {
                        "values_range": "Sheet1!$B$2:$B$4",
                        "categorie_range": "Sheet1!$A$2:$A$4",  # type: ignore[typeddict-unknown-key]
                    }
                ],
            }
        }
        with pytest.raises(ValueError, match="unknown option"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_bare_values_range_raises(self, tmp_xlsx: str) -> None:
        """A values_range without a sheet name gives a clear, actionable error."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        # Intentionally invalid: bare range, missing the 'Sheet1!' qualifier.
        charts: dict[str, ChartOptions] = {"D2": {"type": "bar", "values_range": "A2:A3"}}
        with pytest.raises(ValueError, match="must include a sheet name"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_bare_categories_range_raises(self, tmp_xlsx: str) -> None:
        """A categories_range without a sheet name gives a clear, actionable error."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        # Intentionally invalid: bare range, missing the 'Sheet1!' qualifier.
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "bar",
                "values_range": "Sheet1!$A$2:$A$3",
                "categories_range": "A1:A1",
            }
        }
        with pytest.raises(ValueError, match="must include a sheet name"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_valid_values_range_with_wrong_type_values_alias_raises(
        self, tmp_xlsx: str
    ) -> None:
        """A malformed later alias still raises a type error.

        Even though an earlier alias ('values_range') in the same config is
        already valid, a malformed 'values' alias must still be checked.
        """
        df = pd.DataFrame({"A": [1, 2, 3]})
        # Intentionally invalid: 'values' (a non-string) must still be checked
        # even though 'values_range' already resolved to a valid value.
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "line",
                "values_range": "Sheet1!$A$2:$A$3",
                "values": 123,  # type: ignore[typeddict-item]
            }
        }
        with pytest.raises(ValueError, match="must be a string"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_series_bare_chart_level_categories_fallback_raises(self, tmp_xlsx: str) -> None:
        """A bare chart-level categories fallback (reused by series without their own) is rejected."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
        # Intentionally invalid: chart-level 'categories_range' is bare and is
        # reused as the fallback by the series item below.
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "column",
                "categories_range": "A2:A4",
                "series": [{"values_range": "Sheet1!$B$2:$B$4"}],
            }
        }
        with pytest.raises(ValueError, match="must include a sheet name"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    @pytest.mark.parametrize(
        "key", ["data_range", "values_range", "values", "name", "series_name"]
    )
    def test_chart_level_series_key_beside_series_raises(self, key: str, tmp_xlsx: str) -> None:
        """A chart-level range or series name beside `series` is rejected, not dropped.

        The series branch reads its ranges and names from the series items only,
        so each of these keys used to be accepted, ignored, and never mentioned
        again -- a chart quietly missing the series the caller asked for. The
        rest of this library rejects contradictory input loudly; this is that.
        """
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6], "C": [7, 8, 9]})
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "column",
                key: "Sheet1!$C$2:$C$4",  # type: ignore[misc]
                "series": [{"values_range": "Sheet1!$B$2:$B$4"}],
            }
        }
        with pytest.raises(ValueError, match=f"'{key}' cannot be combined with 'series'"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)

    def test_chart_level_series_keys_are_all_named_at_once(self, tmp_xlsx: str) -> None:
        """Two offending keys are both reported, so fixing one does not reveal the next."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6], "C": [7, 8, 9]})
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "column",
                "data_range": "Sheet1!$C$2:$C$4",
                "name": "Ignored",
                "series": [{"values_range": "Sheet1!$B$2:$B$4"}],
            }
        }
        with pytest.raises(ValueError, match="cannot be combined with 'series'") as excinfo:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)
        message = str(excinfo.value)
        assert "'data_range'" in message
        assert "'name'" in message

    def test_chart_level_non_series_keys_still_work_with_series(self, tmp_xlsx: str) -> None:
        """Control: the rejection is scoped to series-level keys and nothing else.

        A guard written one key too wide would break the documented chart-level
        `categories_range` fallback and every chart-level option that is not a
        series concept -- and the parametrized test above cannot see that,
        because every input it feeds is meant to be rejected. This is the input
        that must still be accepted.
        """
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
        charts: dict[str, ChartOptions] = {
            "D2": {
                "type": "column",
                "series": [{"values_range": "Sheet1!$B$2:$B$4"}],
                "categories_range": "Sheet1!$A$2:$A$4",
                "title": "Kept",
                "x_axis_name": "X",
                "style": 12,
            }
        }
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, charts=charts)
        with zipfile.ZipFile(tmp_xlsx) as zf:
            chart_xml = zf.read("xl/charts/chart1.xml").decode("utf-8")
        # The chart-level categories are the documented fallback for a series
        # item that specifies none, so they must reach the XML.
        assert "Sheet1!$A$2:$A$4" in chart_xml
        assert "Kept" in chart_xml
        assert '<c:style val="12"/>' in chart_xml


class TestSparklines:
    """Tests for native Excel sparklines (mini in-cell charts)."""

    @staticmethod
    def _sheet_xml(path: str, sheet: str = "sheet1") -> str:
        """Return the decoded worksheet XML (sparklines live there, not a separate part)."""
        with zipfile.ZipFile(path) as zf:
            return zf.read(f"xl/worksheets/{sheet}.xml").decode("utf-8")

    def test_single_sparkline(self, tmp_xlsx: str) -> None:
        """A single-cell location places one sparkline with its type and color."""
        df = pd.DataFrame({"q1": [10], "q2": [30], "q3": [20]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            sparklines={
                "D2": {"range": "Sheet1!A2:C2", "type": "column", "color": "#FF0000"}
            },
        )
        xml = self._sheet_xml(tmp_xlsx)
        assert xml.count("<x14:sparklineGroup ") == 1
        assert 'type="column"' in xml
        # The data range and location anchor must reach the XML.
        assert "<xm:f>Sheet1!A2:C2</xm:f>" in xml
        assert "<xm:sqref>D2</xm:sqref>" in xml
        # The custom series color must land as an ARGB value.
        assert 'rgb="FFFF0000"' in xml

    def test_grouped_sparkline_expands_per_row(self, tmp_xlsx: str) -> None:
        """A range location makes a grouped sparkline: one per row with a sliced data range."""
        df = pd.DataFrame({"q1": [10, 30, 20], "q2": [15, 25, 35], "q3": [25, 20, 45]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            sparklines={
                "D2:D4": {
                    "range": "Sheet1!A2:C4",
                    "type": "line",
                    "markers": True,
                    "high_point": True,
                }
            },
        )
        xml = self._sheet_xml(tmp_xlsx)
        # One group containing three per-row sparklines.
        assert xml.count("<x14:sparklineGroup ") == 1
        assert xml.count("<x14:sparkline>") == 3
        assert 'markers="1"' in xml
        assert 'high="1"' in xml
        # Each row's data range is sliced and anchored to its own cell.
        assert "<xm:f>Sheet1!A2:C2</xm:f>" in xml
        assert "<xm:sqref>D2</xm:sqref>" in xml
        assert "<xm:f>Sheet1!A4:C4</xm:f>" in xml
        assert "<xm:sqref>D4</xm:sqref>" in xml

    def test_sparklines_with_dfs_to_xlsx_per_sheet(self, tmp_xlsx: str) -> None:
        """Sparklines work via per-sheet options in dfs_to_xlsx."""
        df1 = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
        df2 = pd.DataFrame({"a": [5, 6], "b": [7, 8]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "North", {"sparklines": {"D2": {"range": "North!A2:B2"}}}),
                (df2, "South", {"sparklines": {"D2": {"range": "South!A2:B2"}}}),
            ],
            tmp_xlsx,
        )
        assert "<x14:sparklineGroup" in self._sheet_xml(tmp_xlsx, "sheet1")
        assert "<x14:sparklineGroup" in self._sheet_xml(tmp_xlsx, "sheet2")

    def test_sparkline_missing_range_raises(self, tmp_xlsx: str) -> None:
        """Omitting the required 'range' key raises a clear error."""
        df = pd.DataFrame({"A": [1, 2]})
        # Intentionally invalid: no 'range' key.
        sparklines: dict[str, SparklineOptions] = {"D2": {"type": "line"}}  # type: ignore[typeddict-item]
        with pytest.raises(ValueError, match="missing required 'range'"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    def test_sparkline_invalid_type_raises(self, tmp_xlsx: str) -> None:
        """An unknown sparkline type is rejected."""
        df = pd.DataFrame({"A": [1, 2]})
        # Intentionally invalid: 'squiggle' is not a known sparkline type.
        sparklines: dict[str, SparklineOptions] = {
            "D2": {"range": "Sheet1!A2:A3", "type": "squiggle"}  # type: ignore[typeddict-item]
        }
        with pytest.raises(ValueError, match="Unknown sparkline type"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    def test_sparkline_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A typo in a sparkline option key is rejected, not silently dropped."""
        df = pd.DataFrame({"A": [1, 2]})
        # Intentionally invalid: 'marker' is a typo for 'markers'.
        sparklines: dict[str, SparklineOptions] = {
            "D2": {"range": "Sheet1!A2:A3", "marker": True}  # type: ignore[typeddict-unknown-key]
        }
        with pytest.raises(ValueError, match="unknown option"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    def test_sparkline_bare_range_raises(self, tmp_xlsx: str) -> None:
        """A 'range' without a sheet name gives a clear, actionable error."""
        df = pd.DataFrame({"A": [1, 2]})
        # Intentionally invalid: bare range, missing the 'Sheet1!' qualifier.
        sparklines: dict[str, SparklineOptions] = {"D2": {"range": "A2:A3"}}
        with pytest.raises(ValueError, match="must include a sheet name"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    def test_sparkline_rectangular_location_raises(self, tmp_xlsx: str) -> None:
        """A 2D block location for a grouped sparkline is rejected as ambiguous."""
        df = pd.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6], "c": [7, 8, 9]})
        # Intentionally invalid: 'E2:F4' spans two rows and two columns.
        sparklines: dict[str, SparklineOptions] = {"E2:F4": {"range": "Sheet1!A2:C4"}}
        with pytest.raises(ValueError, match="single row or column"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    @pytest.mark.parametrize("style", [0, 37, 300, -1])
    def test_sparkline_invalid_style_raises(self, style: int, tmp_xlsx: str) -> None:
        """A style outside 1-36 is rejected with the 1-36 message, including values >255 (300) and negatives (-1)."""
        df = pd.DataFrame({"A": [1, 2]})
        sparklines: dict[str, SparklineOptions] = {
            "D2": {"range": "Sheet1!A2:A3", "style": style}
        }
        with pytest.raises(ValueError, match="'style' must be in the range 1-36"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines=sparklines)

    def test_sparkline_win_loss_and_group_axes(self, tmp_xlsx: str) -> None:
        """win_loss type, negative points, and grouped min/max axes reach the XML."""
        df = pd.DataFrame({"q1": [10, -30, 20], "q2": [15, 25, -35], "q3": [25, 20, 45]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            sparklines={
                "E2:E4": {
                    "range": "Sheet1!A2:C4",
                    "type": "win_loss",
                    "negative_points": True,
                    "group_max": True,
                    "group_min": True,
                    "negative_points_color": "#00B050",
                }
            },
        )
        xml = self._sheet_xml(tmp_xlsx)
        assert 'type="stacked"' in xml  # win_loss serializes as a stacked sparkline
        assert 'negative="1"' in xml
        assert 'maxAxisType="group"' in xml
        assert 'minAxisType="group"' in xml
        assert 'rgb="FF00B050"' in xml  # negative-points color

    def test_sparkline_custom_axes_and_date_range(self, tmp_xlsx: str) -> None:
        """custom_max/min and a date_range serialize as manual axes and a date axis."""
        df = pd.DataFrame({"q1": [10], "q2": [30], "q3": [20]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            sparklines={
                "F2": {
                    "range": "Sheet1!A2:C2",
                    "custom_max": 100.0,
                    "custom_min": 0.0,
                    "date_range": "Sheet1!A1:C1",
                }
            },
        )
        xml = self._sheet_xml(tmp_xlsx)
        assert 'manualMax="100"' in xml
        assert 'manualMin="0"' in xml
        assert 'maxAxisType="custom"' in xml
        assert 'minAxisType="custom"' in xml
        assert 'dateAxis="1"' in xml
        assert "<xm:f>Sheet1!A1:C1</xm:f>" in xml  # the supplied date range

    @pytest.mark.parametrize(
        ("key", "value", "expected"),
        [
            ("markers", True, 'markers="1"'),
            ("high_point", True, 'high="1"'),
            ("low_point", True, 'low="1"'),
            ("first_point", True, 'first="1"'),
            ("last_point", True, 'last="1"'),
            ("negative_points", True, 'negative="1"'),
            ("show_axis", True, 'displayXAxis="1"'),
            ("show_hidden_data", True, 'displayHidden="1"'),
            ("group_max", True, 'maxAxisType="group"'),
            ("group_min", True, 'minAxisType="group"'),
            ("right_to_left", True, 'rightToLeft="1"'),
            # Read by column: the first sparkline of the group takes column A.
            ("column_order", True, "<xm:f>Sheet1!A2:A3</xm:f>"),
            ("color", "#112233", '<x14:colorSeries rgb="FF112233"/>'),
            ("high_point_color", "#112234", '<x14:colorHigh rgb="FF112234"/>'),
            ("low_point_color", "#112235", '<x14:colorLow rgb="FF112235"/>'),
            ("first_point_color", "#112236", '<x14:colorFirst rgb="FF112236"/>'),
            ("last_point_color", "#112237", '<x14:colorLast rgb="FF112237"/>'),
            ("negative_points_color", "#112238", '<x14:colorNegative rgb="FF112238"/>'),
            ("markers_color", "#112239", '<x14:colorMarkers rgb="FF112239"/>'),
            # Style 5 draws from theme colour 8; the default, style 1, from 4.
            ("style", 5, '<x14:colorSeries theme="8"'),
            ("line_weight", 2.5, 'lineWeight="2.5"'),
        ],
    )
    def test_every_sparkline_toggle_and_colour_reaches_the_xml(
        self, tmp_xlsx_factory: Callable[..., str], key: str, value: object, expected: str
    ) -> None:
        """One sparkline per key, and each key's own attribute or colour element.

        Every case is checked against a sparkline without the key as well, so
        an expected string that is there by default cannot pass. A grouped
        location (one sparkline per row of A2:B3) is used for all of them,
        because ``column_order`` changes nothing on a single sparkline.

        The keys are those ``SPARKLINE_KEYS`` in src/apply/sparklines.rs lists
        besides ``range``, ``type``, ``custom_max``, ``custom_min`` and
        ``date_range``, which the tests above cover.
        """
        df = pd.DataFrame({"a": [1, -2], "b": [3, 4]})
        base: Any = {"range": "Sheet1!A2:B3"}
        with_key: Any = {**base, key: value}
        plain, configured = tmp_xlsx_factory(), tmp_xlsx_factory()
        xlsxturbo.df_to_xlsx(df, plain, sparklines={"D2:D3": base})
        xlsxturbo.df_to_xlsx(df, configured, sparklines={"D2:D3": with_key})
        assert expected not in self._sheet_xml(plain), f"{expected!r} is there without {key!r}"
        assert expected in self._sheet_xml(configured), f"{key!r} did not reach the XML as {expected!r}"


class TestNumbersReachingTheXmlAreChecked:
    """NaN, infinity and negative sizes were written into the XML verbatim.

    ``line_weight=nan`` became ``lineWeight="NaN"``; nothing raised. The cases
    that stay valid (a weight of zero, a large scale) are the controls: a screen
    refusing everything would pass the rejections alone.
    """

    @pytest.mark.parametrize(
        ("key", "value"),
        [
            ("line_weight", float("nan")),
            ("line_weight", -1.0),
            ("custom_max", float("inf")),
            ("custom_min", float("nan")),
        ],
    )
    def test_sparkline_numbers(self, tmp_xlsx: str, key: str, value: float) -> None:
        """Each sparkline number is refused, naming the sparkline and the key."""
        df = pd.DataFrame({"q1": [10], "q2": [30]})
        config: Any = {"range": "Sheet1!A2:B2", key: value}
        with pytest.raises(xlsxturbo.ConfigurationError, match=rf"sparklines\['C2'\]: '{key}' must"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines={"C2": config})

    def test_sparkline_weight_zero_is_accepted(self, tmp_xlsx: str) -> None:
        """The control: zero is a weight, not a defect."""
        df = pd.DataFrame({"q1": [10], "q2": [30]})
        config: Any = {"range": "Sheet1!A2:B2", "line_weight": 0.0}
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, sparklines={"C2": config})
        with zipfile.ZipFile(tmp_xlsx) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert 'lineWeight="0"' in xml

    @pytest.mark.parametrize("size", [float("nan"), -3.0], ids=["nan", "negative"])
    def test_textbox_font_size(self, tmp_xlsx: str, size: float) -> None:
        """A textbox font size is refused the same way a cell font size is."""
        df = pd.DataFrame({"A": [1]})
        config: Any = {"text": "T", "font": {"size": size}}
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"textboxes\['B2'\].*'size' must"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": config})

    @pytest.mark.parametrize("key", ["scale_width", "scale_height"])
    def test_image_scale_must_be_finite(
        self, tmp_xlsx_factory: Callable[..., str], key: str
    ) -> None:
        """An image scale of NaN is refused; a large finite one is kept (the control)."""
        img_path = tmp_xlsx_factory(".png")
        Path(img_path).write_bytes(base64.b64decode(TINY_PNG_B64))
        df = pd.DataFrame({"A": [1]})
        bad: Any = {"path": img_path, key: float("nan")}
        with pytest.raises(xlsxturbo.ConfigurationError, match=rf"images\['B5'\]: '{key}' must"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx_factory(), images={"B5": bad})
        good: Any = {"path": img_path, key: 50.0}
        xlsxturbo.df_to_xlsx(df, tmp_xlsx_factory(), images={"B5": good})


class TestDeterministicFeatureMapOrder:
    """Regression tests for N3: feature maps must preserve Python dict order.

    Rust's `std::collections::HashMap` seeds a fresh random hasher per
    instance, so two otherwise-identical `df_to_xlsx` calls used to be able
    to emit images/comments/checkboxes/textboxes in a different order in the
    generated XML even within the same process. The cell_ref-keyed feature
    maps (images, checkboxes, textboxes, comments, charts, sparklines,
    rich_text) now use `IndexMap`, which preserves insertion order
    deterministically. These tests build the same options twice in one
    process (so a per-instance random hash seed would show up as a diff) and
    assert the emitted XML parts are byte-for-byte identical.
    """

    def test_images_and_comments_order_is_stable_across_runs(
        self, tmp_xlsx_factory: Callable[..., str]
    ) -> None:
        """Two writes with the same images/comments dicts emit identical drawing/comment XML."""
        df = pd.DataFrame({"A": [1, 2, 3, 4]})
        png_path = tmp_xlsx_factory(".png")
        with Path(png_path).open("wb") as f:
            f.write(base64.b64decode(TINY_PNG_B64))

        # Insertion order deliberately out of cell-reference order, so a
        # HashMap's random iteration order would be very likely to visibly
        # reorder at least one of these two 4-entry maps between the two runs.
        images: dict[str, str | ImageOptions] = {
            "D1": png_path,
            "B1": png_path,
            "F1": png_path,
            "H1": png_path,
        }
        comments: dict[str, str | CommentOptions] = {
            "A1": "one",
            "C1": "two",
            "E1": "three",
            "G1": "four",
        }

        path1 = tmp_xlsx_factory()
        path2 = tmp_xlsx_factory()
        xlsxturbo.df_to_xlsx(df, path1, images=images, comments=comments)
        xlsxturbo.df_to_xlsx(df, path2, images=images, comments=comments)

        with zipfile.ZipFile(path1) as zf1, zipfile.ZipFile(path2) as zf2:
            drawing1 = zf1.read("xl/drawings/drawing1.xml")
            drawing2 = zf2.read("xl/drawings/drawing1.xml")
            comments_xml1 = zf1.read("xl/comments1.xml")
            comments_xml2 = zf2.read("xl/comments1.xml")

        assert drawing1 == drawing2, "image placement order must be reproducible across runs"
        assert comments_xml1 == comments_xml2, "comment order must be reproducible across runs"

    def test_checkboxes_and_textboxes_order_is_stable_across_runs(
        self, tmp_xlsx_factory: Callable[..., str]
    ) -> None:
        """Two writes with the same checkboxes/textboxes dicts emit identical drawing XML."""
        df = pd.DataFrame({"A": [1, 2, 3, 4]})
        checkboxes: dict[str, bool | CheckboxOptions] = {
            "D1": True,
            "B1": False,
            "F1": True,
            "H1": False,
        }
        textboxes: dict[str, str | TextboxOptions] = {
            "C1": "one",
            "E1": "two",
            "G1": "three",
            "I1": "four",
        }

        path1 = tmp_xlsx_factory()
        path2 = tmp_xlsx_factory()
        xlsxturbo.df_to_xlsx(df, path1, checkboxes=checkboxes, textboxes=textboxes)
        xlsxturbo.df_to_xlsx(df, path2, checkboxes=checkboxes, textboxes=textboxes)

        with zipfile.ZipFile(path1) as zf1, zipfile.ZipFile(path2) as zf2:
            drawing1 = zf1.read("xl/drawings/drawing1.xml")
            drawing2 = zf2.read("xl/drawings/drawing1.xml")

        assert drawing1 == drawing2, "checkbox/textbox order must be reproducible across runs"
