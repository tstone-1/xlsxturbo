"""Tests for media features: images, checkboxes, textboxes, and charts."""

from __future__ import annotations

import base64
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, TINY_PNG_B64, active_ws, load_workbook

if TYPE_CHECKING:
    from xlsxturbo.xlsxturbo import (
        ChartOptions,
        CheckboxOptions,
        CommentOptions,
        ImageOptions,
        SparklineOptions,
        TextboxOptions,
    )

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


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
        """Image with scaling options."""
        df = pd.DataFrame({"A": [1]})
        path = tmp_xlsx_factory()
        png_data = base64.b64decode(TINY_PNG_B64)
        img_path = tmp_xlsx_factory(".png")
        with Path(img_path).open("wb") as f:
            f.write(png_data)

        xlsxturbo.df_to_xlsx(
            df,
            path,
            images={"B5": {"path": img_path, "scale_width": 2.0, "scale_height": 2.0}},
        )
        with zipfile.ZipFile(path) as zf:
            media = [n for n in zf.namelist() if n.startswith("xl/media/")]
            assert media, "no embedded image found in xl/media/"
            # A drawing relationship must anchor the image to the sheet.
            assert any(n.startswith("xl/drawings/") for n in zf.namelist())


class TestCheckboxes:
    """Tests for checkboxes feature (v0.13.0)."""

    def test_checkbox_simple_bool(self, tmp_xlsx: str) -> None:
        """Checkboxes with plain bool values."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"D1": True, "D2": False})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Checkboxes render as boolean TRUE/FALSE in cells
        assert ws["D1"].value is True
        assert ws["D2"].value is False

    def test_checkbox_dict_form(self, tmp_xlsx: str) -> None:
        """Checkbox with dict specifying checked state."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": {"checked": True}})
        assert Path(tmp_xlsx).exists()
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
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["C3"].value is True

    def test_checkbox_missing_checked_key_raises(self, tmp_xlsx: str) -> None:
        """Dict form without 'checked' raises a clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="checked"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, checkboxes={"B2": {"format": {"bold": True}}})

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
        with pytest.raises(ValueError):  # noqa: PT011
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
        assert Path(tmp_xlsx).exists()
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
        assert Path(tmp_xlsx).exists()


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
        assert Path(tmp_xlsx).exists()

    def test_textbox_with_all_options(self, tmp_xlsx: str) -> None:
        """Textbox with every supported option."""
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
                    "line_color": "#000000",
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

    def test_textbox_font_partial(self, tmp_xlsx: str) -> None:
        """Partial font options work (only some keys set)."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "font": {"bold": True}}})
        assert Path(tmp_xlsx).exists()

    def test_textbox_missing_text_raises(self, tmp_xlsx: str) -> None:
        """Dict form without 'text' raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="text"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"width": 100}})

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
        """Offsets without size/font still writes file (exercises insert_shape_with_offset path)."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "x_offset": 15, "y_offset": 25}})
        assert Path(tmp_xlsx).exists()

    def test_textbox_wrong_value_type_raises(self, tmp_xlsx: str) -> None:
        """Non-string, non-dict value raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="textboxes"):
            # Intentionally invalid: value must be a string or dict.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": 123})  # type: ignore[dict-item]

    def test_textbox_invalid_color_raises(self, tmp_xlsx: str) -> None:
        """Invalid color string surfaces the parse error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="color"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, textboxes={"B2": {"text": "T", "fill_color": "not_a_color"}})

    def test_textbox_invalid_cell_ref_raises(self, tmp_xlsx: str) -> None:
        """Invalid cell reference raises."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError):  # noqa: PT011
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
        assert Path(tmp_xlsx).exists()

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
        assert Path(tmp_xlsx).exists()


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
        sparklines: dict[str, SparklineOptions] = {"D2": {"type": "line"}}
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
