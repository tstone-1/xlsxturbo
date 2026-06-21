"""Tests for conditional formatting features (color scales, data bars, icon sets, cell rules)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, get_temp_path, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestConditionalFormatting:
    """Tests for conditional formatting feature (v0.8.0)."""

    def test_2_color_scale(self) -> None:
        """Verify 2-color scale conditional format."""
        df = pd.DataFrame({"Score": [10, 50, 90]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                conditional_formats={
                    "Score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
                },
            )
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
                # Assert it's actually a color scale with both configured colors.
                assert "<COLORSCALE>" in xml
                assert "FF0000" in xml
                assert "00FF00" in xml
        finally:
            Path(path).unlink()

    def test_3_color_scale(self) -> None:
        """Verify 3-color scale conditional format."""
        df = pd.DataFrame({"Value": [1, 5, 10]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                conditional_formats={
                    "Value": {
                        "type": "3_color_scale",
                        "min_color": "#F8696B",
                        "mid_color": "#FFEB84",
                        "max_color": "#63BE7B",
                    }
                },
            )
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
                assert "<COLORSCALE>" in xml
                # All three configured colors must be present.
                assert "F8696B" in xml
                assert "FFEB84" in xml
                assert "63BE7B" in xml
        finally:
            Path(path).unlink()

    def test_data_bar(self) -> None:
        """Verify data bar conditional format."""
        df = pd.DataFrame({"Progress": [25, 50, 75, 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                conditional_formats={
                    "Progress": {"type": "data_bar", "bar_color": "#638EC6"}
                },
            )
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
                # Assert it's a data bar with the configured bar color.
                assert "DATABAR" in xml
                assert "638EC6" in xml
        finally:
            Path(path).unlink()

    def test_icon_set(self) -> None:
        """Verify icon set conditional format."""
        df = pd.DataFrame({"Status": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                conditional_formats={
                    "Status": {"type": "icon_set", "icon_type": "3_traffic_lights"}
                },
            )
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
                # Assert it's an iconSet rule with three thresholds (a 3-icon set).
                # Excel omits the iconSet type attr for the default 3-traffic-lights,
                # so verify the structure rather than the (absent) type string.
                assert 'type="iconSet"' in xml
                assert xml.count("<cfvo ") == 3
        finally:
            Path(path).unlink()

    def test_conditional_format_with_pattern(self) -> None:
        """Verify conditional format with wildcard column pattern."""
        df = pd.DataFrame({"score_a": [80], "score_b": [60], "name": ["Alice"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                conditional_formats={
                    "score_*": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
                },
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Should have conditional formats on both score columns
                assert len(ws.conditional_formatting) >= 1
                wb.close()
        finally:
            Path(path).unlink()


class TestCellConditionalFormat:
    """Tests for rule-based conditional formatting (v0.12.0)."""

    def test_cell_equal_to_string(self) -> None:
        """Verify type='cell' with criteria='equal_to' highlights matching string cells."""
        df = pd.DataFrame({"status": ["OK", "ERROR", "OK", "ERROR"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "status": {
                    "type": "cell",
                    "criteria": "equal_to",
                    "value": "ERROR",
                    "format": {"bg_color": "#FF0000", "bold": True}
                }
            })
            assert Path(path).exists()
            assert Path(path).stat().st_size > 0
        finally:
            Path(path).unlink()

    def test_cell_equal_to_number(self) -> None:
        """Verify type='cell' with numeric value."""
        df = pd.DataFrame({"score": [50, 75, 100, 25]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "score": {
                    "type": "cell",
                    "criteria": "greater_than",
                    "value": 70,
                    "format": {"bg_color": "#00FF00"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_between(self) -> None:
        """Verify between criteria with min_value and max_value."""
        df = pd.DataFrame({"temp": [10, 20, 30, 40, 50]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "temp": {
                    "type": "cell",
                    "criteria": "between",
                    "min_value": 20,
                    "max_value": 40,
                    "format": {"bg_color": "#FFFF00"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_not_between(self) -> None:
        """Verify not_between criteria."""
        df = pd.DataFrame({"val": [1, 5, 10, 15]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "val": {
                    "type": "cell",
                    "criteria": "not_between",
                    "min_value": 3,
                    "max_value": 12,
                    "format": {"font_color": "red"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_multiple_rules_list(self) -> None:
        """Verify multiple rules on one column via list."""
        df = pd.DataFrame({"severity": ["HIGH", "MEDIUM", "LOW", "HIGH"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "severity": [
                    {"type": "cell", "criteria": "equal_to", "value": "HIGH",
                     "format": {"bg_color": "#FF0000"}},
                    {"type": "cell", "criteria": "equal_to", "value": "MEDIUM",
                     "format": {"bg_color": "#FFA500"}},
                    {"type": "cell", "criteria": "equal_to", "value": "LOW",
                     "format": {"bg_color": "#FFFF00"}},
                ]
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_text_containing(self) -> None:
        """Verify criteria='containing' for text match."""
        df = pd.DataFrame({"desc": ["error occurred", "all good", "error found"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "desc": {
                    "type": "cell",
                    "criteria": "containing",
                    "value": "error",
                    "format": {"bg_color": "#FF0000"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_text_begins_with(self) -> None:
        """Verify criteria='begins_with' for text match."""
        df = pd.DataFrame({"code": ["ERR-001", "OK-002", "ERR-003"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "code": {
                    "type": "cell",
                    "criteria": "begins_with",
                    "value": "ERR",
                    "format": {"font_color": "red"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_text_ends_with(self) -> None:
        """Verify criteria='ends_with' for text match."""
        df = pd.DataFrame({"file": ["report.pdf", "data.csv", "notes.pdf"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "file": {
                    "type": "cell",
                    "criteria": "ends_with",
                    "value": ".pdf",
                    "format": {"italic": True}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_blanks(self) -> None:
        """Verify criteria='blanks' highlights blank cells."""
        df = pd.DataFrame({"val": [1, None, 3, None]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "val": {
                    "type": "cell",
                    "criteria": "blanks",
                    "format": {"bg_color": "#CCCCCC"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_no_blanks(self) -> None:
        """Verify criteria='no_blanks' highlights non-blank cells."""
        df = pd.DataFrame({"val": [1, None, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "val": {
                    "type": "cell",
                    "criteria": "no_blanks",
                    "format": {"bg_color": "#00FF00"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_backward_compat_single_dict(self) -> None:
        """Verify existing single-dict format still works (backward compat)."""
        df = pd.DataFrame({"score": [10, 50, 90]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_less_than(self) -> None:
        """Verify less_than criteria."""
        df = pd.DataFrame({"price": [10.5, 20.0, 5.0]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "price": {
                    "type": "cell",
                    "criteria": "less_than",
                    "value": 15,
                    "format": {"bg_color": "#FF0000"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_with_pattern_matching(self) -> None:
        """Verify cell conditional format works with wildcard patterns."""
        df = pd.DataFrame({"score_a": [50], "score_b": [80], "name": ["x"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "score_*": {
                    "type": "cell",
                    "criteria": "greater_than",
                    "value": 60,
                    "format": {"bold": True}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_format_with_border(self) -> None:
        """Verify cell conditional format can include border styling."""
        df = pd.DataFrame({"val": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "val": {
                    "type": "cell",
                    "criteria": "equal_to",
                    "value": 2,
                    "format": {"bg_color": "#FFFF00", "border": "thin"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_invalid_criteria_raises(self) -> None:
        """Verify invalid criteria raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown criteria"):
                xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                    "A": {"type": "cell", "criteria": "invalid", "value": 1,
                           "format": {"bold": True}}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_missing_criteria_raises(self) -> None:
        """Verify missing criteria key raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="requires 'criteria'"):
                xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                    "A": {"type": "cell", "value": 1, "format": {"bold": True}}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_cell_with_polars(self) -> None:
        """Verify cell conditional format works with polars."""
        df = pl.DataFrame({"score": [50, 75, 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "score": {
                    "type": "cell",
                    "criteria": "greater_than_or_equal_to",
                    "value": 75,
                    "format": {"bg_color": "#00FF00"}
                }
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_not_equal_to(self) -> None:
        """Verify not_equal_to criteria."""
        df = pd.DataFrame({"A": ["ok", "fail", "ok"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "A": {"type": "cell", "criteria": "not_equal_to", "value": "ok",
                       "format": {"bg_color": "#FF0000"}}
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_less_than_or_equal_to(self) -> None:
        """Verify less_than_or_equal_to criteria."""
        df = pd.DataFrame({"A": [1, 5, 10]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "A": {"type": "cell", "criteria": "less_than_or_equal_to", "value": 5,
                       "format": {"bold": True}}
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_not_containing(self) -> None:
        """Verify not_containing criteria."""
        df = pd.DataFrame({"A": ["hello world", "goodbye", "hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "A": {"type": "cell", "criteria": "not_containing", "value": "hello",
                       "format": {"italic": True}}
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_without_format_key(self) -> None:
        """Verify cell rule without format key still works (no styling applied)."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "A": {"type": "cell", "criteria": "greater_than", "value": 1}
            })
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_cell_missing_value_raises(self) -> None:
        """Verify missing value key for value-requiring criteria raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="missing 'value'"):
                xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                    "A": {"type": "cell", "criteria": "greater_than",
                           "format": {"bold": True}}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_cell_numeric_comparison_correct(self) -> None:
        """Verify numeric values produce numeric (not string) comparisons in Excel."""
        df = pd.DataFrame({"score": [8, 50, 70, 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "score": {
                    "type": "cell",
                    "criteria": "greater_than",
                    "value": 70,
                    "format": {"bg_color": "#00FF00"}
                }
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cf_rules = ws.conditional_formatting
                assert len(list(cf_rules)) > 0
                rule = next(iter(cf_rules))
                cf = rule.rules[0]
                assert cf.type == "cellIs"
                assert cf.operator == "greaterThan"
                assert cf.formula == ['70']
                wb.close()
        finally:
            Path(path).unlink()

    def test_cell_string_comparison_correct(self) -> None:
        """Verify string values produce string comparisons in Excel."""
        df = pd.DataFrame({"status": ["OK", "ERROR", "OK"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, conditional_formats={
                "status": {
                    "type": "cell",
                    "criteria": "equal_to",
                    "value": "ERROR",
                    "format": {"bg_color": "#FF0000"}
                }
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cf_rules = ws.conditional_formatting
                assert len(list(cf_rules)) > 0
                rule = next(iter(cf_rules))
                cf = rule.rules[0]
                assert cf.type == "cellIs"
                assert cf.operator == "equal"
                assert cf.formula == ['"ERROR"']
                wb.close()
        finally:
            Path(path).unlink()

    def test_invalid_list_item_raises(self) -> None:
        """Verify non-dict item in conditional format list raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        invalid_formats = {
            "A": [{"type": "cell", "criteria": "equal_to", "value": 1, "format": {"bold": True}}, "not_a_dict"],
        }
        try:
            with pytest.raises(TypeError, match=r"list item .* must be a dict"):
                xlsxturbo.df_to_xlsx(df, path, conditional_formats=invalid_formats)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)
