"""Tests for conditional formatting features (color scales, data bars, icon sets, cell rules)."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

if TYPE_CHECKING:
    from xlsxturbo.xlsxturbo import ConditionalFormat

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestConditionalFormatting:
    """Tests for conditional formatting feature (v0.8.0)."""

    def test_2_color_scale(self, tmp_xlsx: str) -> None:
        """Verify 2-color scale conditional format."""
        df = pd.DataFrame({"Score": [10, 50, 90]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={
                "Score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
            # Assert it's actually a color scale with both configured colors.
            assert "<COLORSCALE>" in xml
            assert "FF0000" in xml
            assert "00FF00" in xml

    def test_3_color_scale(self, tmp_xlsx: str) -> None:
        """Verify 3-color scale conditional format."""
        df = pd.DataFrame({"Value": [1, 5, 10]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={
                "Value": {
                    "type": "3_color_scale",
                    "min_color": "#F8696B",
                    "mid_color": "#FFEB84",
                    "max_color": "#63BE7B",
                }
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
            assert "<COLORSCALE>" in xml
            # All three configured colors must be present.
            assert "F8696B" in xml
            assert "FFEB84" in xml
            assert "63BE7B" in xml

    def test_data_bar(self, tmp_xlsx: str) -> None:
        """Verify data bar conditional format."""
        df = pd.DataFrame({"Progress": [25, 50, 75, 100]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={
                "Progress": {"type": "data_bar", "bar_color": "#638EC6"}
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
            # Assert it's a data bar with the configured bar color.
            assert "DATABAR" in xml
            assert "638EC6" in xml

    def test_icon_set(self, tmp_xlsx: str) -> None:
        """Verify icon set conditional format."""
        df = pd.DataFrame({"Status": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={
                "Status": {"type": "icon_set", "icon_type": "3_traffic_lights"}
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            # Assert it's an iconSet rule with three thresholds (a 3-icon set).
            # Excel omits the iconSet type attr for the default 3-traffic-lights,
            # so verify the structure rather than the (absent) type string.
            assert 'type="iconSet"' in xml
            assert xml.count("<cfvo ") == 3

    def test_2_color_scale_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A typo'd key ('min_colour') is rejected with the valid keys for the resolved type."""
        df = pd.DataFrame({"Score": [10, 50, 90]})
        with pytest.raises(ValueError, match="unknown option 'min_colour'"):
            # Intentionally invalid: 'min_colour' is a typo for 'min_color'.
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                conditional_formats={
                    "Score": {"type": "2_color_scale", "min_colour": "#FF0000", "max_color": "#00FF00"}  # type: ignore[typeddict-unknown-key]
                },
            )

    def test_data_bar_wrong_family_key_raises(self, tmp_xlsx: str) -> None:
        """A key valid for another type ('min_color' belongs to color scales) is rejected on data_bar."""
        df = pd.DataFrame({"Progress": [25, 50, 75]})
        with pytest.raises(ValueError, match="unknown option 'min_color'"):
            # Intentionally invalid: 'min_color' is a color-scale key, not a data_bar key.
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                conditional_formats={
                    "Progress": {"type": "data_bar", "min_color": "#638EC6"}  # type: ignore[typeddict-unknown-key]
                },
            )

    def test_conditional_format_with_pattern(self, tmp_xlsx: str) -> None:
        """Verify conditional format with wildcard column pattern."""
        df = pd.DataFrame({"score_a": [80], "score_b": [60], "name": ["Alice"]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={
                "score_*": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
            },
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Should have conditional formats on both score columns
        assert len(ws.conditional_formatting) >= 1
        wb.close()


class TestCellConditionalFormat:
    """Tests for rule-based conditional formatting (v0.12.0)."""

    @pytest.mark.parametrize(
        ("criteria", "extra", "expected_type", "expected_operator", "expected_formula"),
        [
            pytest.param("equal_to", {"value": "ERROR"}, "cellIs", "equal", ['"ERROR"'], id="equal_to_string"),
            pytest.param("equal_to", {"value": 70}, "cellIs", "equal", ["70"], id="equal_to_number"),
            pytest.param("not_equal_to", {"value": "ok"}, "cellIs", "notEqual", ['"ok"'], id="not_equal_to"),
            pytest.param("greater_than", {"value": 70}, "cellIs", "greaterThan", ["70"], id="greater_than"),
            pytest.param(
                "greater_than_or_equal_to",
                {"value": 75},
                "cellIs",
                "greaterThanOrEqual",
                ["75"],
                id="greater_than_or_equal_to",
            ),
            pytest.param("less_than", {"value": 15}, "cellIs", "lessThan", ["15"], id="less_than"),
            pytest.param(
                "less_than_or_equal_to",
                {"value": 5},
                "cellIs",
                "lessThanOrEqual",
                ["5"],
                id="less_than_or_equal_to",
            ),
            pytest.param(
                "between",
                {"min_value": 20, "max_value": 40},
                "cellIs",
                "between",
                ["20", "40"],
                id="between",
            ),
            pytest.param(
                "not_between",
                {"min_value": 3, "max_value": 12},
                "cellIs",
                "notBetween",
                ["3", "12"],
                id="not_between",
            ),
            pytest.param(
                "containing",
                {"value": "error"},
                "containsText",
                "containsText",
                ['NOT(ISERROR(SEARCH("error",A2)))'],
                id="containing",
            ),
            pytest.param(
                "not_containing",
                {"value": "hello"},
                "notContainsText",
                "notContains",
                ['ISERROR(SEARCH("hello",A2))'],
                id="not_containing",
            ),
            pytest.param(
                "begins_with",
                {"value": "ERR"},
                "beginsWith",
                "beginsWith",
                ['LEFT(A2,3)="ERR"'],
                id="begins_with",
            ),
            pytest.param(
                "ends_with",
                {"value": ".pdf"},
                "endsWith",
                "endsWith",
                ['RIGHT(A2,4)=".pdf"'],
                id="ends_with",
            ),
            pytest.param("blanks", {}, "containsBlanks", None, ["LEN(TRIM(A2))=0"], id="blanks"),
            pytest.param("no_blanks", {}, "notContainsBlanks", None, ["LEN(TRIM(A2))>0"], id="no_blanks"),
        ],
    )
    def test_cell_criteria_content(
        self,
        tmp_xlsx: str,
        criteria: str,
        extra: dict[str, str | int | float],
        expected_type: str,
        expected_operator: str | None,
        expected_formula: list[str],
    ) -> None:
        """Verify every 'cell' criteria produces the documented openpyxl rule type/operator/formula.

        Replaces a set of prior existence-only smoke tests (one per criteria)
        with a single content-verified parametrization, so a regression in the
        underlying rule mapping (src/apply/conditional_formats.rs) is caught
        instead of only "the file was written".
        """
        df = pd.DataFrame({"A": ["ok", "fail", None, 5, 10]})
        config: ConditionalFormat = {"type": "cell", "criteria": criteria, **extra}  # type: ignore[typeddict-item]
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            conditional_formats={"A": config},
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cf_rules = ws.conditional_formatting
        rule = next(iter(cf_rules)).rules[0]
        assert rule.type == expected_type
        assert rule.operator == expected_operator
        assert rule.formula == expected_formula
        wb.close()

    def test_cell_multiple_rules_list(self, tmp_xlsx: str) -> None:
        """Verify multiple rules on one column via list."""
        df = pd.DataFrame({"severity": ["HIGH", "MEDIUM", "LOW", "HIGH"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "severity": [
                {"type": "cell", "criteria": "equal_to", "value": "HIGH",
                 "format": {"bg_color": "#FF0000"}},
                {"type": "cell", "criteria": "equal_to", "value": "MEDIUM",
                 "format": {"bg_color": "#FFA500"}},
                {"type": "cell", "criteria": "equal_to", "value": "LOW",
                 "format": {"bg_color": "#FFFF00"}},
            ]
        })
        assert Path(tmp_xlsx).exists()

    def test_backward_compat_single_dict(self, tmp_xlsx: str) -> None:
        """Verify existing single-dict format still works (backward compat)."""
        df = pd.DataFrame({"score": [10, 50, 90]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}
        })
        assert Path(tmp_xlsx).exists()

    def test_cell_with_pattern_matching(self, tmp_xlsx: str) -> None:
        """Verify cell conditional format works with wildcard patterns."""
        df = pd.DataFrame({"score_a": [50], "score_b": [80], "name": ["x"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "score_*": {
                "type": "cell",
                "criteria": "greater_than",
                "value": 60,
                "format": {"bold": True}
            }
        })
        assert Path(tmp_xlsx).exists()

    def test_cell_format_with_border(self, tmp_xlsx: str) -> None:
        """Verify cell conditional format can include border styling."""
        df = pd.DataFrame({"val": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "val": {
                "type": "cell",
                "criteria": "equal_to",
                "value": 2,
                "format": {"bg_color": "#FFFF00", "border": "thin"}
            }
        })
        assert Path(tmp_xlsx).exists()

    def test_invalid_criteria_raises(self, tmp_xlsx: str) -> None:
        """Verify invalid criteria raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="Unknown criteria"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
                "A": {"type": "cell", "criteria": "invalid", "value": 1,
                       "format": {"bold": True}}
            })

    def test_missing_criteria_raises(self, tmp_xlsx: str) -> None:
        """Verify missing criteria key raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="requires 'criteria'"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
                "A": {"type": "cell", "value": 1, "format": {"bold": True}}
            })

    def test_cell_with_polars(self, tmp_xlsx: str) -> None:
        """Verify cell conditional format works with polars."""
        df = pl.DataFrame({"score": [50, 75, 100]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "score": {
                "type": "cell",
                "criteria": "greater_than_or_equal_to",
                "value": 75,
                "format": {"bg_color": "#00FF00"}
            }
        })
        assert Path(tmp_xlsx).exists()

    def test_cell_without_format_key(self, tmp_xlsx: str) -> None:
        """Verify cell rule without format key still works (no styling applied)."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "A": {"type": "cell", "criteria": "greater_than", "value": 1}
        })
        assert Path(tmp_xlsx).exists()

    def test_cell_missing_value_raises(self, tmp_xlsx: str) -> None:
        """Verify missing value key for value-requiring criteria raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="missing 'value'"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
                "A": {"type": "cell", "criteria": "greater_than",
                       "format": {"bold": True}}
            })

    def test_cell_numeric_comparison_correct(self, tmp_xlsx: str) -> None:
        """Verify numeric values produce numeric (not string) comparisons in Excel."""
        df = pd.DataFrame({"score": [8, 50, 70, 100]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "score": {
                "type": "cell",
                "criteria": "greater_than",
                "value": 70,
                "format": {"bg_color": "#00FF00"}
            }
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cf_rules = ws.conditional_formatting
        assert len(list(cf_rules)) > 0
        rule = next(iter(cf_rules))
        cf = rule.rules[0]
        assert cf.type == "cellIs"
        assert cf.operator == "greaterThan"
        assert cf.formula == ['70']
        wb.close()

    def test_cell_string_comparison_correct(self, tmp_xlsx: str) -> None:
        """Verify string values produce string comparisons in Excel."""
        df = pd.DataFrame({"status": ["OK", "ERROR", "OK"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats={
            "status": {
                "type": "cell",
                "criteria": "equal_to",
                "value": "ERROR",
                "format": {"bg_color": "#FF0000"}
            }
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cf_rules = ws.conditional_formatting
        assert len(list(cf_rules)) > 0
        rule = next(iter(cf_rules))
        cf = rule.rules[0]
        assert cf.type == "cellIs"
        assert cf.operator == "equal"
        assert cf.formula == ['"ERROR"']
        wb.close()

    def test_invalid_list_item_raises(self, tmp_xlsx: str) -> None:
        """Verify non-dict item in conditional format list raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        invalid_formats = {
            "A": [{"type": "cell", "criteria": "equal_to", "value": 1, "format": {"bold": True}}, "not_a_dict"],
        }
        with pytest.raises(TypeError, match=r"list item .* must be a dict"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, conditional_formats=invalid_formats)  # type: ignore[arg-type]
