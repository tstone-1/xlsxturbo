"""Tests for the data validation feature (dropdown lists, numeric ranges, messages)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestValidations:
    """Tests for data validation feature (v0.10.0)."""

    @pytest.mark.parametrize(("field", "attribute"), [("input_title", "promptTitle"), ("error_title", "errorTitle")])
    @pytest.mark.parametrize("with_message", [False, True])
    def test_validation_title_without_message_is_preserved(
        self, tmp_xlsx: str, field: str, attribute: str, with_message: bool,
    ) -> None:
        """Each title is independent of its optional message."""
        config: Any = {"type": "whole_number", "min": 0, "max": 1, field: "Synthetic title"}
        if with_message:
            config[field.replace("title", "message")] = "Synthetic message"
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})
        wb = load_workbook(tmp_xlsx)
        validation = active_ws(wb).data_validations.dataValidation[0]
        assert getattr(validation, attribute) == "Synthetic title"
        assert validation.formula1 == "0"
        assert validation.formula2 == "1"
        wb.close()

    @pytest.mark.parametrize("field", ["input_title", "error_title"])
    @pytest.mark.parametrize("with_message", [False, True])
    def test_validation_titles_always_validate_types(self, tmp_xlsx: str, field: str, with_message: bool) -> None:
        """Omitting a message cannot bypass validation of the title."""
        config: Any = {"type": "whole_number", field: 123}
        if with_message:
            config[field.replace("title", "message")] = "Synthetic message"
        with pytest.raises(xlsxturbo.ConfigurationError, match=rf"validations\['a'\].*{field}.*string"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})

    @pytest.mark.parametrize("field", ["min", "max"])
    @pytest.mark.parametrize("value", [float("nan"), float("inf"), float("-inf")])
    def test_decimal_validation_requires_finite_bounds(self, tmp_xlsx: str, field: str, value: float) -> None:
        """Non-finite bounds must not become validation formulas."""
        config: Any = {"type": "decimal", field: value}
        with pytest.raises(xlsxturbo.ConfigurationError, match=rf"validations\['a'\].*{field}.*finite"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})

    @pytest.mark.parametrize("field", ["min", "max"])
    @pytest.mark.parametrize("value", [False, True])
    def test_whole_number_validation_rejects_bool_bounds(self, tmp_xlsx: str, field: str, value: bool) -> None:
        """Booleans do not become zero/one numeric bounds."""
        config: Any = {"type": "whole_number", field: value}
        with pytest.raises(xlsxturbo.ConfigurationError, match=rf"validations\['a'\].*{field}.*bool"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})

    @pytest.mark.parametrize("kind", ["decimal", "whole_number"])
    def test_finite_zero_one_bounds_are_preserved(self, tmp_xlsx: str, kind: str) -> None:
        """The invalid-number screen keeps ordinary finite bounds working."""
        config: Any = {"type": kind, "min": 0, "max": 1}
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})
        wb = load_workbook(tmp_xlsx)
        validation = active_ws(wb).data_validations.dataValidation[0]
        assert validation.formula1 == "0"
        assert validation.formula2 == "1"
        wb.close()

    @pytest.mark.parametrize(
        ("kind", "bound"),
        [("whole_number", 7), ("decimal", 0.5), ("text_length", 3)],
    )
    @pytest.mark.parametrize(
        ("field", "operator"),
        [("min", "greaterThanOrEqual"), ("max", "lessThanOrEqual")],
    )
    def test_one_bound_writes_a_one_sided_rule(
        self, tmp_xlsx: str, kind: str, bound: float, field: str, operator: str,
    ) -> None:
        """Only ``min`` or only ``max`` writes one formula, with no stand-in for the other.

        The stand-in for a missing decimal bound was ``f64::MAX``, written as a
        309-digit number above Excel's largest, and Excel refused to open the file.
        """
        config: Any = {"type": kind, field: bound}
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})
        wb = load_workbook(tmp_xlsx)
        dv = active_ws(wb).data_validations.dataValidation[0]
        assert dv.operator == operator
        assert dv.formula1 == str(bound)
        assert dv.formula2 is None
        wb.close()

    def test_decimal_without_bounds_is_refused(self, tmp_xlsx: str) -> None:
        """A decimal rule with no bound has no value Excel can open, so it is refused."""
        with pytest.raises(
            xlsxturbo.ConfigurationError,
            match=r"validations\['a'\]: decimal validation requires 'min', 'max' or both",
        ):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": {"type": "decimal"}})

    @pytest.mark.parametrize(
        ("kind", "low", "high"),
        [("whole_number", "-2147483648", "2147483647"), ("text_length", "0", "4294967295")],
    )
    def test_no_bounds_keeps_the_full_range(self, tmp_xlsx: str, kind: str, low: str, high: str) -> None:
        """Without bounds these two keep their full-range rule, which Excel opens."""
        config: Any = {"type": kind}
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), tmp_xlsx, validations={"a": config})
        wb = load_workbook(tmp_xlsx)
        dv = active_ws(wb).data_validations.dataValidation[0]
        assert (dv.operator, dv.formula1, dv.formula2) == (None, low, high)
        wb.close()

    @pytest.mark.parametrize(("values", "index"), [(["Smith, John", "c"], 0), (["a", "b", "1,5"], 2)])
    def test_list_value_with_a_comma_is_refused(self, tmp_xlsx: str, values: list[str], index: int) -> None:
        """Excel splits an inline list at commas, so the value would become several entries."""
        with pytest.raises(
            xlsxturbo.ConfigurationError,
            match=rf"validations\['a'\]: 'values' item {index} .* contains a comma",
        ):
            xlsxturbo.df_to_xlsx(
                pd.DataFrame({"a": ["x"]}), tmp_xlsx, validations={"a": {"type": "list", "values": values}},
            )

    def test_list_value_with_a_space_is_written(self, tmp_xlsx: str) -> None:
        """Control for the comma check: other punctuation and spaces pass through."""
        xlsxturbo.df_to_xlsx(
            pd.DataFrame({"a": ["x"]}), tmp_xlsx, validations={"a": {"type": "list", "values": ["a b", "c;d"]}},
        )
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb).data_validations.dataValidation[0].formula1 == '"a b,c;d"'
        wb.close()

    def test_empty_list_is_refused(self, tmp_xlsx: str) -> None:
        """A dropdown with no entries is a rule nothing can satisfy."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"validations\['a'\]: 'values' must not be empty"):
            xlsxturbo.df_to_xlsx(
                pd.DataFrame({"a": ["x"]}), tmp_xlsx, validations={"a": {"type": "list", "values": []}},
            )

    def test_list_validation(self, tmp_xlsx: str) -> None:
        """Verify dropdown list validation."""
        df = pd.DataFrame({"Status": ["Open", "Closed"], "Value": [1, 2]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            validations={"Status": {"type": "list", "values": ["Open", "Closed", "Pending"]}},
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert len(ws.data_validations.dataValidation) > 0
        dv = ws.data_validations.dataValidation[0]
        # Assert the type and that the dropdown values actually landed.
        assert dv.type == "list"
        assert "Open" in dv.formula1
        assert "Pending" in dv.formula1
        wb.close()

    def test_number_validation(self, tmp_xlsx: str) -> None:
        """Verify whole number range validation."""
        df = pd.DataFrame({"Score": [85, 90]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert len(ws.data_validations.dataValidation) > 0
        dv = ws.data_validations.dataValidation[0]
        # Assert the type and the min/max bounds, not just presence.
        assert dv.type == "whole"
        assert dv.formula1 == "0"
        assert dv.formula2 == "100"
        wb.close()

    def test_validation_with_messages(self, tmp_xlsx: str) -> None:
        """Verify validation with input and error messages."""
        df = pd.DataFrame({"Value": [50]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            validations={
                "Value": {
                    "type": "decimal",
                    "min": 0,
                    "max": 100,
                    "input_title": "Enter Value",
                    "input_message": "Must be between 0 and 100",
                    "error_title": "Invalid",
                    "error_message": "Value out of range",
                }
            },
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert len(ws.data_validations.dataValidation) > 0
        dv = ws.data_validations.dataValidation[0]
        assert dv.promptTitle == "Enter Value"
        assert dv.errorTitle == "Invalid"
        wb.close()

    def test_whole_number_max_outside_i32_range_raises(self, tmp_xlsx: str) -> None:
        """A whole_number max outside the i32 range raises a clear, actionable error."""
        df = pd.DataFrame({"Score": [85]})
        with pytest.raises(ValueError, match="must be within the i32 range"):
            # Intentionally invalid: 3_000_000_000 exceeds i32::MAX.
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                validations={"Score": {"type": "whole_number", "min": 0, "max": 3_000_000_000}},
            )

    def test_whole_number_max_beyond_i64_raises_i32_range_message(self, tmp_xlsx: str) -> None:
        """A whole_number max too large even for i64 gets the i32-range message.

        Not the misleading generic 'must be an integer' fallback.
        """
        df = pd.DataFrame({"Score": [85]})
        with pytest.raises(ValueError, match="must be within the i32 range"):
            # Intentionally invalid: 2**70 overflows even i64, let alone i32.
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                validations={"Score": {"type": "whole_number", "min": 0, "max": 2**70}},
            )

    def test_validation_pattern_matching(self, tmp_xlsx: str) -> None:
        """Verify validation with column pattern."""
        df = pd.DataFrame({"score_a": [80], "score_b": [90], "name": ["Test"]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, validations={"score_*": {"type": "whole_number", "min": 0, "max": 100}}
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # One validation per score column, none on 'name'.
        assert sorted(str(dv.sqref) for dv in ws.data_validations.dataValidation) == ["A2", "B2"]
        wb.close()

    def test_validation_pattern_must_match(self, tmp_xlsx: str) -> None:
        """Validation patterns that match no columns raise an error."""
        df = pd.DataFrame({"Score": [85]})
        with pytest.raises(ValueError, match=r"validations.*Missing.*matched no columns"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                validations={"Missing": {"type": "whole_number", "min": 0, "max": 100}},
            )


class TestEmptyFrame:
    """A frame without data rows still has its validations checked; only the write is skipped."""

    @staticmethod
    def _empty() -> pd.DataFrame:
        return pd.DataFrame({"a": pd.Series([], dtype="int64")})

    def test_an_unknown_type_is_refused(self, tmp_xlsx: str) -> None:
        """A typo fails on the empty day, not first on the day with data."""
        bad: Any = {"type": "bogus"}
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"validations\['a'\].*Unknown validation type 'bogus'"):
            xlsxturbo.df_to_xlsx(self._empty(), tmp_xlsx, validations={"a": bad})

    def test_an_unknown_key_is_refused(self, tmp_xlsx: str) -> None:
        """Unknown-key rejection runs without rows."""
        bad: Any = {"type": "whole_number", "min": 0, "zzz": 1}
        with pytest.raises(xlsxturbo.ConfigurationError, match="unknown option 'zzz'"):
            xlsxturbo.df_to_xlsx(self._empty(), tmp_xlsx, validations={"a": bad})

    def test_a_pattern_must_still_match_a_column(self, tmp_xlsx: str) -> None:
        """An empty frame still has columns, so a pattern matching none of them is a typo."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"validations\['nosuch'\]: pattern matched no columns"):
            xlsxturbo.df_to_xlsx(self._empty(), tmp_xlsx, validations={"nosuch": {"type": "whole_number", "min": 0}})

    def test_a_valid_config_writes_nothing(self, tmp_xlsx: str) -> None:
        """A valid config on the empty frame exports, and no validation is written."""
        xlsxturbo.df_to_xlsx(self._empty(), tmp_xlsx, validations={"a": {"type": "whole_number", "min": 0}})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb).data_validations.dataValidation == []
        wb.close()
