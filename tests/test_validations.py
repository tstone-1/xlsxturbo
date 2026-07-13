"""Tests for the data validation feature (dropdown lists, numeric ranges, messages)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestValidations:
    """Tests for data validation feature (v0.10.0)."""

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
        # Should have validations on the score columns
        assert len(ws.data_validations.dataValidation) > 0
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
