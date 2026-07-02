"""Integration tests exercising multiple xlsxturbo features together."""

from __future__ import annotations

import base64
from collections.abc import Callable
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, TINY_PNG_B64, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestAllFeaturesCombined:
    """Test using all new features together."""

    def test_all_features_df_to_xlsx(self, tmp_xlsx: str) -> None:
        """All features work together in df_to_xlsx."""
        df = pd.DataFrame(
            {"Name": ["Alice", "Bob"], "Score": [95, 87], "Grade": ["A", "B"]}
        )
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            autofit=True,
            table_style="Medium2",
            table_name="StudentScores",
            column_widths={"_all": 30, 0: 20},
            header_format={"bold": True, "bg_color": "#4F81BD", "font_color": "white"},
            freeze_panes=True,
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert "StudentScores" in ws.tables
        # freeze_panes=True must freeze the header row (split below row 1).
        assert ws.freeze_panes == "A2"
        # Note: table_style overrides header_format styling
        # This is expected Excel behavior - tables have their own header styles
        wb.close()

    def test_all_features_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """All features work together in dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        df2 = pd.DataFrame({"X": ["a", "b"], "Y": ["c", "d"]})
        xlsxturbo.dfs_to_xlsx(
            [
                (
                    df1,
                    "Numbers",
                    {
                        "table_style": "Medium2",
                        "table_name": "NumbersTable",
                        "header_format": {"bold": True},
                        "column_widths": {"_all": 15},
                    },
                ),
                (
                    df2,
                    "Letters",
                    {
                        "table_style": "Medium9",
                        "table_name": "LettersTable",
                        "header_format": {"bg_color": "#FF0000"},
                    },
                ),
            ],
            tmp_xlsx,
            autofit=True,
            freeze_panes=True,
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        assert "NumbersTable" in wb["Numbers"].tables
        assert "LettersTable" in wb["Letters"].tables
        wb.close()


class TestV10AllFeatures:
    """Tests combining v0.10.0 features."""

    def test_all_new_features_together(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """All v0.10.0 features work together."""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [85, 92]})
        path = tmp_xlsx_factory()

        png_data = base64.b64decode(TINY_PNG_B64)
        img_path = tmp_xlsx_factory(".png")
        with Path(img_path).open("wb") as f:
            f.write(png_data)

        xlsxturbo.df_to_xlsx(
            df,
            path,
            comments={"A1": "Names column", "B1": {"text": "Scores", "author": "Test"}},
            validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
            rich_text={"D1": [("Legend:", {"bold": True}), " student scores"]},
            images={"E1": img_path},
        )
        assert Path(path).exists()

    def test_new_features_with_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """New features work with dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": [1, 2]})
        df2 = pd.DataFrame({"B": [3, 4]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "Sheet1", {"comments": {"A1": "First sheet header"}}),
                (df2, "Sheet2", {"validations": {"B": {"type": "whole_number", "min": 0, "max": 10}}}),
            ],
            tmp_xlsx,
        )
        assert Path(tmp_xlsx).exists()
