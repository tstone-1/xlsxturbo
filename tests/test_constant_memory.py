"""Tests for constant_memory mode and its warning emission."""

from __future__ import annotations

import warnings

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestConstantMemoryMode:
    """Tests for constant_memory mode (v0.4.0)."""

    def test_basic_constant_memory(self, tmp_xlsx: str) -> None:
        """File is created in constant memory mode."""
        df = pd.DataFrame({"A": list(range(100)), "B": list(range(100, 200))})
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx, constant_memory=True)
        assert rows > 0
        assert cols == 2
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].value == "A"  # header
        assert ws["A2"].value == 0  # first data row
        assert ws["B2"].value == 100
        wb.close()

    def test_constant_memory_warns_when_disabling_features(self, tmp_xlsx: str) -> None:
        """Features are disabled with one actionable warning and no crash."""
        df = pd.DataFrame({"Score": [1, 2, 3]})
        with pytest.warns(RuntimeWarning, match="constant_memory=True disables these features"):
            rows, _cols = xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                constant_memory=True,
                table_style="Medium9",
                freeze_panes=True,
                autofit=True,
                row_heights={0: 30},
                conditional_formats={"Score": {"type": "data_bar", "bar_color": "#638EC6"}},
                formula_columns={"Double": "=A{row}*2"},
                merged_ranges=[("A1:A1", "Merge")],
                hyperlinks=[("A1", "https://example.com", "Link")],
                comments={"A1": "Comment"},
                validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
                rich_text={"B1": [("Bold", {"bold": True})]},
            )
        assert rows > 0
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Table should NOT be created
        assert len(ws.tables) == 0
        # Data should still be written
        assert ws["A1"].value == "Score"
        assert ws["A2"].value == 1
        wb.close()

    def test_constant_memory_with_column_widths(self, tmp_xlsx: str) -> None:
        """Column widths still work in constant memory mode."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, constant_memory=True, column_widths={0: 25})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        width = ws.column_dimensions["A"].width
        assert width is not None
        assert width > 20
        wb.close()


class TestConstantMemoryWarning:
    """Tests for constant_memory warning emission."""

    def test_constant_memory_warns_on_incompatible_options(self, tmp_xlsx: str) -> None:
        """constant_memory=True with incompatible options emits RuntimeWarning."""
        df = pd.DataFrame({"A": [1, 2]})
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            xlsxturbo.df_to_xlsx(
                df, tmp_xlsx,
                constant_memory=True,
                table_style="Medium2",
                freeze_panes=True,
            )
            assert len(w) == 1
            assert issubclass(w[0].category, RuntimeWarning)
            assert "table_style" in str(w[0].message)
            assert "freeze_panes" in str(w[0].message)

    def test_constant_memory_no_warning_when_clean(self, tmp_xlsx: str) -> None:
        """constant_memory=True without incompatible options emits no warning."""
        df = pd.DataFrame({"A": [1, 2]})
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, constant_memory=True)
            assert len(w) == 0

    def test_constant_memory_empty_per_sheet_dict_excluded_from_warning(self, tmp_xlsx: str) -> None:
        """An explicitly empty per-sheet dict is a no-op and must not appear in the disabled-features list.

        `present_complex_options` excludes present-but-empty collections, since
        they have nothing to apply and thus nothing for constant_memory to
        skip. `freeze_panes=True` is included alongside the empty `comments`
        dict so the warning still fires (proving the empty dict isn't just
        suppressing the whole warning) while confirming 'comments' is absent
        from its message.
        """
        df = pd.DataFrame({"A": [1, 2]})
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            xlsxturbo.dfs_to_xlsx(
                [(df, "S1", {"comments": {}})],
                tmp_xlsx,
                constant_memory=True,
                freeze_panes=True,
            )
            assert len(w) == 1
            message = str(w[0].message)
            assert "freeze_panes" in message
            assert "comments" not in message

    def test_dfs_to_xlsx_two_sheets_each_warn_once(self, tmp_xlsx: str) -> None:
        """dfs_to_xlsx with constant_memory=True warns once per sheet that has a disabled feature.

        The skip-warning is emitted inside write_sheet_data, which runs once
        per sheet, so table_style set on both sheets of a two-sheet workbook
        produces two separate RuntimeWarnings (one per sheet), not a single
        workbook-level warning.
        """
        df1 = pd.DataFrame({"A": [1, 2]})
        df2 = pd.DataFrame({"B": [3, 4]})
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            xlsxturbo.dfs_to_xlsx(
                [(df1, "S1"), (df2, "S2")],
                tmp_xlsx,
                constant_memory=True,
                table_style="Medium2",
            )
            assert len(w) == 2
            for warning in w:
                assert issubclass(warning.category, RuntimeWarning)
                assert "table_style" in str(warning.message)
        wb = load_workbook(tmp_xlsx)
        assert wb["S1"]["A2"].value == 1
        assert wb["S2"]["A2"].value == 3
        assert len(wb["S1"].tables) == 0
        assert len(wb["S2"].tables) == 0
        wb.close()


class TestFeatureConstantMemoryWarnings:
    """Tests that each individually-incompatible feature warns by name."""

    @pytest.mark.parametrize(
        ("option_name", "kwargs"),
        [
            pytest.param("cells", {"cells": {"C1": "test"}}, id="cells"),
            pytest.param("checkboxes", {"checkboxes": {"B2": True}}, id="checkboxes"),
            pytest.param("textboxes", {"textboxes": {"B2": "note"}}, id="textboxes"),
            pytest.param(
                "charts",
                {"charts": {"D2": {"type": "bar", "data_range": "Sheet1!$A$2:$A$3"}}},
                id="charts",
            ),
            pytest.param(
                "sparklines",
                {"sparklines": {"D2": {"range": "Sheet1!A2:A3"}}},
                id="sparklines",
            ),
        ],
    )
    def test_feature_constant_memory_warns(
        self, option_name: str, kwargs: dict[str, object], tmp_xlsx: str
    ) -> None:
        """constant_memory=True with an incompatible feature emits a RuntimeWarning naming it."""
        df = pd.DataFrame({"A": [1, 2]})
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, constant_memory=True, **kwargs)  # type: ignore[arg-type]
            assert len(w) == 1
            assert issubclass(w[0].category, RuntimeWarning)
            assert option_name in str(w[0].message)
