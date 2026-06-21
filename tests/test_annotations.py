"""Tests for annotation features: merged ranges, hyperlinks, comments, and defined names."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, get_temp_path, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestMergedRanges:
    """Tests for merged_ranges feature (v0.9.0)."""

    def test_simple_merge(self) -> None:
        """Merge a range with text."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4], "C": [5, 6]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                merged_ranges=[("A1:C1", "Title Row")],
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Cell A1 should contain the merge text
                assert ws["A1"].value == "Title Row"
                # The range should be merged
                merged = [str(m) for m in ws.merged_cells.ranges]
                assert "A1:C1" in merged
                wb.close()
        finally:
            Path(path).unlink()

    def test_merge_with_format(self) -> None:
        """Merge a range with custom formatting."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                merged_ranges=[
                    ("A1:B1", "Styled Merge", {"bold": True, "bg_color": "#4F81BD"})
                ],
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].value == "Styled Merge"
                assert ws["A1"].font.bold is True
                merged = [str(m) for m in ws.merged_cells.ranges]
                assert "A1:B1" in merged
                wb.close()
        finally:
            Path(path).unlink()

    def test_multiple_merges(self) -> None:
        """Merge multiple ranges in the same sheet."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6], "C": [7, 8, 9]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                merged_ranges=[
                    ("A1:C1", "Top Title"),
                    ("A5:C5", "Bottom Title"),
                ],
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                merged = [str(m) for m in ws.merged_cells.ranges]
                assert "A1:C1" in merged
                assert "A5:C5" in merged
                wb.close()
        finally:
            Path(path).unlink()

    def test_merge_with_dfs_to_xlsx(self) -> None:
        """Verify merged ranges work per-sheet."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df, "Sheet1", {"merged_ranges": [("A1:B1", "Per-Sheet Merge")]})],
                path,
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb["Sheet1"]
                assert ws["A1"].value == "Per-Sheet Merge"
                merged = [str(m) for m in ws.merged_cells.ranges]
                assert "A1:B1" in merged
                wb.close()
        finally:
            Path(path).unlink()


class TestHyperlinks:
    """Tests for hyperlinks feature (v0.9.0)."""

    def test_basic_hyperlink(self) -> None:
        """Write a hyperlink with URL and display text."""
        df = pd.DataFrame({"Name": ["Example"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                hyperlinks=[("B2", "https://example.com", "Example Site")],
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["B2"].hyperlink is not None
                assert "example.com" in ws["B2"].hyperlink.target
                wb.close()
        finally:
            Path(path).unlink()

    def test_hyperlink_without_display_text(self) -> None:
        """Write a hyperlink with URL only (no display text)."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                hyperlinks=[("A2", "https://example.com")],
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].hyperlink is not None
                wb.close()
        finally:
            Path(path).unlink()

    def test_multiple_hyperlinks(self) -> None:
        """Write multiple hyperlinks in the same sheet."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                hyperlinks=[
                    ("B1", "https://one.com", "One"),
                    ("B2", "https://two.com", "Two"),
                    ("B3", "https://three.com", "Three"),
                ],
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["B1"].hyperlink is not None
                assert ws["B2"].hyperlink is not None
                assert ws["B3"].hyperlink is not None
                wb.close()
        finally:
            Path(path).unlink()

    def test_hyperlinks_with_dfs_to_xlsx(self) -> None:
        """Verify hyperlinks work per-sheet in multi-sheet mode."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (
                        df,
                        "Sheet1",
                        {"hyperlinks": [("B1", "https://example.com", "Link")]},
                    )
                ],
                path,
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb["Sheet1"]
                assert ws["B1"].hyperlink is not None
                wb.close()
        finally:
            Path(path).unlink()


class TestComments:
    """Tests for comments/notes feature (v0.10.0)."""

    def test_simple_comment(self) -> None:
        """Write a simple string comment."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, comments={"A1": "This is a header note"})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # openpyxl stores comments in ws.comments
                assert ws["A1"].comment is not None
                assert "header note" in ws["A1"].comment.text
                wb.close()
        finally:
            Path(path).unlink()

    def test_comment_with_author(self) -> None:
        """Write a comment with an author."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, comments={"A2": {"text": "Data note", "author": "John"}}
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                comment = ws["A2"].comment
                assert comment is not None
                assert "Data note" in comment.text
                assert comment.author == "John"
                wb.close()
        finally:
            Path(path).unlink()

    def test_multiple_comments(self) -> None:
        """Write multiple comments on different cells."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, comments={"A1": "Column A", "B1": "Column B", "A2": "First value"}
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].comment is not None
                assert ws["B1"].comment is not None
                assert ws["A2"].comment is not None
                assert "Column A" in ws["A1"].comment.text
                wb.close()
        finally:
            Path(path).unlink()


class TestDefinedNames:
    """Tests for workbook-level defined names (v0.11.0)."""

    def test_single_defined_name(self) -> None:
        """Create a single defined name in the workbook."""
        df = pd.DataFrame({"a": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path,
                defined_names={"MyRange": "=Sheet1!$A$1:$A$4"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                names = {dn.name: dn.attr_text for dn in wb.defined_names.values()}
                assert "MyRange" in names
                assert "$A$1:$A$4" in names["MyRange"]
                wb.close()
        finally:
            Path(path).unlink()

    def test_multiple_defined_names(self) -> None:
        """Create multiple defined names."""
        df = pd.DataFrame({"a": [1], "b": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, defined_names={
                "Range1": "=Sheet1!$A$1:$A$2",
                "Range2": "=Sheet1!$B$1:$B$2",
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                names = {dn.name for dn in wb.defined_names.values()}
                assert "Range1" in names
                assert "Range2" in names
                wb.close()
        finally:
            Path(path).unlink()

    def test_defined_names_dfs_to_xlsx(self) -> None:
        """Verify defined_names works in multi-sheet mode."""
        df1 = pd.DataFrame({"x": [1]})
        df2 = pd.DataFrame({"y": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df1, "S1"), (df2, "S2")], path,
                defined_names={"AllData": "=S1!$A$1:$A$2"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                names = {dn.name for dn in wb.defined_names.values()}
                assert "AllData" in names
                wb.close()
        finally:
            Path(path).unlink()

    def test_defined_name_with_quoted_sheet(self) -> None:
        """Create a defined name with a quoted sheet name containing a space."""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path,
                sheet_name="LCA Calculator Parameters",
                defined_names={"Settings": "='LCA Calculator Parameters'!$B$13:$D$155"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                names = {dn.name for dn in wb.defined_names.values()}
                assert "Settings" in names
                wb.close()
        finally:
            Path(path).unlink()

    def test_empty_local_defined_name_raises_df_to_xlsx(self) -> None:
        """Empty local defined names raise ValueError instead of panicking."""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="name must not be empty"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    defined_names={"Sheet1!": "=Sheet1!$A$1:$A$2"},
                )
        finally:
            Path(path).unlink(missing_ok=True)

    def test_empty_local_defined_name_raises_dfs_to_xlsx(self) -> None:
        """Empty local defined names raise ValueError in multi-sheet mode."""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="name must not be empty"):
                xlsxturbo.dfs_to_xlsx(
                    [(df, "Sheet1")],
                    path,
                    defined_names={"Sheet1!": "=Sheet1!$A$1:$A$2"},
                )
        finally:
            Path(path).unlink(missing_ok=True)


class TestDefinedNamesVerification:
    """Tests for defined_names with content verification."""

    def test_defined_names_written(self) -> None:
        """Verify defined_names are written to the workbook."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path,
                defined_names={"MyRange": "=Sheet1!$A$1:$A$4"},
            )
            wb = load_workbook(path)
            assert "MyRange" in wb.defined_names
            wb.close()
        finally:
            Path(path).unlink()

    def test_defined_names_multi_sheet(self) -> None:
        """Verify defined_names work with dfs_to_xlsx."""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df, "Data")], path,
                defined_names={"Total": "=Data!$A$1:$A$3"},
            )
            wb = load_workbook(path)
            assert "Total" in wb.defined_names
            wb.close()
        finally:
            Path(path).unlink()
