"""Tests for annotation features: merged ranges, hyperlinks, comments, and defined names."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestMergedRanges:
    """Tests for merged_ranges feature (v0.9.0)."""

    def test_simple_merge(self, tmp_xlsx: str) -> None:
        """Merge a range with text."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4], "C": [5, 6]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            merged_ranges=[("A1:C1", "Title Row")],
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Cell A1 should contain the merge text
        assert ws["A1"].value == "Title Row"
        # The range should be merged
        merged = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:C1" in merged
        wb.close()

    def test_merge_with_format(self, tmp_xlsx: str) -> None:
        """Merge a range with custom formatting."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            merged_ranges=[
                ("A1:B1", "Styled Merge", {"bold": True, "bg_color": "#4F81BD"})
            ],
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].value == "Styled Merge"
        assert ws["A1"].font.bold is True
        merged = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:B1" in merged
        wb.close()

    def test_multiple_merges(self, tmp_xlsx: str) -> None:
        """Merge multiple ranges in the same sheet."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6], "C": [7, 8, 9]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            merged_ranges=[
                ("A1:C1", "Top Title"),
                ("A5:C5", "Bottom Title"),
            ],
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        merged = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:C1" in merged
        assert "A5:C5" in merged
        wb.close()

    def test_merge_with_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """Verify merged ranges work per-sheet."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.dfs_to_xlsx(
            [(df, "Sheet1", {"merged_ranges": [("A1:B1", "Per-Sheet Merge")]})],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "Per-Sheet Merge"
        merged = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:B1" in merged
        wb.close()


class TestHyperlinks:
    """Tests for hyperlinks feature (v0.9.0)."""

    def test_basic_hyperlink(self, tmp_xlsx: str) -> None:
        """Write a hyperlink with URL and display text."""
        df = pd.DataFrame({"Name": ["Example"]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            hyperlinks=[("B2", "https://example.com", "Example Site")],
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["B2"].hyperlink is not None
        assert "example.com" in ws["B2"].hyperlink.target
        wb.close()

    def test_hyperlink_without_display_text(self, tmp_xlsx: str) -> None:
        """Write a hyperlink with URL only (no display text)."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            hyperlinks=[("A2", "https://example.com")],
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].hyperlink is not None
        wb.close()

    def test_multiple_hyperlinks(self, tmp_xlsx: str) -> None:
        """Write multiple hyperlinks in the same sheet."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            hyperlinks=[
                ("B1", "https://one.com", "One"),
                ("B2", "https://two.com", "Two"),
                ("B3", "https://three.com", "Three"),
            ],
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["B1"].hyperlink is not None
        assert ws["B2"].hyperlink is not None
        assert ws["B3"].hyperlink is not None
        wb.close()

    def test_hyperlinks_with_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """Verify hyperlinks work per-sheet in multi-sheet mode."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.dfs_to_xlsx(
            [
                (
                    df,
                    "Sheet1",
                    {"hyperlinks": [("B1", "https://example.com", "Link")]},
                )
            ],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        ws = wb["Sheet1"]
        assert ws["B1"].hyperlink is not None
        wb.close()


class TestComments:
    """Tests for comments/notes feature (v0.10.0)."""

    def test_simple_comment(self, tmp_xlsx: str) -> None:
        """Write a simple string comment."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, comments={"A1": "This is a header note"})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # openpyxl stores comments in ws.comments
        assert ws["A1"].comment is not None
        assert "header note" in ws["A1"].comment.text
        wb.close()

    def test_comment_with_author(self, tmp_xlsx: str) -> None:
        """Write a comment with an author."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, comments={"A2": {"text": "Data note", "author": "John"}}
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        comment = ws["A2"].comment
        assert comment is not None
        assert "Data note" in comment.text
        assert comment.author == "John"
        wb.close()

    def test_comment_dict_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A typo'd/extra key in the comment dict form is rejected, not silently dropped."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="unknown option 'auhtor'"):
            # Intentionally invalid: 'auhtor' is a typo for 'author'.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, comments={"A1": {"text": "note", "auhtor": "John"}})  # type: ignore[typeddict-unknown-key]

    def test_multiple_comments(self, tmp_xlsx: str) -> None:
        """Write multiple comments on different cells."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, comments={"A1": "Column A", "B1": "Column B", "A2": "First value"}
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].comment is not None
        assert ws["B1"].comment is not None
        assert ws["A2"].comment is not None
        assert "Column A" in ws["A1"].comment.text
        wb.close()

    def test_empty_per_sheet_comments_overrides_global(self, tmp_xlsx: str) -> None:
        """An explicitly empty per-sheet 'comments' dict disables the global default for that sheet.

        Per-sheet complex options now shadow ("turn off") a non-empty global
        default when passed as an explicit empty dict/list, instead of
        silently falling back to the global value.
        """
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"A": [2]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "S1"),
                (df2, "S2", {"comments": {}}),
            ],
            tmp_xlsx,
            comments={"A1": "note"},
        )
        wb = load_workbook(tmp_xlsx)
        assert wb["S1"]["A1"].comment is not None
        assert wb["S2"]["A1"].comment is None
        wb.close()


class TestDefinedNames:
    """Tests for workbook-level defined names (v0.11.0)."""

    def test_single_defined_name(self, tmp_xlsx: str) -> None:
        """Create a single defined name in the workbook."""
        df = pd.DataFrame({"a": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx,
            defined_names={"MyRange": "=Sheet1!$A$1:$A$4"})
        wb = load_workbook(tmp_xlsx)
        names = {dn.name: dn.attr_text for dn in wb.defined_names.values()}
        assert "MyRange" in names
        assert "$A$1:$A$4" in names["MyRange"]
        wb.close()

    def test_multiple_defined_names(self, tmp_xlsx: str) -> None:
        """Create multiple defined names."""
        df = pd.DataFrame({"a": [1], "b": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, defined_names={
            "Range1": "=Sheet1!$A$1:$A$2",
            "Range2": "=Sheet1!$B$1:$B$2",
        })
        wb = load_workbook(tmp_xlsx)
        names = {dn.name for dn in wb.defined_names.values()}
        assert "Range1" in names
        assert "Range2" in names
        wb.close()

    def test_defined_names_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """Verify defined_names works in multi-sheet mode."""
        df1 = pd.DataFrame({"x": [1]})
        df2 = pd.DataFrame({"y": [2]})
        xlsxturbo.dfs_to_xlsx(
            [(df1, "S1"), (df2, "S2")], tmp_xlsx,
            defined_names={"AllData": "=S1!$A$1:$A$2"})
        wb = load_workbook(tmp_xlsx)
        names = {dn.name for dn in wb.defined_names.values()}
        assert "AllData" in names
        wb.close()

    def test_defined_name_with_quoted_sheet(self, tmp_xlsx: str) -> None:
        """Create a defined name with a quoted sheet name containing a space."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx,
            sheet_name="LCA Calculator Parameters",
            defined_names={"Settings": "='LCA Calculator Parameters'!$B$13:$D$155"})
        wb = load_workbook(tmp_xlsx)
        names = {dn.name for dn in wb.defined_names.values()}
        assert "Settings" in names
        wb.close()

    # An empty name used to abort the process: rust_xlsxwriter's define_name
    # called `chars().next().unwrap()` on the local part, so "" and "Sheet1!"
    # panicked instead of returning an error, and xlsxturbo screened for them
    # before the call. Fixed upstream in 0.98.2 (jmcnamara/rust_xlsxwriter#186);
    # the screen is gone and the crate's own ParameterError is what surfaces.
    # Pinned back to 0.98.1 these two fail: a Rust panic reaches Python as
    # pyo3_runtime.PanicException, which is not a ValueError.
    @pytest.mark.parametrize("name", ["", "Sheet1!"])
    def test_empty_defined_name_raises_df_to_xlsx(
        self, tmp_xlsx: str, name: str
    ) -> None:
        """An empty name and an empty local part both raise, neither panics."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError, match="cannot be empty in Excel"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                defined_names={name: "=Sheet1!$A$1:$A$2"},
            )

    def test_empty_defined_name_raises_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """The same failure reaches the caller in multi-sheet mode."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError, match="cannot be empty in Excel"):
            xlsxturbo.dfs_to_xlsx(
                [(df, "Sheet1")],
                tmp_xlsx,
                defined_names={"Sheet1!": "=Sheet1!$A$1:$A$2"},
            )

    def test_empty_defined_name_error_names_the_offending_key(
        self, tmp_xlsx: str
    ) -> None:
        """The message still identifies which `defined_names` key failed.

        The crate reports the *local* part, so its own message says `Name ''`
        for both `""` and `"Sheet1!"`. The wrapper in `apply_defined_names` is
        what puts the caller's own key back in front of it — without that, a
        dict of several names would not say which one was rejected.
        """
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError, match="cannot be empty in Excel") as excinfo:
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                defined_names={"Sheet1!": "=Sheet1!$A$1:$A$2"},
            )
        assert "Sheet1!" in str(excinfo.value)


class TestDefinedNameCellReferenceCollision:
    """A defined name Excel would read as a cell address is refused.

    Table names are sanitized instead, and the difference is deliberate: a
    defined name is what the user's own formulas say, so renaming ``Q1`` to
    ``_Q1`` behind their back would leave every ``=SUM(Q1)`` pointing at a name
    the workbook no longer defines. ``rust_xlsxwriter`` 0.98.2 validates the
    first character and the character set but not this.
    """

    @pytest.mark.parametrize("name", ["Q1", "A1", "A01", "XFD1048576", "R", "c", "R1C1"])
    def test_reference_shaped_name_raises(self, name: str, tmp_xlsx: str) -> None:
        """The error names the offending key and the rule."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(
            xlsxturbo.ConfigurationError, match="is a cell reference"
        ) as excinfo:
            xlsxturbo.df_to_xlsx(
                df, tmp_xlsx, defined_names={name: "=Sheet1!$A$1:$A$2"}
            )
        assert isinstance(excinfo.value, ValueError)
        assert f"defined_names['{name}']" in str(excinfo.value)

    def test_sheet_qualified_name_is_judged_on_its_local_part(
        self, tmp_xlsx: str
    ) -> None:
        """``Sheet1!Q1`` is rejected, and the message names both halves.

        ``define_name`` applies its own rules to the part after the ``!``, so a
        check that looked at the whole string would let this through.
        """
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(
            xlsxturbo.ConfigurationError, match="is a cell reference"
        ) as excinfo:
            xlsxturbo.df_to_xlsx(
                df, tmp_xlsx, defined_names={"Sheet1!Q1": "=Sheet1!$A$1:$A$2"}
            )
        assert isinstance(excinfo.value, ValueError)
        assert "defined_names['Sheet1!Q1']" in str(excinfo.value)
        assert "'Q1' is a cell reference" in str(excinfo.value)

    @pytest.mark.parametrize("name", ["MyRange", "Q1_Sales", "R1C1D"])
    def test_ordinary_names_still_work(self, name: str, tmp_xlsx: str) -> None:
        """An ordinary name is still accepted.

        The control: without it the rejections above are satisfied by a check
        that refuses everything.
        """
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, defined_names={name: "=Sheet1!$A$1:$A$2"})
        wb = load_workbook(tmp_xlsx)
        assert name in {dn.name for dn in wb.defined_names.values()}
        wb.close()

    def test_a_sheet_qualified_ordinary_name_still_works(self, tmp_xlsx: str) -> None:
        """A sheet-scoped name with a fine local part is still accepted.

        The qualified control: the local-part check must not over-reach into
        names that merely carry a sheet prefix.
        """
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, defined_names={"Sheet1!Totals": "=Sheet1!$A$1:$A$2"}
        )
        wb = load_workbook(tmp_xlsx)
        assert "Totals" in {
            dn.name for ws in wb.worksheets for dn in ws.defined_names.values()
        } | {dn.name for dn in wb.defined_names.values()}
        wb.close()

    @pytest.mark.parametrize("name", ["AAAA1", "A1048577"])
    def test_shapes_beyond_the_grid_stay_legal_names(
        self, name: str, tmp_xlsx: str
    ) -> None:
        """Letters-then-digits is not enough — the cell has to exist.

        The grid ends at XFD1048576, so these address nothing and Excel accepts
        them; refusing them would reject input Excel takes.
        """
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, defined_names={name: "=Sheet1!$A$1:$A$2"})
        wb = load_workbook(tmp_xlsx)
        assert name in {dn.name for dn in wb.defined_names.values()}
        wb.close()


class TestDefinedNamesVerification:
    """Tests for defined_names with content verification."""

    def test_defined_names_written(self, tmp_xlsx: str) -> None:
        """Verify defined_names are written to the workbook."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx,
            defined_names={"MyRange": "=Sheet1!$A$1:$A$4"},
        )
        wb = load_workbook(tmp_xlsx)
        assert "MyRange" in wb.defined_names
        wb.close()

    def test_defined_names_multi_sheet(self, tmp_xlsx: str) -> None:
        """Verify defined_names work with dfs_to_xlsx."""
        df = pd.DataFrame({"A": [1, 2]})
        xlsxturbo.dfs_to_xlsx(
            [(df, "Data")], tmp_xlsx,
            defined_names={"Total": "=Data!$A$1:$A$3"},
        )
        wb = load_workbook(tmp_xlsx)
        assert "Total" in wb.defined_names
        wb.close()
