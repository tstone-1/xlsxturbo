"""Tests that a cell reference must be written in ASCII.

A Unicode uppercase used to run before the ASCII-letter check, so a sharp s
followed by 1 was accepted as column SS, and a dotless i or a long s as I or S.
The non-ASCII letters are built with ``chr`` so the source stays ASCII.
"""

from __future__ import annotations

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")

SHARP_S = chr(0xDF)  # uppercases to "SS"
DOTLESS_I = chr(0x131)  # uppercases to "I"
LONG_S = chr(0x17F)  # uppercases to "S"


class TestNonAsciiCellReference:
    """A reference with a non-ASCII letter is a configuration error."""

    @pytest.mark.parametrize("ref", [f"{SHARP_S}1", f"{DOTLESS_I}1", f"{LONG_S}1", f"A{SHARP_S}1"])
    def test_comment_at_non_ascii_reference_is_refused(self, tmp_xlsx: str, ref: str) -> None:
        """`comments` names the bad key and raises ConfigurationError."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(xlsxturbo.ConfigurationError, match=r"Invalid (cell reference|row number)") as info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, comments={ref: "note"})
        assert f"comments['{ref}']" in str(info.value)

    @pytest.mark.parametrize("ref", ["SS1", "ss1", "I1", "S1"])
    def test_ascii_spelling_is_accepted(self, tmp_xlsx: str, ref: str) -> None:
        """Control: the ASCII references the old fold produced are fine, in either case."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, comments={ref: "note"})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        comment = ws[ref.upper()].comment
        assert comment is not None
        # The writer prefixes the note with an "Author:" line.
        assert comment.text.endswith("note")
        wb.close()
