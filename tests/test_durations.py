"""Tests for duration values: a missing duration is an empty cell in every frame shape.

A pandas frame whose columns are all timedeltas reaches the writer as numpy
``timedelta64`` scalars, where any other column mix yields pandas ``Timedelta``
objects and pandas' ``NaT``. The missing value of the first kind used to be
written as the text ``NaT``, in every unit.

What a duration that is present becomes is not pinned here: it depends on the
frame's dtype mix and on the column's unit (a nanosecond column of an
all-timedelta frame is written as a number, every other unit as text), and
choosing one form is an open decision. The unit is set explicitly for that
reason: pandas 2 defaults to nanoseconds and pandas 3 to microseconds.
"""

from __future__ import annotations

import datetime

import numpy as np
import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")

# The missing value sits in the middle row, so its row exists in the sheet
# whichever way the writer handles it: a trailing empty row would simply be
# absent and read back as None either way.
_TD = pd.to_timedelta(["1 days", "NaT", "01:30:00"])
_PY_TD = [datetime.timedelta(days=1), None, datetime.timedelta(hours=1, minutes=30)]
_NUMPY_SCALARS = [np.timedelta64(1, "D"), np.timedelta64("NaT", "ns"), np.timedelta64(2, "h")]


def _in(unit: str) -> pd.Index:
    return _TD.astype(f"timedelta64[{unit}]")


def _frames() -> list[object]:
    return [
        *(
            pytest.param(pd.DataFrame({"td": _in(unit)}), id=f"pandas-all-timedelta-{unit}")
            for unit in ("s", "ms", "us", "ns")
        ),
        pytest.param(pd.DataFrame({"td": _in("us"), "u": _in("us")}), id="pandas-two-timedelta-columns"),
        pytest.param(
            pd.DataFrame({"td": pd.Series(_NUMPY_SCALARS, dtype=object)}),
            id="pandas-object-column-of-numpy-scalars",
        ),
        # Controls: these shapes already wrote an empty cell.
        pytest.param(pd.DataFrame({"td": _in("us"), "s": ["a", "b", "c"]}), id="pandas-timedelta-and-str"),
        pytest.param(pd.DataFrame({"td": pd.Series(_PY_TD, dtype=object)}), id="pandas-object-column-of-timedelta"),
        pytest.param(pl.DataFrame({"td": _PY_TD}), id="polars-duration"),
    ]


class TestMissingDuration:
    """A missing duration is written as an empty cell."""

    @pytest.mark.parametrize("frame", _frames())
    @pytest.mark.parametrize("constant_memory", [False, True])
    def test_missing_duration_is_an_empty_cell(self, tmp_xlsx: str, frame: object, constant_memory: bool) -> None:
        """The NaT row is empty and the rows around it still carry a value."""
        xlsxturbo.df_to_xlsx(frame, tmp_xlsx, constant_memory=constant_memory)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].value == "td"
        assert ws["A3"].value is None
        for ref in ("A2", "A4"):
            value = ws[ref].value
            assert value is not None, ref
            assert value != "", ref
        wb.close()
