"""Option-coverage guard (W4): every per-sheet option must actually be applied.

Nothing previously forced a sheet option to have a real effect on the written
file -- an option can be accepted by validation, extracted into a typed Rust
struct, and then silently dropped if its ``apply_*`` call is missing from
``apply_worksheet_features`` in ``src/convert.rs``. This module writes a
minimal workbook per option and asserts an observable artifact in the
produced ``.xlsx`` (via openpyxl or a raw zip/XML read-back), so a future
change that wires an option through extraction/validation but forgets to
apply it fails a test instead of shipping silently.

The ``test_coverage_map_is_complete`` test is the actual guard: it
introspects ``xlsxturbo.df_to_xlsx``'s real keyword parameters at runtime
(the same ``inspect.signature`` mechanism used in ``test_api_surface.py``)
and asserts every sheet-option kwarg has an entry in ``COVERAGE`` -- so
adding option #N+1 without a coverage entry is a test failure, not a
silent gap.
"""

from __future__ import annotations

import base64
import inspect
import zipfile
from collections.abc import Callable
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, TINY_PNG_B64, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")

# kwargs of df_to_xlsx that are not per-sheet "options": the DataFrame, the
# output path, the (single-sheet) sheet name, and defined_names, which is a
# workbook-level feature applied once regardless of which sheet(s) exist,
# not a per-sheet option accepted by dfs_to_xlsx's options dict.
NON_SHEET_PARAMS = frozenset({"df", "output_path", "sheet_name", "defined_names"})

# A factory that creates and tracks a new temporary file path (see
# conftest.py's `tmp_xlsx_factory` fixture); used by checks that need an
# extra file (e.g. an image) alongside the workbook under test.
PathFactory = Callable[..., str]


def _base_df() -> pd.DataFrame:
    """Return a fresh minimal DataFrame reused by most coverage checks.

    A new DataFrame is built per call so no check can accidentally mutate a
    shared instance out from under another.

    Returns:
        A 3-row, 2-column DataFrame with columns "Name" and "Score".
    """
    return pd.DataFrame({"Name": ["Alice", "Bob", "Carol"], "Score": [10, 50, 90]})


def _check_header(path: str, _factory: PathFactory) -> None:
    """header=False must omit the header row entirely."""
    xlsxturbo.df_to_xlsx(_base_df(), path, header=False)
    ws = active_ws(load_workbook(path))
    assert ws["A1"].value == "Alice"


def _check_autofit(path: str, _factory: PathFactory) -> None:
    """autofit=True must widen a column beyond Excel's default width."""
    df = pd.DataFrame({"VeryLongColumnNameForAutofit": ["x" * 80]})
    xlsxturbo.df_to_xlsx(df, path, autofit=True)
    ws = active_ws(load_workbook(path))
    width = ws.column_dimensions["A"].width
    assert width is not None
    assert width > 15, f"autofitted width {width} should exceed the Excel default"


def _check_table_style(path: str, _factory: PathFactory) -> None:
    """table_style must add a real Excel table to the sheet."""
    xlsxturbo.df_to_xlsx(_base_df(), path, table_style="Medium9")
    ws = active_ws(load_workbook(path))
    assert len(ws.tables) == 1


def _check_table_name(path: str, _factory: PathFactory) -> None:
    """table_name must be used as the created table's name."""
    xlsxturbo.df_to_xlsx(_base_df(), path, table_style="Medium9", table_name="CoverageTable")
    ws = active_ws(load_workbook(path))
    assert "CoverageTable" in ws.tables


def _check_freeze_panes(path: str, _factory: PathFactory) -> None:
    """freeze_panes=True must freeze the header row."""
    xlsxturbo.df_to_xlsx(_base_df(), path, freeze_panes=True)
    ws = active_ws(load_workbook(path))
    assert ws.freeze_panes == "A2"


def _check_constant_memory(path: str, _factory: PathFactory) -> None:
    """constant_memory=True must emit a RuntimeWarning naming a skipped feature."""
    with pytest.warns(RuntimeWarning, match="constant_memory=True disables these features"):
        xlsxturbo.df_to_xlsx(_base_df(), path, constant_memory=True, table_style="Medium9")
    assert Path(path).exists()


def _check_column_widths(path: str, _factory: PathFactory) -> None:
    """column_widths must set an explicit column width."""
    xlsxturbo.df_to_xlsx(_base_df(), path, column_widths={0: 40})
    ws = active_ws(load_workbook(path))
    width = ws.column_dimensions["A"].width
    assert width is not None
    assert width > 30


def _check_row_heights(path: str, _factory: PathFactory) -> None:
    """row_heights must set an explicit row height."""
    xlsxturbo.df_to_xlsx(_base_df(), path, row_heights={0: 40})
    ws = active_ws(load_workbook(path))
    height = ws.row_dimensions[1].height
    assert height is not None
    assert abs(height - 40) < 1


def _check_header_format(path: str, _factory: PathFactory) -> None:
    """header_format must format the header row."""
    xlsxturbo.df_to_xlsx(_base_df(), path, header_format={"bold": True})
    ws = active_ws(load_workbook(path))
    assert ws["A1"].font.bold is True


def _check_column_formats(path: str, _factory: PathFactory) -> None:
    """column_formats must format the matched column's data cells."""
    xlsxturbo.df_to_xlsx(_base_df(), path, column_formats={"Score": {"bold": True}})
    ws = active_ws(load_workbook(path))
    assert ws["B2"].font.bold is True


def _check_conditional_formats(path: str, _factory: PathFactory) -> None:
    """conditional_formats must add a real conditional format rule to the sheet."""
    xlsxturbo.df_to_xlsx(
        _base_df(),
        path,
        conditional_formats={"Score": {"type": "data_bar", "bar_color": "#638EC6"}},
    )
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8").upper()
        assert "DATABAR" in xml
        assert "638EC6" in xml


def _check_formula_columns(path: str, _factory: PathFactory) -> None:
    """formula_columns must add a computed column with a real Excel formula."""
    xlsxturbo.df_to_xlsx(_base_df(), path, formula_columns={"Double": "=B{row}*2"})
    ws = active_ws(load_workbook(path))
    assert ws["C1"].value == "Double"
    assert ws["C2"].value == "=B2*2"


def _check_merged_ranges(path: str, _factory: PathFactory) -> None:
    """merged_ranges must merge the given range and write its text."""
    xlsxturbo.df_to_xlsx(_base_df(), path, merged_ranges=[("D1:E1", "Merged Title")])
    ws = active_ws(load_workbook(path))
    assert ws["D1"].value == "Merged Title"
    assert "D1:E1" in [str(m) for m in ws.merged_cells.ranges]


def _check_hyperlinks(path: str, _factory: PathFactory) -> None:
    """Hyperlinks must attach a real hyperlink to the target cell."""
    xlsxturbo.df_to_xlsx(_base_df(), path, hyperlinks=[("D1", "https://example.com")])
    ws = active_ws(load_workbook(path))
    assert ws["D1"].hyperlink is not None
    assert "example.com" in ws["D1"].hyperlink.target


def _check_comments(path: str, _factory: PathFactory) -> None:
    """Comments must attach a real cell comment/note."""
    xlsxturbo.df_to_xlsx(_base_df(), path, comments={"D1": "Coverage note"})
    ws = active_ws(load_workbook(path))
    assert ws["D1"].comment is not None
    assert "Coverage note" in ws["D1"].comment.text


def _check_validations(path: str, _factory: PathFactory) -> None:
    """Validations must add a real data validation rule to the sheet."""
    xlsxturbo.df_to_xlsx(
        _base_df(),
        path,
        validations={"Name": {"type": "text_length", "min": 1, "max": 50}},
    )
    ws = active_ws(load_workbook(path))
    assert len(ws.data_validations.dataValidation) > 0


def _check_rich_text(path: str, _factory: PathFactory) -> None:
    """rich_text must write a formatted run, not plain text."""
    xlsxturbo.df_to_xlsx(_base_df(), path, rich_text={"D1": [("Bold", {"bold": True}), " plain"]})
    with zipfile.ZipFile(path) as zf:
        shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
        assert "Bold" in shared
        assert "<b/>" in shared


def _check_images(path: str, factory: PathFactory) -> None:
    """Images must embed a real image in the workbook package."""
    png_path = factory(".png")
    Path(png_path).write_bytes(base64.b64decode(TINY_PNG_B64))
    xlsxturbo.df_to_xlsx(_base_df(), path, images={"D1": png_path})
    with zipfile.ZipFile(path) as zf:
        media = [n for n in zf.namelist() if n.startswith("xl/media/")]
        assert media, "no embedded image found in xl/media/"


def _check_checkboxes(path: str, _factory: PathFactory) -> None:
    """Checkboxes must write a real Excel boolean checkbox cell."""
    xlsxturbo.df_to_xlsx(_base_df(), path, checkboxes={"D1": True})
    ws = active_ws(load_workbook(path))
    assert ws["D1"].value is True


def _check_textboxes(path: str, _factory: PathFactory) -> None:
    """Textboxes must add a floating text shape with the given text."""
    xlsxturbo.df_to_xlsx(_base_df(), path, textboxes={"D1": "Coverage textbox"})
    with zipfile.ZipFile(path) as zf:
        drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        assert "Coverage textbox" in drawing


def _check_charts(path: str, _factory: PathFactory) -> None:
    """Charts must add a real native Excel chart part."""
    xlsxturbo.df_to_xlsx(
        _base_df(),
        path,
        charts={"D2": {"type": "bar", "data_range": "Sheet1!$B$2:$B$4"}},
    )
    with zipfile.ZipFile(path) as zf:
        assert "xl/charts/chart1.xml" in zf.namelist()


def _check_sparklines(path: str, _factory: PathFactory) -> None:
    """Sparklines must add a real sparkline group to the worksheet XML."""
    xlsxturbo.df_to_xlsx(_base_df(), path, sparklines={"D2": {"range": "Sheet1!A2:B2"}})
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "<x14:sparklineGroup" in xml


def _check_cells(path: str, _factory: PathFactory) -> None:
    """Cells must write the given value to the given cell."""
    xlsxturbo.df_to_xlsx(_base_df(), path, cells={"D1": "Coverage label"})
    ws = active_ws(load_workbook(path))
    assert ws["D1"].value == "Coverage label"


# Option name -> a callable that writes a minimal workbook exercising that
# option and asserts its observable effect. Every df_to_xlsx sheet-option
# kwarg must have exactly one entry here (see test_coverage_map_is_complete).
COVERAGE: dict[str, Callable[[str, PathFactory], None]] = {
    "header": _check_header,
    "autofit": _check_autofit,
    "table_style": _check_table_style,
    "table_name": _check_table_name,
    "freeze_panes": _check_freeze_panes,
    "constant_memory": _check_constant_memory,
    "column_widths": _check_column_widths,
    "row_heights": _check_row_heights,
    "header_format": _check_header_format,
    "column_formats": _check_column_formats,
    "conditional_formats": _check_conditional_formats,
    "formula_columns": _check_formula_columns,
    "merged_ranges": _check_merged_ranges,
    "hyperlinks": _check_hyperlinks,
    "comments": _check_comments,
    "validations": _check_validations,
    "rich_text": _check_rich_text,
    "images": _check_images,
    "checkboxes": _check_checkboxes,
    "textboxes": _check_textboxes,
    "charts": _check_charts,
    "sparklines": _check_sparklines,
    "cells": _check_cells,
}


@pytest.mark.parametrize("option_name", sorted(COVERAGE))
def test_option_is_applied(option_name: str, tmp_xlsx_factory: PathFactory) -> None:
    """Each per-sheet option produces an observable effect in the written file.

    Args:
        option_name: The option under test, used to look up its check in COVERAGE.
        tmp_xlsx_factory: Fixture that creates and cleans up temporary file paths.
    """
    path = tmp_xlsx_factory()
    COVERAGE[option_name](path, tmp_xlsx_factory)


def test_coverage_map_is_complete() -> None:
    """Every df_to_xlsx sheet-option kwarg must have a COVERAGE entry.

    This is the actual W4 guard: introspects the compiled `df_to_xlsx`
    function's real keyword parameters (the same mechanism
    `test_api_surface.py` uses) rather than trusting a hand-maintained list,
    so a new option added to the signature without a matching coverage
    function fails this test immediately -- it can no longer be silently
    accepted, extracted, and dropped before reaching `apply_worksheet_features`.
    """
    params = set(inspect.signature(xlsxturbo.df_to_xlsx).parameters.keys())
    sheet_options = params - NON_SHEET_PARAMS

    missing = sheet_options - set(COVERAGE)
    extra = set(COVERAGE) - sheet_options
    assert not missing, f"df_to_xlsx option(s) with no coverage test: {sorted(missing)}"
    assert not extra, f"COVERAGE has entries for nonexistent option(s): {sorted(extra)}"
