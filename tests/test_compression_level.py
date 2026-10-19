"""Tests for ``compression_level``: the zip compression of the saved file.

The level changes how the archive is compressed and nothing inside it, so the
checks read the compressed size of the worksheet part and compare the parts'
uncompressed bytes. ``docProps/core.xml`` is left out of the byte comparison: it
carries a timestamp (see ``TestGeneratedFileDeterminism``).
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
import pytest
import xlsxturbo
from xlsxturbo import cli

SHEET = "xl/worksheets/sheet1.xml"


def _frame() -> pd.DataFrame:
    """Return a frame large enough that the compression level changes the size.

    Returns:
        5000 rows of numbers and repeated strings.
    """
    return pd.DataFrame({"n": range(5000), "s": [f"row {i % 50}" for i in range(5000)]})


def _packed(path: Path) -> int:
    """Return the compressed size of the first worksheet.

    Args:
        path: The workbook.

    Returns:
        The worksheet part's size inside the archive, in bytes.
    """
    with zipfile.ZipFile(path) as archive:
        return archive.getinfo(SHEET).compress_size


def _parts(path: Path) -> dict[str, bytes]:
    """Return every part's uncompressed bytes, except the timestamped core properties.

    Args:
        path: The workbook.

    Returns:
        Part name to content.
    """
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


@pytest.fixture
def csv_file(tmp_path: Path) -> Path:
    """Write the frame as a CSV for the CSV entry points.

    Args:
        tmp_path: pytest's per-test folder.

    Returns:
        The CSV path.
    """
    path = tmp_path / "in.csv"
    _frame().to_csv(path, index=False)
    return path


class TestLevelReachesTheArchive:
    """Each entry point passes the level to the writer."""

    def test_df_to_xlsx(self, tmp_path: Path) -> None:
        """Level 1 packs the sheet larger than level 9, with the same content."""
        fast, small = tmp_path / "1.xlsx", tmp_path / "9.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), fast, compression_level=1)
        xlsxturbo.df_to_xlsx(_frame(), small, compression_level=9)
        assert _packed(fast) > _packed(small)
        assert _parts(fast) == _parts(small)

    def test_dfs_to_xlsx(self, tmp_path: Path) -> None:
        """The multi-sheet entry point takes it as a keyword for the whole workbook."""
        fast, small = tmp_path / "1.xlsx", tmp_path / "9.xlsx"
        xlsxturbo.dfs_to_xlsx([(_frame(), "Data")], fast, compression_level=1)
        xlsxturbo.dfs_to_xlsx([(_frame(), "Data")], small, compression_level=9)
        assert _packed(fast) > _packed(small)
        assert _parts(fast) == _parts(small)

    @pytest.mark.parametrize("parallel", [False, True])
    def test_csv_to_xlsx(self, csv_file: Path, tmp_path: Path, parallel: bool) -> None:
        """Both CSV pipelines take it."""
        fast, small = tmp_path / "1.xlsx", tmp_path / "9.xlsx"
        xlsxturbo.csv_to_xlsx(csv_file, fast, parallel=parallel, compression_level=1)
        xlsxturbo.csv_to_xlsx(csv_file, small, parallel=parallel, compression_level=9)
        assert _packed(fast) > _packed(small)
        assert _parts(fast) == _parts(small)

    def test_constant_memory(self, tmp_path: Path) -> None:
        """It applies to a constant_memory workbook too: it is set at the save."""
        fast, small = tmp_path / "1.xlsx", tmp_path / "9.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), fast, constant_memory=True, compression_level=1)
        xlsxturbo.df_to_xlsx(_frame(), small, constant_memory=True, compression_level=9)
        assert _packed(fast) > _packed(small)

    def test_omitted_is_level_6(self, tmp_path: Path) -> None:
        """Leaving it out gives the same archive as asking for 6, the writer's default."""
        omitted, six, one = tmp_path / "omitted.xlsx", tmp_path / "6.xlsx", tmp_path / "1.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), omitted)
        xlsxturbo.df_to_xlsx(_frame(), six, compression_level=6)
        xlsxturbo.df_to_xlsx(_frame(), one, compression_level=1)
        assert _packed(omitted) == _packed(six)
        # Control: the size does move with the level, so the equality above is not vacuous.
        assert _packed(one) != _packed(six)


class TestValidation:
    """The writer only prints a warning for a level out of range, so xlsxturbo checks it."""

    @pytest.mark.parametrize("level", [0, 10, -1, 256])
    def test_out_of_range(self, tmp_path: Path, level: int) -> None:
        """A level outside 1 to 9 is a ConfigurationError naming the value."""
        with pytest.raises(xlsxturbo.ConfigurationError, match=f"between 1 and 9, got {level}"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", compression_level=level)

    @pytest.mark.parametrize("level", [True, 1.0, "6"])
    def test_wrong_type(self, tmp_path: Path, level: Any) -> None:
        """A bool, a float or a string is a ConfigurationTypeError, not level 1."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="compression_level must be an integer"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", compression_level=level)

    def test_csv_to_xlsx_out_of_range(self, csv_file: Path, tmp_path: Path) -> None:
        """The CSV entry point validates it the same way."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="between 1 and 9, got 0"):
            xlsxturbo.csv_to_xlsx(csv_file, tmp_path / "out.xlsx", compression_level=0)

    def test_not_a_per_sheet_option(self, tmp_path: Path) -> None:
        """It describes the file, so a per-sheet options dict refuses it."""
        with pytest.raises(xlsxturbo.ConfigurationError, match="compression_level"):
            xlsxturbo.dfs_to_xlsx(
                [(_frame(), "Data", {"compression_level": 1})],  # type: ignore[list-item]
                tmp_path / "out.xlsx",
            )


class TestCommandLine:
    """``--compression-level`` in the Python command."""

    def test_flag_is_passed_through(self, csv_file: Path, tmp_path: Path) -> None:
        """``-c 1`` packs larger than ``-c 9``."""
        fast, small = tmp_path / "1.xlsx", tmp_path / "9.xlsx"
        assert cli.main([str(csv_file), str(fast), "-c", "1"]) == 0
        assert cli.main([str(csv_file), str(small), "--compression-level", "9"]) == 0
        assert _packed(fast) > _packed(small)

    @pytest.mark.parametrize("level", ["0", "10", "x", "-1"])
    def test_invalid_level_exits_2(self, csv_file: Path, tmp_path: Path, level: str) -> None:
        """An invalid level is a usage error, as in the Rust binary."""
        with pytest.raises(SystemExit) as exc:
            cli.main([str(csv_file), str(tmp_path / "out.xlsx"), "-c", level])
        assert exc.value.code == cli.EXIT_USAGE
