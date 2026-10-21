"""A write the operating system refuses is also a ``PermissionError``.

Code moved over from ``pandas.DataFrame.to_excel`` often has ``except
PermissionError`` around the write, to ask the user to close a workbook that is
still open in Excel. ``FileError`` alone is an ``OSError`` and not a
``PermissionError``, so that handler stopped catching it. ``FilePermissionError``
is the subclass raised for ``EACCES`` and ``EPERM``, and it is both.

The refusal itself is produced per platform. On Windows a handle held on the
target is enough, and is the real case. On POSIX an open file can be replaced, so
the tests use a folder without write permission, which does not stop root.
"""

from __future__ import annotations

import errno
import os
import subprocess
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo


def _frame() -> pd.DataFrame:
    """One column, one row."""
    return pd.DataFrame({"a": [1]})


@contextmanager
def _refused_target(tmp_path: Path) -> Iterator[Path]:
    """Yield an output path that cannot be written while the context is open."""
    if sys.platform == "win32":
        target = tmp_path / "held.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), target)
        with target.open("rb"):
            yield target
        return
    if os.geteuid() == 0:
        pytest.skip("root is not refused by a read-only folder")
    folder = tmp_path / "read-only"
    folder.mkdir()
    folder.chmod(0o555)
    try:
        yield folder / "out.xlsx"
    finally:
        folder.chmod(0o755)


def _icacls(folder: Path, *args: str) -> None:
    """Change a folder's ACL with the icacls that ships with Windows."""
    subprocess.run(  # noqa: S603 - fixed argv, no shell
        ["icacls", str(folder), *args],  # noqa: S607 - icacls is in System32 on every Windows
        check=True,
        capture_output=True,
    )


@contextmanager
def _refused_folder(tmp_path: Path) -> Iterator[Path]:
    """Yield a folder in which no file can be created while the context is open."""
    folder = tmp_path / "refused"
    folder.mkdir()
    if sys.platform == "win32":
        # Deny ACEs bind administrators too; S-1-1-0 is Everyone in every locale.
        everyone = "*S-1-1-0"
        _icacls(folder, "/deny", f"{everyone}:(OI)(CI)(W)")
        try:
            yield folder
        finally:
            _icacls(folder, "/remove:d", everyone)
        return
    if os.geteuid() == 0:
        pytest.skip("root is not refused by a read-only folder")
    folder.chmod(0o555)
    try:
        yield folder
    finally:
        folder.chmod(0o755)


class TestRefusedTempFolder:
    """A constant_memory sheet's temporary file, refused by the OS, is a FilePermissionError.

    It was a plain FileError until rust_xlsxwriter 0.101.0 kept the ``io::Error``
    (upstream #213).
    """

    def test_the_class_and_errno(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Exactly FilePermissionError with errno EACCES, naming the folder."""
        with _refused_folder(tmp_path) as folder:
            for name in ("TMPDIR", "TEMP", "TMP"):
                monkeypatch.setenv(name, str(folder))
            with pytest.raises(PermissionError) as caught:
                xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", constant_memory=True)
        assert type(caught.value) is xlsxturbo.FilePermissionError
        assert caught.value.errno == errno.EACCES
        assert "temporary file" in str(caught.value)
        assert folder.name in str(caught.value)


class TestRefusedWrite:
    """What a refused write raises."""

    @pytest.mark.parametrize(
        "handler", [PermissionError, xlsxturbo.FilePermissionError, xlsxturbo.FileError, OSError, ValueError]
    )
    def test_every_handler_catches_it(self, tmp_path: Path, handler: type[Exception]) -> None:
        """The new builtin, and each class that caught this failure before."""
        with _refused_target(tmp_path) as target, pytest.raises(handler):
            xlsxturbo.df_to_xlsx(_frame(), target)

    def test_the_class_errno_and_message(self, tmp_path: Path) -> None:
        """Exactly FilePermissionError, with errno EACCES and the library's own message."""
        with _refused_target(tmp_path) as target, pytest.raises(xlsxturbo.XlsxTurboError) as caught:
            xlsxturbo.df_to_xlsx(_frame(), target)
        assert type(caught.value) is xlsxturbo.FilePermissionError
        assert caught.value.errno == errno.EACCES  # type: ignore[attr-defined]
        assert str(caught.value).startswith("Failed to save workbook to ")
        assert target.name in str(caught.value)

    def test_the_multi_sheet_entry_point(self, tmp_path: Path) -> None:
        """dfs_to_xlsx saves through the same path and raises the same class."""
        with _refused_target(tmp_path) as target, pytest.raises(PermissionError):
            xlsxturbo.dfs_to_xlsx([(_frame(), "A")], target)

    def test_the_retry_loop_from_to_excel(self, tmp_path: Path) -> None:
        """The handler the class exists for: retry once the target is free again."""
        attempts = 0
        with _refused_target(tmp_path) as target:
            try:
                xlsxturbo.df_to_xlsx(_frame(), target)
            except PermissionError:
                attempts += 1
        assert attempts == 1
        assert xlsxturbo.df_to_xlsx(_frame(), target) == (2, 1)


class TestOtherFileFailures:
    """Controls: a file failure that is not a refusal stays a plain FileError."""

    def test_a_missing_folder_is_not_a_permission_error(self, tmp_path: Path) -> None:
        """ENOENT raises FileError itself, which except PermissionError does not catch."""
        with pytest.raises(xlsxturbo.FileError) as caught:
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "absent" / "out.xlsx")
        assert type(caught.value) is xlsxturbo.FileError
        assert not isinstance(caught.value, PermissionError)
        assert caught.value.errno == errno.ENOENT  # type: ignore[attr-defined]

    def test_a_missing_csv_input_is_not_a_permission_error(self, tmp_path: Path) -> None:
        """The read side: a CSV that does not exist is a plain FileError too."""
        with pytest.raises(xlsxturbo.FileError) as caught:
            xlsxturbo.csv_to_xlsx(tmp_path / "absent.csv", tmp_path / "out.xlsx")
        assert type(caught.value) is xlsxturbo.FileError


@pytest.mark.skipif(sys.platform == "win32", reason="chmod does not make a file unreadable on Windows")
class TestRefusedRead:
    """The read side of the same rule."""

    def test_an_unreadable_csv_input(self, tmp_path: Path) -> None:
        """A CSV the process may not read raises what open() raises for it."""
        if os.geteuid() == 0:
            pytest.skip("root can read a file without read permission")
        source = tmp_path / "in.csv"
        source.write_text("a\n1\n")
        source.chmod(0o000)
        try:
            with pytest.raises(PermissionError) as caught:
                xlsxturbo.csv_to_xlsx(source, tmp_path / "out.xlsx")
        finally:
            source.chmod(0o644)
        assert type(caught.value) is xlsxturbo.FilePermissionError
