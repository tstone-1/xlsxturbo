"""Writing a workbook into a binary file-like object instead of a path.

``output_path`` accepts ``io.BytesIO``, a file opened ``"wb"``, or anything with a
``write(bytes)`` method. The archive is serialised in memory during the detached
save and handed to the object afterwards, so conversion/save failures write nothing.
Delivery failures may leave partial bytes in a writer that cannot be rolled back.
"""

from __future__ import annotations

import errno
import io
import os
import sys
import zipfile
from functools import partial
from pathlib import Path
from typing import ClassVar

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, TIMESTAMPED_PART, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


def archive_parts(data: bytes) -> dict[str, bytes]:
    """Every member of an .xlsx archive except the one carrying a timestamp.

    Args:
        data: The archive bytes.

    Returns:
        Member name to member bytes.
    """
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != TIMESTAMPED_PART}


class NoneReturningWriter:
    """A response-body style writer whose ``write`` returns ``None``, as Django's does."""

    def __init__(self) -> None:
        """Start with nothing written."""
        self.chunks: list[bytes] = []

    def write(self, data: bytes) -> None:
        """Accept everything, report nothing.

        Args:
            data: The bytes to store.
        """
        self.chunks.append(bytes(data))


class TrickleWriter:
    """A raw-stream style writer that accepts at most ``limit`` bytes per call."""

    def __init__(self, limit: int) -> None:
        """Accept ``limit`` bytes per call.

        Args:
            limit: The most bytes one ``write`` call keeps.
        """
        self.limit = limit
        self.calls = 0
        self.data = bytearray()

    def write(self, data: bytes) -> int:
        """Keep a prefix of ``data`` and say how much.

        Args:
            data: The bytes offered.

        Returns:
            How many bytes were kept.
        """
        self.calls += 1
        kept = bytes(data[: self.limit])
        self.data += kept
        return len(kept)


class StuckWriter:
    """A writer that never accepts a byte."""

    def write(self, data: bytes) -> int:
        """Accept nothing.

        Args:
            data: The bytes offered.

        Returns:
            Always 0.
        """
        return 0


class FailingWriter:
    """A writer whose ``write`` raises, like a closed socket."""

    def write(self, data: bytes) -> int:
        """Raise.

        Args:
            data: The bytes offered.

        Raises:
            BrokenPipeError: Always.
        """
        raise BrokenPipeError("client went away")


class TestBytesIO:
    """Each entry point writes the same workbook into a buffer as into a file."""

    def test_df_to_xlsx_buffer_matches_file(self, tmp_xlsx: str) -> None:
        """The buffer holds the same archive the path export writes."""
        df = pd.DataFrame({"n": [1, 2, 3], "s": ["a", "b", "c"]})
        buffer = io.BytesIO()

        assert xlsxturbo.df_to_xlsx(df, buffer, autofit=True, table_style="Medium2") == (4, 2)
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, table_style="Medium2")

        assert archive_parts(buffer.getvalue()) == archive_parts(Path(tmp_xlsx).read_bytes())

    def test_dfs_to_xlsx_buffer_matches_file(self, tmp_xlsx: str) -> None:
        """The multi-sheet path writes into a buffer through the same finish step."""
        sheets = [
            (pd.DataFrame({"a": [1]}), "Pandas"),
            (pl.DataFrame({"b": ["x", "y"]}), "Polars"),
        ]
        buffer = io.BytesIO()

        assert xlsxturbo.dfs_to_xlsx(sheets, buffer) == [(2, 1), (3, 1)]
        xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)

        assert archive_parts(buffer.getvalue()) == archive_parts(Path(tmp_xlsx).read_bytes())

    @pytest.mark.parametrize("parallel", [False, True])
    def test_csv_to_xlsx_buffer_matches_file(self, tmp_path: Path, parallel: bool) -> None:
        """Both CSV pipelines write into a buffer."""
        source = tmp_path / "in.csv"
        source.write_text("id,day,flag\n1,2024-01-02,true\n2,2024-01-03,false\n")
        target = tmp_path / "out.xlsx"
        buffer = io.BytesIO()

        assert xlsxturbo.csv_to_xlsx(source, buffer, parallel=parallel) == (3, 3)
        xlsxturbo.csv_to_xlsx(source, target, parallel=parallel)

        assert archive_parts(buffer.getvalue()) == archive_parts(target.read_bytes())

    def test_constant_memory_into_buffer(self) -> None:
        """constant_memory stages rows in temporary files and still lands in the buffer."""
        df = pd.DataFrame({"n": range(1000)})
        buffer = io.BytesIO()

        xlsxturbo.df_to_xlsx(df, buffer, constant_memory=True)

        buffer.seek(0)
        ws = active_ws(load_workbook(buffer))
        assert ws.max_row == 1001
        assert ws["A1001"].value == 999

    def test_buffer_is_left_open_at_its_new_position(self) -> None:
        """The caller owns the buffer: it is appended to at its position and not closed."""
        buffer = io.BytesIO()
        buffer.write(b"PREFIX")

        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), buffer)

        assert not buffer.closed
        data = buffer.getvalue()
        assert data.startswith(b"PREFIX")
        assert buffer.tell() == len(data)
        assert zipfile.is_zipfile(io.BytesIO(data[len(b"PREFIX") :]))

    def test_file_opened_wb(self, tmp_xlsx: str) -> None:
        """A real binary file object is a writer too."""
        with Path(tmp_xlsx).open("wb") as handle:
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [7]}), handle)

        assert active_ws(load_workbook(tmp_xlsx))["A2"].value == 7


class TestWriterProtocol:
    """How the ``write`` return value is honoured."""

    def test_none_return_means_everything_was_accepted(self) -> None:
        """A writer that returns None is called once with the whole archive."""
        writer = NoneReturningWriter()

        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), writer)

        assert len(writer.chunks) == 1
        assert zipfile.is_zipfile(io.BytesIO(writer.chunks[0]))

    def test_short_writes_are_retried_with_the_rest(self) -> None:
        """A writer that keeps a prefix is called again until the archive is complete."""
        writer = TrickleWriter(limit=1000)
        reference = io.BytesIO()
        df = pd.DataFrame({"a": list(range(50))})

        xlsxturbo.df_to_xlsx(df, writer)
        xlsxturbo.df_to_xlsx(df, reference)

        assert writer.calls > 1
        assert archive_parts(bytes(writer.data)) == archive_parts(reference.getvalue())

    def test_writer_that_accepts_nothing_raises_file_error(self) -> None:
        """A write() returning 0 fails instead of looping forever."""
        with pytest.raises(xlsxturbo.FileError, match="accepted 0 of the remaining"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), StuckWriter())

    def test_raw_writer_none_is_not_response_writer_success(self) -> None:
        """RawIOBase's None means would-block on every platform."""
        with io.RawIOBase() as writer:
            # RawIOBase.write is normally unsupported; this stream instead
            # reports the documented nonblocking result without an OS pipe.
            writer.write = lambda _data: None  # type: ignore[method-assign]
            with pytest.raises(xlsxturbo.FileError, match="would block") as error:
                xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), writer)
        assert error.value.errno == errno.EAGAIN

    @pytest.mark.skipif(sys.platform == "win32", reason="nonblocking anonymous pipes differ on Windows")
    @pytest.mark.parametrize("entry", ["df", "dfs", "csv", "csv_parallel"])
    def test_nonblocking_raw_writer_would_block_is_not_success(self, tmp_path: Path, entry: str) -> None:
        """A real full nonblocking pipe refuses all four export paths without data loss being hidden."""
        read_fd, write_fd = os.pipe()
        with os.fdopen(read_fd, "rb", buffering=0) as reader, os.fdopen(write_fd, "wb", buffering=0) as writer:
            os.set_blocking(write_fd, False)
            os.set_blocking(read_fd, False)
            filled = 0
            while True:
                try:
                    filled += os.write(write_fd, b"x" * 4096)
                except BlockingIOError:
                    break
            assert writer.write(b"x") is None
            df = pd.DataFrame({"a": [1]})
            if entry == "df":
                export = partial(xlsxturbo.df_to_xlsx, df, writer)
            elif entry == "dfs":
                export = partial(xlsxturbo.dfs_to_xlsx, [(df, "Sheet1")], writer)
            else:
                source = tmp_path / "synthetic.csv"
                source.write_text("a\n1\n", encoding="utf-8")
                export = partial(xlsxturbo.csv_to_xlsx, source, writer, parallel=entry == "csv_parallel")
            with pytest.raises(xlsxturbo.FileError, match="would block") as error:
                export()
            assert error.value.errno == errno.EAGAIN
            drained = 0
            while chunk := reader.read(65536):
                drained += len(chunk)
            assert drained == filled

    def test_exception_from_write_propagates_unchanged(self) -> None:
        """The writer's own exception reaches the caller as itself."""
        with pytest.raises(BrokenPipeError, match="client went away"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), FailingWriter())


class TestFailedExportWritesNothing:
    """The archive is handed over only after the save succeeds."""

    BAD_CHART: ClassVar[dict[str, dict[str, str]]] = {
        "D2": {"type": "bar", "data_range": "NoSuchSheet!$A$2:$A$3"}
    }

    def test_save_time_failure_leaves_buffer_empty(self) -> None:
        """A failure during serialisation reports as before and writes nothing."""
        buffer = io.BytesIO()

        with pytest.raises(xlsxturbo.FileError, match="Failed to save workbook to the output buffer"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), buffer, charts=self.BAD_CHART)  # type: ignore[arg-type]

        assert buffer.getvalue() == b""

    def test_option_failure_leaves_buffer_empty(self) -> None:
        """A failure before the save writes nothing either."""
        buffer = io.BytesIO()

        with pytest.raises(xlsxturbo.ConfigurationError):
            xlsxturbo.dfs_to_xlsx([], buffer)

        assert buffer.getvalue() == b""


class TestRejectedTargets:
    """Objects that cannot receive a workbook are refused before any work is done."""

    def test_text_stream_is_refused_by_name(self) -> None:
        """A text stream gets a message about binary mode, not a TypeError from write()."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match=r"text stream .*StringIO.*'wb'"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), io.StringIO())  # type: ignore[arg-type]

    def test_text_file_is_refused(self, tmp_path: Path) -> None:
        """A file opened in text mode is a text stream too."""
        with (tmp_path / "out.xlsx").open("w") as handle, pytest.raises(TypeError, match="text stream"):
            xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), handle)  # type: ignore[arg-type]

    def test_object_without_write_lists_what_is_accepted(self) -> None:
        """The message names every accepted kind of target."""
        with pytest.raises(TypeError, match="binary file-like object with a write\\(\\) method, got int"):
            xlsxturbo.csv_to_xlsx("in.csv", 42)  # type: ignore[arg-type]

    def test_str_subclass_with_write_is_still_a_path(self, tmp_path: Path) -> None:
        """A path is tried first, so adding a write method does not change what a str means."""

        class WritablePath(str):
            """A str that also looks like a stream."""

            def write(self, data: bytes) -> int:
                """Must never be called.

                Args:
                    data: The bytes offered.

                Raises:
                    AssertionError: Always.
                """
                raise AssertionError("a str path was treated as a writer")

        target = tmp_path / "out.xlsx"
        xlsxturbo.df_to_xlsx(pd.DataFrame({"a": [1]}), WritablePath(target))

        assert target.exists()
