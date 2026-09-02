"""Tests for what the library does while another thread is inside it.

Both classes exist because `df_to_xlsx` and `dfs_to_xlsx` release the GIL for the
archive write. `TestConcurrentWrites` covers the hazard that creates -- Rust code
running with the interpreter unlocked -- and `TestSaveReleasesTheGil` covers the
benefit, which is the only reason to accept the hazard.
"""

from __future__ import annotations

import errno
import os
import threading
import time
from collections.abc import Callable
from typing import TYPE_CHECKING

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

if TYPE_CHECKING:
    from pathlib import Path

pytestmark = pytest.mark.skipif(
    not HAS_OPENPYXL, reason="openpyxl required for content verification"
)

ROWS = 4000

# A stuck save must fail the run rather than hang it: without these, a deadlock in the
# detached region -- the exact regression these tests exist to catch -- would stall CI
# until the job timeout with no diagnosis attached.
BARRIER_TIMEOUT_S = 30.0
JOIN_TIMEOUT_S = 300.0

# `scripts/coverage_report.py` builds the extension with `-Cinstrument-coverage` and sets
# this for both of the pytest passes it drives. Under that build the timing comparison
# below inverts -- see `TestSaveReleasesTheGil` -- so it is skipped there.
#
# Keyed on the script's own marker rather than on `LLVM_PROFILE_FILE`: only the first of
# those two passes sets that one, so keying on it left the second pass -- which runs
# against the same instrumented extension, under coverage.py tracing on top -- still
# making a wall-clock assertion, and failing it intermittently.
INSTRUMENTED = "XLSXTURBO_COVERAGE" in os.environ


def _frame() -> pd.DataFrame:
    """Build a frame big enough that the archive write is a real share of the call."""
    return pd.DataFrame(
        {
            "i": range(ROWS),
            "f": [x * 1.5 for x in range(ROWS)],
            "s": [f"row-{x}" for x in range(ROWS)],
        }
    )


def _run_threads(n_threads: int, body: Callable[[int], None]) -> tuple[float, list[str]]:
    """Release `n_threads` at once through a barrier and time them to completion.

    Args:
        n_threads: How many threads to start.
        body: Called once per thread with the thread's index.

    Returns:
        The wall time in seconds, and a description of every failure. Failures are
        returned rather than raised so the caller reports all of them, not merely
        whichever thread happened to finish first. A thread still running after
        `JOIN_TIMEOUT_S` is itself reported as a failure.
    """
    barrier = threading.Barrier(n_threads)
    failures: list[str] = []
    lock = threading.Lock()

    def worker(index: int) -> None:
        barrier.wait(timeout=BARRIER_TIMEOUT_S)
        try:
            body(index)
        except BaseException as exc:
            with lock:
                failures.append(f"thread {index}: {type(exc).__name__}: {exc}")

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(n_threads)]
    start = time.perf_counter()
    for thread in threads:
        thread.start()
    for index, thread in enumerate(threads):
        thread.join(timeout=JOIN_TIMEOUT_S)
        if thread.is_alive():
            failures.append(f"thread {index}: still running after {JOIN_TIMEOUT_S}s")
    return time.perf_counter() - start, failures


class TestConcurrentWrites:
    """Concurrent calls must produce the same files serial ones would.

    The archive write runs without the GIL, so these exercise Rust code executing
    while other threads are live in the interpreter. Every output is read back: a
    thread that raised nothing but wrote a truncated file would otherwise pass.

    These do *not* cover the GIL release itself -- they stay green with it removed.
    They cover the hazard it introduces, which is the part that could corrupt data.
    """

    def test_eight_threads_share_one_frame(self, tmp_path: Path) -> None:
        """One DataFrame, eight concurrent readers, twenty-four verified workbooks."""
        frame = _frame()  # deliberately shared: every thread reads this one object

        def body(index: int) -> None:
            for k in range(3):
                xlsxturbo.df_to_xlsx(frame, str(tmp_path / f"t{index}-{k}.xlsx"), sheet_name="S")

        _elapsed, failures = _run_threads(8, body)
        assert not failures, failures

        written = sorted(p.name for p in tmp_path.iterdir())
        assert len(written) == 24, f"expected 24 files, got {len(written)}"
        for name in written:
            worksheet = active_ws(load_workbook(str(tmp_path / name)))
            assert worksheet.max_row == ROWS + 1, f"{name}: {worksheet.max_row} rows"
            assert worksheet.cell(ROWS + 1, 1).value == ROWS - 1, f"{name}: last row wrong"
            assert worksheet.cell(ROWS + 1, 3).value == f"row-{ROWS - 1}", f"{name}: text wrong"

    def test_multi_sheet_writes_concurrently(self, tmp_path: Path) -> None:
        """`dfs_to_xlsx` releases the GIL at the same point and needs the same cover."""
        frame = _frame()

        def body(index: int) -> None:
            xlsxturbo.dfs_to_xlsx(
                [(frame, "one"), (frame, "two")], str(tmp_path / f"multi-{index}.xlsx")
            )

        _elapsed, failures = _run_threads(6, body)
        assert not failures, failures

        for index in range(6):
            workbook = load_workbook(str(tmp_path / f"multi-{index}.xlsx"))
            assert workbook.sheetnames == ["one", "two"]
            assert workbook["two"].max_row == ROWS + 1

    def test_configuration_errors_raise_correctly_from_many_threads(self, tmp_path: Path) -> None:
        """Twelve concurrent readers of the cached exception classes get one class object.

        This does *not* race the `PyOnceLock` that holds those classes: it is filled by
        `errors::register` during module import, so by the time any test runs it has been
        initialised for a long while. What it covers is concurrent *reads* of the cache
        while raising, and that `except` still matches -- which is identity, not equality.
        """
        frame = _frame()
        caught: list[type[BaseException]] = []
        lock = threading.Lock()

        def body(index: int) -> None:
            try:
                xlsxturbo.df_to_xlsx(
                    frame,
                    str(tmp_path / f"never-{index}.xlsx"),
                    column_widths={0: "wide"},  # type: ignore[dict-item]  # invalid on purpose
                )
            except BaseException as exc:
                with lock:
                    caught.append(type(exc))

        _elapsed, failures = _run_threads(12, body)
        assert not failures, failures
        assert len(caught) == 12, f"only {len(caught)} of 12 threads raised"
        # One class object, not twelve equal-looking ones: `except` matches by identity.
        assert set(caught) == {xlsxturbo.ConfigurationTypeError}

    def test_file_errors_from_inside_the_detached_save(self, tmp_path: Path) -> None:
        """A failure raised from *within* the GIL-free region must still arrive intact.

        The configuration error above is rejected before the save is reached, so it
        never crosses the detach boundary. This one does: a missing parent directory
        is detected inside `save_workbook`, and the `PyErr` for it is built after the
        GIL comes back. Eight threads do it at once.
        """
        frame = _frame()
        caught: list[tuple[type[BaseException], int | None]] = []
        lock = threading.Lock()

        def body(index: int) -> None:
            try:
                xlsxturbo.df_to_xlsx(frame, str(tmp_path / "no_such_dir" / f"{index}.xlsx"))
            except xlsxturbo.FileError as exc:
                with lock:
                    caught.append((type(exc), exc.errno))

        _elapsed, failures = _run_threads(8, body)
        assert not failures, failures
        assert len(caught) == 8, f"only {len(caught)} of 8 threads raised FileError"
        assert {kind for kind, _ in caught} == {xlsxturbo.FileError}
        assert {number for _, number in caught} == {errno.ENOENT}


@pytest.mark.skipif(
    (os.cpu_count() or 1) < 4, reason="needs 4 cores to observe threads running in parallel"
)
@pytest.mark.skipif(
    INSTRUMENTED,
    reason="an -Cinstrument-coverage build inverts this comparison; see the class docstring",
)
class TestSaveReleasesTheGil:
    """Threaded exports must genuinely run in parallel, not merely interleave.

    This is a timing assertion, so it is built to be dull rather than precise. The
    gap it sits in is wide: with the GIL held across the save, four threads finish
    in 0.95-1.02x the single-threaded wall time; with it released they finish in
    about 0.43x. The threshold is 0.80x, and each leg is the best of three runs,
    because best-of is stable against a scheduler hiccup where a mean is not.

    It is skipped under `-Cinstrument-coverage`, and that is not the threshold being
    too tight. On an instrumented build four threads are *slower* than one: 1.693s
    against 1.347s, measured on a 32-core machine where the ordinary build gives
    0.43x. So the comparison there measures the instrumentation rather than the GIL,
    at any threshold. The uninstrumented CI legs on Linux, macOS and Windows all
    passed the same assertion on 4-core runners, which is what pins the cause to the
    build rather than to the core count.

    Both entry points are measured separately, and that is not redundancy: they
    detach at two different call sites. Removing only the one in `lib.rs` left the
    whole suite green while `df_to_xlsx` was still covered -- measured, which is how
    the second case got written.

    The two legs are **interleaved** -- serial, parallel, serial, parallel,
    serial, parallel -- rather than run as two blocks, and best-of is taken per
    leg. Best-of protects a leg against a hiccup *inside* it; it cannot protect
    against machine load that covers the whole serial block and lifts before the
    parallel one, which inflates `serial` and produces a pass with the feature
    absent -- the one outcome this test exists to refuse. That is not
    hypothetical: during a review the `df_to_xlsx` leg passed against a build
    with neither `py.detach`, while a `cargo test` ran in the background; in
    isolation the same leg failed 10 of 10 runs. Interleaving makes any such
    load fall on both legs.

    A passing run records its ratio with `record_property`, so the number
    reaches the junit output instead of vanishing -- an assertion that prints
    nothing when it passes leaves nothing to compare a later run against.
    Measured: `--junitxml` yields `<property name="ratio" value="0.386..."/>`
    per case. pytest warns that `record_property` is "incompatible with
    junit_family 'xunit2'" and writes the property anyway; setting
    `junit_family = "legacy"` would silence it, and is not worth a change to
    `pyproject.toml` while nothing consumes the file.

    To prove either case can fail, delete the matching `py.detach` wrapper
    around `save_workbook` in `src/convert.rs` or `src/lib.rs`, rebuild, and run
    this class: the affected case reports about 0.95x and goes red.
    """

    @staticmethod
    def _interleaved_best_of_three(
        total: int, export: Callable[[str], None]
    ) -> tuple[float, float]:
        """Time the serial and parallel legs alternately, best of three each.

        Args:
            total: Exports per timed run, spread over that run's threads.
            export: Called with a unique output path; performs one export.

        Returns:
            `(serial, parallel)`, the shortest wall time of each leg's three
            runs, in seconds.
        """
        best = {"s": float("inf"), "p": float("inf")}
        legs = (("s", 1, total), ("p", 4, total // 4))
        for run in range(3):
            for leg, n_threads, per_thread in legs:

                def body(
                    index: int, leg: str = leg, run: int = run, per_thread: int = per_thread
                ) -> None:
                    for k in range(per_thread):
                        export(f"{leg}{run}-{index}-{k}.xlsx")

                elapsed, failures = _run_threads(n_threads, body)
                assert not failures, failures
                best[leg] = min(best[leg], elapsed)
        return best["s"], best["p"]

    @pytest.mark.parametrize("entry_point", ["df_to_xlsx", "dfs_to_xlsx"])
    def test_four_threads_beat_one_by_a_wide_margin(
        self,
        tmp_path: Path,
        entry_point: str,
        record_property: Callable[[str, object], None],
    ) -> None:
        """Four threads must finish a batch in under 80% of one thread's time.

        Args:
            tmp_path: pytest's per-test temporary directory.
            entry_point: The exported function under test.
            record_property: pytest's junit-property recorder, used so a passing
                run leaves its measured ratio behind.
        """
        frame = _frame()

        if entry_point == "df_to_xlsx":
            total = 12

            def export(name: str) -> None:
                xlsxturbo.df_to_xlsx(frame, str(tmp_path / name), sheet_name="S")
        else:
            total = 8  # each call writes two sheets, so fewer calls for the same work

            def export(name: str) -> None:
                xlsxturbo.dfs_to_xlsx([(frame, "one"), (frame, "two")], str(tmp_path / name))

        # Warm up: the first call in a process pays for lazy initialisation that would
        # otherwise land entirely on whichever leg happens to run first.
        export("warmup.xlsx")

        serial, parallel = self._interleaved_best_of_three(total, export)
        record_property("ratio", parallel / serial)

        assert parallel < serial * 0.80, (
            f"{entry_point}: four threads took {parallel:.3f}s against {serial:.3f}s on "
            f"one thread ({serial / parallel:.2f}x). The archive write is not running "
            f"without the GIL."
        )
