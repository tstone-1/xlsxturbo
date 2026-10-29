#!/usr/bin/env python3
"""Measure how a batch of ``df_to_xlsx`` and of ``csv_to_xlsx`` exports scales across threads.

The thread figures on ``docs/performance.md`` come from this script. One seeded frame
is written ``--jobs`` times to ``io.BytesIO`` from a ``ThreadPoolExecutor``, once per
thread count; the same frame, saved once as a CSV file, is converted ``--jobs`` times
the same way. ``df_to_xlsx`` holds the GIL while it reads the frame, ``csv_to_xlsx``
releases it for the whole conversion, so the two show how much that difference is
worth. The function and thread-count pairs run interleaved, rotating which goes first,
and the reported batch time is the median after a warmup round.

The warmup round also checks the outputs: every export of every thread count must
unpack to the same workbook as a serial export of the same function, and each serial
export is compared cell by cell with the frame, so a race that corrupts or drops data
cannot look fast.

Usage:
    python benchmarks/benchmark_threads.py              # 1, 2, 4, 8 threads; 32 jobs of 8,000 rows
    python benchmarks/benchmark_threads.py --markdown   # the table on docs/performance.md
    python benchmarks/benchmark_threads.py --threads 1 2 4 8 16 --jobs 64
"""

from __future__ import annotations

import argparse
import io
import os
import statistics
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from pathlib import Path
from typing import TYPE_CHECKING

from benchmark import archive_contents, check_output, generate_test_data, get_system_info, run_order

if TYPE_CHECKING:
    from collections.abc import Callable

    import pandas as pd

FUNCTIONS = ("df_to_xlsx", "csv_to_xlsx")


def export(frame: pd.DataFrame) -> io.BytesIO:
    """Write the frame to a new buffer.

    Args:
        frame: The frame, shared read-only by every thread.

    Returns:
        The buffer.
    """
    import xlsxturbo

    buffer = io.BytesIO()
    xlsxturbo.df_to_xlsx(frame, buffer)
    return buffer


def export_csv(csv_path: str) -> io.BytesIO:
    """Convert the CSV file to a new buffer.

    Args:
        csv_path: The CSV file, read by every thread.

    Returns:
        The buffer.
    """
    import xlsxturbo

    buffer = io.BytesIO()
    xlsxturbo.csv_to_xlsx(csv_path, buffer)
    return buffer


def run_batch(job: Callable[[], io.BytesIO], threads: int, jobs: int) -> tuple[float, list[io.BytesIO]]:
    """Time one batch of exports on a pool of the given size.

    The pool is started before the clock and shut down after it, so the time is the
    batch alone.

    Args:
        job: One export, called ``jobs`` times.
        threads: Pool size.
        jobs: Exports in the batch.

    Returns:
        The elapsed seconds and the buffers.
    """
    with ThreadPoolExecutor(max_workers=threads) as pool:
        start = time.perf_counter()
        futures = [pool.submit(job) for _ in range(jobs)]
        buffers = [future.result() for future in futures]
        elapsed = time.perf_counter() - start
    return elapsed, buffers


def check_buffers(
    reference: io.BytesIO, outputs: dict[int, list[io.BytesIO]], frame: pd.DataFrame, label: str = "export",
) -> None:
    """Refuse any threaded export that is not the serial export's workbook.

    Args:
        reference: A serial export of the frame.
        outputs: The warmup buffers per thread count.
        frame: The written frame.
        label: Names the exports in the error message.

    Raises:
        RuntimeError: When the reference differs from the frame or any export from the reference.
    """
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_threads_") as temp_dir:
        path = Path(temp_dir) / "reference.xlsx"
        path.write_bytes(reference.getvalue())
        check_output(str(path), frame)
    expected = archive_contents(reference)
    for threads, buffers in outputs.items():
        for index, buffer in enumerate(buffers):
            if archive_contents(buffer) != expected:
                raise RuntimeError(f"{label} {index} on {threads} threads differs from the serial export")


def measure(
    frame: pd.DataFrame, csv_path: str, thread_counts: list[int], jobs: int, runs: int, *, verbose: bool = True,
) -> dict[str, dict[int, list[float]]]:
    """Run the warmup round with its checks, then the timed rounds.

    Args:
        frame: The frame to write.
        csv_path: The frame saved as CSV, for ``csv_to_xlsx``.
        thread_counts: Pool sizes; must include 1, the baseline.
        jobs: Exports per batch.
        runs: Timed rounds.
        verbose: Print progress.

    Returns:
        The batch times per function name and thread count.
    """
    job_for: dict[str, Callable[[], io.BytesIO]] = {
        "df_to_xlsx": partial(export, frame),
        "csv_to_xlsx": partial(export_csv, csv_path),
    }
    if verbose:
        print("Warmup and output check...", flush=True)
    for name, job in job_for.items():
        outputs = {threads: run_batch(job, threads, jobs)[1] for threads in thread_counts}
        check_buffers(job(), outputs, frame, f"{name} export")
        del outputs

    pairs = [(name, threads) for name in FUNCTIONS for threads in thread_counts]
    times: dict[str, dict[int, list[float]]] = {name: {threads: [] for threads in thread_counts} for name in FUNCTIONS}
    for run_num in range(1, runs + 1):
        if verbose:
            print(f"Run {run_num}/{runs}...", flush=True)
        for name, threads in run_order(pairs, run_num):
            times[name][threads].append(run_batch(job_for[name], threads, jobs)[0])
    return times


def format_markdown(times: dict[str, dict[int, list[float]]], rows: int, cols: int, jobs: int, runs: int) -> str:
    """The table and its test-system line, as published on ``docs/performance.md``.

    Args:
        times: Batch times per function name and thread count.
        rows: Rows per frame.
        cols: Columns per frame.
        jobs: Exports per batch.
        runs: Timed rounds.

    Returns:
        Markdown.
    """
    info = get_system_info()
    lines = [
        "| Threads | `df_to_xlsx` batch | Speedup | `csv_to_xlsx` batch | Speedup |",
        "|--------:|-------------------:|--------:|--------------------:|--------:|",
    ]
    for threads in sorted(times["df_to_xlsx"]):
        cells = []
        for name in FUNCTIONS:
            median = statistics.median(times[name][threads])
            base = statistics.median(times[name][1])
            cells.append(f"{median:.2f} s | {base / median:.2f}x")
        lines.append(f"| {threads} | {' | '.join(cells)} |")
    lines.append("")
    lines.append(
        f"*Test system: {info['platform']} {info['platform_release']}, {info['cpu_count']} CPUs, "
        f"Python {info['python_version']}, xlsxturbo {info['xlsxturbo_version']}. {jobs} exports to "
        f"`io.BytesIO` per batch, of one seeded {rows:,}-row x {cols}-column mixed frame and of the same "
        f"frame saved as CSV; functions and thread counts interleaved, median of {runs} runs after a "
        "warmup. Every export unpacks to the same workbook as a serial one of the same function. "
        "Regenerate with `python benchmarks/benchmark_threads.py --markdown`.*"
    )
    return "\n".join(lines)


def format_console(times: dict[str, dict[int, list[float]]]) -> str:
    """A plain-text table for the terminal.

    Args:
        times: Batch times per function name and thread count.

    Returns:
        The table.
    """
    lines = [f"{'Function':<12} {'Threads':>7} {'Batch (s)':>10} {'Stdev':>7} {'Speedup':>8}"]
    for name in FUNCTIONS:
        base = statistics.median(times[name][1])
        for threads in sorted(times[name]):
            samples = times[name][threads]
            median = statistics.median(samples)
            stdev = statistics.stdev(samples) if len(samples) > 1 else 0.0
            lines.append(f"{name:<12} {threads:>7} {median:>10.2f} {stdev:>7.3f} {base / median:>7.2f}x")
    return "\n".join(lines)


def main() -> int:
    """Parse arguments, measure, and print the results.

    Returns:
        0 on success, 1 when an output check failed.
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=8_000, help="Rows per frame (default: 8000)")
    parser.add_argument("--cols", type=int, default=8, help="Columns per frame (default: 8)")
    parser.add_argument("--jobs", type=int, default=32, help="Exports per batch (default: 32)")
    parser.add_argument("--runs", type=int, default=7, help="Timed rounds (default: 7)")
    parser.add_argument(
        "--threads", type=int, nargs="+", default=[1, 2, 4, 8], help="Pool sizes (default: 1 2 4 8)",
    )
    parser.add_argument("--markdown", action="store_true", help="Print the table for docs/performance.md")
    args = parser.parse_args()
    if min(args.rows, args.cols, args.jobs, args.runs, *args.threads) <= 0:
        parser.error("all counts must be positive")
    thread_counts = sorted({1, *args.threads})

    frame = generate_test_data(args.rows, args.cols)
    if not args.markdown:
        print(f"{args.jobs} exports of {args.rows:,} x {args.cols} per batch, {os.cpu_count()} CPUs")
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_threads_") as temp_dir:
        csv_path = str(Path(temp_dir) / "frame.csv")
        frame.to_csv(csv_path, index=False)
        try:
            times = measure(frame, csv_path, thread_counts, args.jobs, args.runs, verbose=not args.markdown)
        except RuntimeError as error:
            print(f"[FAIL] {error}", file=sys.stderr)
            return 1
    print(
        format_markdown(times, args.rows, args.cols, args.jobs, args.runs) if args.markdown
        else format_console(times)
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
