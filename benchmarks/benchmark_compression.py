#!/usr/bin/env python3
"""Measure what ``compression_level`` costs and saves, for ``df_to_xlsx`` and ``csv_to_xlsx``.

The table on ``docs/performance.md`` comes from this script. Every level writes the
same seeded frame to ``io.BytesIO``. The levels run interleaved, rotating which goes
first, and the reported time is the median after a warmup round.

The warmup round also checks the outputs: every level must unpack to the same workbook
as level 6, because a level changes only how the archive is packed, and the level-6
``df_to_xlsx`` output is compared cell by cell with the frame.

Usage:
    python benchmarks/benchmark_compression.py              # levels 1-9, 100,000 rows, 9 runs
    python benchmarks/benchmark_compression.py --markdown   # the table on docs/performance.md
    python benchmarks/benchmark_compression.py --rows 20000 --runs 3 --levels 1 6 9
"""

from __future__ import annotations

import argparse
import io
import statistics
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from benchmark import archive_contents, check_output, get_system_info, run_order

if TYPE_CHECKING:
    import pandas as pd

DEFAULT_LEVEL = 6
# The two functions that take compression_level and read their data differently.
FUNCTIONS = ("df_to_xlsx", "csv_to_xlsx")


@dataclass
class LevelResult:
    """Timings and output size for one compression level."""

    level: int
    times: dict[str, list[float]]
    size_bytes: int

    def median(self, function: str) -> float:
        """The median time of one function at this level, in seconds.

        Args:
            function: A name from ``FUNCTIONS``.

        Returns:
            The median of the timed runs.
        """
        return statistics.median(self.times[function])


def generate_frame(rows: int, seed: int = 42) -> pd.DataFrame:
    """A seeded frame shaped like a typical export, which is what decides compressibility.

    Seven columns: integers, two float columns, a category with 20 repeated values,
    a unique id per row, booleans and dates. Random letters, the strings of
    ``benchmark.py``, barely compress, so they would understate what a level saves.

    Args:
        rows: Number of rows.
        seed: Random seed.

    Returns:
        The frame.
    """
    import numpy as np
    import pandas as pd

    rng = np.random.default_rng(seed)
    categories = np.array([f"category_{i:02d}" for i in range(20)])
    return pd.DataFrame({
        "quantity": rng.integers(0, 10_000, rows),
        "price": (rng.random(rows) * 1000).round(2),
        "ratio": rng.random(rows),
        "category": categories[rng.integers(0, len(categories), rows)],
        "id": [f"ID-{i:08d}" for i in range(rows)],
        "active": rng.integers(0, 2, rows).astype(bool),
        "date": np.datetime64("2020-01-01", "ns") + rng.integers(0, 1000, rows).astype("timedelta64[D]"),
    })


def write(function: str, frame: pd.DataFrame, csv_path: str, level: int) -> tuple[float, io.BytesIO]:
    """Time one export at one level.

    Args:
        function: A name from ``FUNCTIONS``.
        frame: The frame ``df_to_xlsx`` writes.
        csv_path: The same data as CSV, for ``csv_to_xlsx``.
        level: The compression level.

    Returns:
        The elapsed seconds and the written buffer.
    """
    import xlsxturbo

    buffer = io.BytesIO()
    start = time.perf_counter()
    if function == "df_to_xlsx":
        xlsxturbo.df_to_xlsx(frame, buffer, compression_level=level)
    else:
        xlsxturbo.csv_to_xlsx(csv_path, buffer, compression_level=level)
    return time.perf_counter() - start, buffer


def check_levels(outputs: dict[tuple[str, int], io.BytesIO], frame: pd.DataFrame) -> None:
    """Refuse a level whose output is not the default level's workbook.

    Args:
        outputs: The warmup buffer per function and level.
        frame: The written frame.

    Raises:
        RuntimeError: When a level unpacks to different content, or the default level's
            ``df_to_xlsx`` output differs from the frame.
    """
    for function in FUNCTIONS:
        reference = archive_contents(outputs[function, DEFAULT_LEVEL])
        for (name, level), buffer in outputs.items():
            if name == function and archive_contents(buffer) != reference:
                raise RuntimeError(f"{function} at level {level} unpacks to a different workbook than level 6")
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_level_") as temp_dir:
        path = Path(temp_dir) / "level6.xlsx"
        path.write_bytes(outputs["df_to_xlsx", DEFAULT_LEVEL].getvalue())
        check_output(str(path), frame)


def measure(
    frame: pd.DataFrame, levels: list[int], runs: int, *, verbose: bool = True,
) -> dict[int, LevelResult]:
    """Run the warmup round with its checks, then the timed rounds.

    Args:
        frame: The frame to write.
        levels: The levels to measure; must include the default level.
        runs: Timed rounds.
        verbose: Print progress.

    Returns:
        The result per level.
    """
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_compression_") as temp_dir:
        csv_path = str(Path(temp_dir) / "data.csv")
        frame.to_csv(csv_path, index=False)
        jobs = [(function, level) for level in levels for function in FUNCTIONS]

        if verbose:
            print("Warmup and output check...", flush=True)
        outputs = {job: write(job[0], frame, csv_path, job[1])[1] for job in jobs}
        check_levels(outputs, frame)
        sizes = {level: len(outputs["df_to_xlsx", level].getvalue()) for level in levels}
        del outputs

        results = {level: LevelResult(level, {f: [] for f in FUNCTIONS}, sizes[level]) for level in levels}
        for run_num in range(1, runs + 1):
            if verbose:
                print(f"Run {run_num}/{runs}...", flush=True)
            for function, level in run_order(jobs, run_num):
                elapsed, _ = write(function, frame, csv_path, level)
                results[level].times[function].append(elapsed)
    return results


def _change(value: float, reference: float) -> str:
    """A relative change as a signed whole percentage, or ``+0.x%`` below one percent.

    Args:
        value: The measured value.
        reference: The default level's value.

    Returns:
        The change, e.g. ``-12%``.
    """
    change = (value / reference - 1) * 100
    return f"{change:+.1f}%" if abs(change) < 1 else f"{change:+.0f}%"


def format_markdown(results: dict[int, LevelResult], rows: int, runs: int) -> str:
    """The table and its test-system line, as published on ``docs/performance.md``.

    Args:
        results: The result per level.
        rows: Rows in the frame.
        runs: Timed rounds.

    Returns:
        Markdown.
    """
    info = get_system_info()
    base = results[DEFAULT_LEVEL]
    lines = [
        "| Level | `df_to_xlsx` | `csv_to_xlsx` | File size |",
        "|------:|-------------:|--------------:|----------:|",
    ]
    for level in sorted(results):
        result = results[level]
        cells = [f"{level} (default)" if level == DEFAULT_LEVEL else str(level)]
        for function in FUNCTIONS:
            ms = result.median(function) * 1000
            cells.append(
                f"{ms:.0f} ms" if level == DEFAULT_LEVEL
                else f"{ms:.0f} ms ({_change(result.median(function), base.median(function))})"
            )
        cells.append(
            f"{result.size_bytes / 2**20:.1f} MiB" if level == DEFAULT_LEVEL
            else _change(result.size_bytes, base.size_bytes)
        )
        lines.append("| " + " | ".join(cells) + " |")
    lines.append("")
    lines.append(
        f"*Test system: {info['platform']} {info['platform_release']}, {info['cpu_count']} CPUs, "
        f"Python {info['python_version']}, xlsxturbo {info['xlsxturbo_version']}. A seeded {rows:,}-row "
        "frame of seven columns (integers, two float columns, 20 repeated strings, unique ids, "
        f"booleans, dates), written to `io.BytesIO`; levels interleaved, median of {runs} runs after "
        "a warmup. File size is the `df_to_xlsx` output. Every level unpacks to the same workbook. "
        "Regenerate with `python benchmarks/benchmark_compression.py --markdown`.*"
    )
    return "\n".join(lines)


def format_console(results: dict[int, LevelResult]) -> str:
    """A plain-text table for the terminal.

    Args:
        results: The result per level.

    Returns:
        The table.
    """
    lines = [f"{'Level':>5} {'df_to_xlsx (ms)':>16} {'csv_to_xlsx (ms)':>17} {'Size (KiB)':>11}"]
    for level in sorted(results):
        result = results[level]
        lines.append(
            f"{level:>5} {result.median('df_to_xlsx') * 1000:>16.0f} "
            f"{result.median('csv_to_xlsx') * 1000:>17.0f} {result.size_bytes / 1024:>11,.0f}"
        )
    return "\n".join(lines)


def main() -> int:
    """Parse arguments, measure, and print the results.

    Returns:
        0 on success, 1 when an output check failed.
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=100_000, help="Rows in the frame (default: 100000)")
    parser.add_argument("--runs", type=int, default=9, help="Timed rounds (default: 9)")
    parser.add_argument(
        "--levels", type=int, nargs="+", default=list(range(1, 10)), help="Levels to measure (default: 1-9)",
    )
    parser.add_argument("--markdown", action="store_true", help="Print the table for docs/performance.md")
    args = parser.parse_args()
    if args.rows <= 0 or args.runs <= 0:
        parser.error("--rows and --runs must be positive")
    if any(not 1 <= level <= 9 for level in args.levels):
        parser.error("--levels must be between 1 and 9")
    levels = sorted({*args.levels, DEFAULT_LEVEL})

    frame = generate_frame(args.rows)
    try:
        results = measure(frame, levels, args.runs, verbose=not args.markdown)
    except RuntimeError as error:
        print(f"[FAIL] {error}", file=sys.stderr)
        return 1
    print(format_markdown(results, args.rows, args.runs) if args.markdown else format_console(results))
    return 0


if __name__ == "__main__":
    sys.exit(main())
