#!/usr/bin/env python3
"""Benchmark ``csv_to_xlsx(parallel=True)`` against the single-threaded conversion.

The table on ``docs/performance.md`` comes from this script. The two modes run in
pairs, alternating which goes first, and the reported time is the median after a
warmup pair. The warmup pair also checks the outputs: both modes must report the same
row and column counts and unpack to the same workbook, so a parallel path that drops
or reorders rows cannot look fast.

Usage:
    python benchmarks/benchmark_parallel.py                 # 500,000 rows x 50 columns, 3 runs
    python benchmarks/benchmark_parallel.py --markdown      # the table on docs/performance.md
    python benchmarks/benchmark_parallel.py --rows 1000000 --cols 20
"""

from __future__ import annotations

import argparse
import random
import statistics
import string
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from benchmark import archive_contents, get_system_info


def generate_test_csv(filepath: str, rows: int, cols: int, seed: int = 42, *, verbose: bool = True) -> str:
    """Generate a seeded test CSV with mixed data types.

    Columns cycle through integers, floats, booleans, dates, datetimes and strings.

    Args:
        filepath: Where to write the CSV.
        rows: Number of data rows.
        cols: Number of columns.
        seed: Random seed.
        verbose: Print progress.

    Returns:
        The path.
    """
    if verbose:
        print(f"Generating test CSV: {rows:,} rows x {cols} columns...")
    rng = random.Random(seed)

    start = time.perf_counter()
    with Path(filepath).open("w", encoding="utf-8") as f:
        f.write(",".join(f"col_{i}" for i in range(cols)) + "\n")

        base_date = date(2020, 1, 1)
        base_datetime = datetime(2020, 1, 1, 0, 0, 0)
        for _row in range(rows):
            values: list[str] = []
            for col in range(cols):
                col_type = col % 6
                if col_type == 0:
                    values.append(str(rng.randint(-10000, 10000)))
                elif col_type == 1:
                    values.append(f"{rng.uniform(-1000, 1000):.4f}")
                elif col_type == 2:
                    values.append(rng.choice(["true", "false", "TRUE", "FALSE"]))
                elif col_type == 3:
                    d = base_date + timedelta(days=rng.randint(0, 1000))
                    values.append(d.strftime("%Y-%m-%d"))
                elif col_type == 4:
                    dt = base_datetime + timedelta(
                        days=rng.randint(0, 1000), hours=rng.randint(0, 23), minutes=rng.randint(0, 59),
                    )
                    values.append(dt.strftime("%Y-%m-%dT%H:%M:%S"))
                else:
                    length = rng.randint(5, 20)
                    values.append("".join(rng.choices(string.ascii_letters, k=length)))
            f.write(",".join(values) + "\n")

    if verbose:
        elapsed = time.perf_counter() - start
        size_mb = Path(filepath).stat().st_size / (1024 * 1024)
        print(f"  Generated in {elapsed:.2f}s ({size_mb:.1f} MB)")
    return filepath


def check_modes(outputs: dict[bool, tuple[tuple[int, int], Path]]) -> None:
    """Refuse a parallel conversion that differs from the single-threaded one.

    Args:
        outputs: Per mode, the reported (rows, cols) and the written workbook.

    Raises:
        RuntimeError: When the counts or the unpacked workbooks differ.
    """
    (single_shape, single_path), (parallel_shape, parallel_path) = outputs[False], outputs[True]
    if single_shape != parallel_shape:
        raise RuntimeError(f"parallel reported {parallel_shape}, single-threaded {single_shape}")
    if archive_contents(single_path) != archive_contents(parallel_path):
        raise RuntimeError("parallel output unpacks to a different workbook than single-threaded")


def benchmark_conversion(
    csv_path: str,
    runs: int = 3,
    warmup: bool = True,
    *,
    verbose: bool = True,
) -> dict[bool, tuple[float, float]]:
    """Measure both modes in pairs, alternating which mode runs first.

    Warm up each mode once before collecting samples, and check the warmup outputs.
    Pairing prevents a sustained load or thermal change from favouring one entire
    measurement block. Without a warmup the first pair's outputs are checked instead.

    Args:
        csv_path: The CSV to convert.
        runs: Timed pairs.
        warmup: Run one untimed pair first.
        verbose: Print progress.

    Returns:
        Median and standard deviation per mode, keyed by ``parallel``.

    Raises:
        ValueError: When ``runs`` is below 1.
    """
    import xlsxturbo

    if runs < 1:
        raise ValueError("runs must be at least 1")
    times: dict[bool, list[float]] = {False: [], True: []}
    total_runs = runs + (1 if warmup else 0)
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_parallel_") as temp_dir:
        for run in range(total_runs):
            is_warmup = warmup and run == 0
            modes = (False, True) if run % 2 == 0 else (True, False)
            outputs: dict[bool, tuple[tuple[int, int], Path]] = {}
            for parallel in modes:
                mode = "parallel" if parallel else "single-threaded"
                xlsx_path = Path(temp_dir) / f"{mode}.xlsx"
                if is_warmup and verbose:
                    print(f"  warmup ({mode})...", flush=True)
                start = time.perf_counter()
                shape = xlsxturbo.csv_to_xlsx(csv_path, str(xlsx_path), parallel=parallel)
                elapsed = time.perf_counter() - start
                if not is_warmup:
                    times[parallel].append(elapsed)
                outputs[parallel] = (shape, xlsx_path)
                if run == 0 and verbose:
                    xlsx_size = xlsx_path.stat().st_size / (1024 * 1024)
                    print(f"  {mode}: {shape[0]:,} rows x {shape[1]} cols -> {xlsx_size:.1f} MB")
            if run == 0:
                check_modes(outputs)

    return {
        mode: (statistics.median(samples), statistics.stdev(samples) if len(samples) > 1 else 0.0)
        for mode, samples in times.items()
    }


def format_markdown(results: dict[bool, tuple[float, float]], rows: int, cols: int, runs: int) -> str:
    """The table and its test-system line, as published on ``docs/performance.md``.

    Args:
        results: Median and standard deviation per mode.
        rows: Data rows in the CSV.
        cols: Columns in the CSV.
        runs: Timed pairs.

    Returns:
        Markdown.
    """
    info = get_system_info()
    single, _ = results[False]
    lines = [
        "| Mode | Time (s) | Stdev | Speedup |",
        "|------|---------:|------:|--------:|",
    ]
    for parallel, label in ((False, "single-threaded"), (True, "`parallel=True`")):
        median, stdev = results[parallel]
        lines.append(f"| {label} | {median:.2f} | {stdev:.3f} | {single / median:.2f}x |")
    lines.append("")
    lines.append(
        f"*Test system: {info['platform']} {info['platform_release']}, {info['cpu_count']} CPUs, "
        f"Python {info['python_version']}, xlsxturbo {info['xlsxturbo_version']}. A seeded CSV of "
        f"{rows:,} rows x {cols} columns (integers, floats, booleans, dates, datetimes, strings); "
        f"modes paired and alternated, median of {runs} runs after a warmup. Both modes write the same "
        "workbook. Regenerate with `python benchmarks/benchmark_parallel.py --markdown`.*"
    )
    return "\n".join(lines)


def main() -> int:
    """Run the parallel-vs-single-threaded conversion benchmark and print results.

    Returns:
        0 on success, 1 when the two modes wrote different workbooks.
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=500000, help="Number of rows (default: 500000)")
    parser.add_argument("--cols", type=int, default=50, help="Number of columns (default: 50)")
    parser.add_argument("--runs", type=int, default=3, help="Number of benchmark runs (default: 3)")
    parser.add_argument("--markdown", action="store_true", help="Print the table for docs/performance.md")
    args = parser.parse_args()
    if min(args.rows, args.cols, args.runs) <= 0:
        parser.error("--rows, --cols and --runs must be positive")
    verbose = not args.markdown

    if verbose:
        info = get_system_info()
        print("=" * 60)
        print("xlsxturbo Parallel Processing Benchmark")
        print("=" * 60)
        print(f"xlsxturbo version: {info['xlsxturbo_version']}")
        print(f"CPU cores: {info['cpu_count']}")
        print()

    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
        csv_path = tmp.name
    try:
        generate_test_csv(csv_path, args.rows, args.cols, verbose=verbose)
        if verbose:
            print()
            print(f"Benchmarking paired modes ({args.runs} runs each + warmup)...")
        try:
            results = benchmark_conversion(csv_path, runs=args.runs, verbose=verbose)
        except RuntimeError as error:
            print(f"[FAIL] {error}", file=sys.stderr)
            return 1
    finally:
        Path(csv_path).unlink(missing_ok=True)

    if args.markdown:
        print(format_markdown(results, args.rows, args.cols, args.runs))
        return 0
    single_med, single_std = results[False]
    parallel_med, parallel_std = results[True]
    print()
    print("=" * 60)
    print("Results:")
    print("=" * 60)
    print(f"Single-threaded: {single_med:.2f}s (stdev {single_std:.2f}s)")
    print(f"Parallel:        {parallel_med:.2f}s (stdev {parallel_std:.2f}s)")
    print(f"Speedup:         {single_med / parallel_med:.2f}x")
    return 0


if __name__ == "__main__":
    sys.exit(main())
