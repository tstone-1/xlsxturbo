#!/usr/bin/env python3
"""
xlsxturbo Benchmark Suite

Professional benchmark comparing Excel writing performance across libraries:
- xlsxturbo (Rust-based)
- pandas + openpyxl
- pandas + xlsxwriter
- polars.write_excel

Usage:
    python benchmarks/benchmark.py           # Quick benchmark (medium size only)
    python benchmarks/benchmark.py --full    # Full benchmark (all sizes)
    python benchmarks/benchmark.py --markdown # Output as markdown table
    python benchmarks/benchmark.py --json    # Output as JSON for CI
    python benchmarks/benchmark.py --rows 1000000 --cols 100  # Custom size

Examples:
    python benchmarks/benchmark.py --full --markdown > benchmark_results.md
    python benchmarks/benchmark.py --json > benchmark_results.json
"""

import argparse
import gc
import json
import os
import platform
import statistics
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import date, timedelta

# Number of data type categories for column generation:
# 0,1 = integers (25%), 2,3 = floats (25%), 4,5 = strings (25%), 6 = dates (12.5%), 7 = booleans (12.5%)
NUM_TYPE_CATEGORIES = 8


@dataclass
class BenchmarkResult:
    """Result from a single benchmark run."""
    library: str
    time_seconds: float
    rows_per_second: float
    file_size_mb: float
    success: bool
    error: str | None = None


@dataclass
class BenchmarkSummary:
    """Summary of multiple runs for a library."""
    library: str
    median_time: float
    rows_per_second: float
    file_size_mb: float
    speedup_vs_xlsxturbo: float
    all_times: list[float]


def get_system_info() -> dict:
    """Collect system information for reproducibility."""
    import xlsxturbo

    info = {
        "python_version": platform.python_version(),
        "platform": platform.system(),
        "platform_release": platform.release(),
        "processor": platform.processor() or "Unknown",
        "xlsxturbo_version": xlsxturbo.version(),
    }

    # Try to get CPU count
    info["cpu_count"] = os.cpu_count() or "Unknown"

    return info


def generate_test_data(rows: int, cols: int):
    """
    Generate test DataFrame with realistic mixed types:
    - 25% integers
    - 25% floats
    - 25% strings (5-20 chars)
    - 12.5% dates
    - 12.5% booleans
    """
    import numpy as np
    import pandas as pd

    data = {}
    base_date = date(2020, 1, 1)

    for i in range(cols):
        col_type = i % NUM_TYPE_CATEGORIES

        if col_type in (0, 1):
            # Integer column (25%)
            data[f"int_{i}"] = np.random.randint(0, 1_000_000, rows)
        elif col_type in (2, 3):
            # Float column (25%)
            data[f"float_{i}"] = np.random.random(rows) * 10000
        elif col_type in (4, 5):
            # String column (25%) - variable length 5-20 chars
            lengths = np.random.randint(5, 21, rows)
            data[f"str_{i}"] = [
                "".join(np.random.choice(list("abcdefghijklmnopqrstuvwxyz"), length))
                for length in lengths
            ]
        elif col_type == 6:
            # Date column (12.5%)
            days_offset = np.random.randint(0, 1000, rows)
            data[f"date_{i}"] = [base_date + timedelta(days=int(d)) for d in days_offset]
        else:
            # Boolean column (12.5%)
            data[f"bool_{i}"] = np.random.choice([True, False], rows)

    return pd.DataFrame(data)


def get_file_size_mb(filepath: str) -> float:
    """Get file size in megabytes."""
    return os.path.getsize(filepath) / (1024 * 1024)


def run_benchmark_xlsxturbo(df_pd, output_path: str, rows: int) -> BenchmarkResult:
    """Benchmark xlsxturbo df_to_xlsx."""
    import xlsxturbo

    try:
        start = time.perf_counter()
        xlsxturbo.df_to_xlsx(df_pd, output_path)
        elapsed = time.perf_counter() - start
        size_mb = get_file_size_mb(output_path)
        return BenchmarkResult(
            library="xlsxturbo",
            time_seconds=elapsed,
            rows_per_second=rows / elapsed,
            file_size_mb=size_mb,
            success=True,
        )
    except Exception as e:
        return BenchmarkResult(
            library="xlsxturbo",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error=str(e),
        )


def run_benchmark_pandas_openpyxl(df_pd, output_path: str, rows: int) -> BenchmarkResult:
    """Benchmark pandas with openpyxl engine."""
    try:
        start = time.perf_counter()
        df_pd.to_excel(output_path, index=False, engine="openpyxl")
        elapsed = time.perf_counter() - start
        size_mb = get_file_size_mb(output_path)
        return BenchmarkResult(
            library="pandas + openpyxl",
            time_seconds=elapsed,
            rows_per_second=rows / elapsed,
            file_size_mb=size_mb,
            success=True,
        )
    except ImportError:
        return BenchmarkResult(
            library="pandas + openpyxl",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error="openpyxl not installed",
        )
    except Exception as e:
        return BenchmarkResult(
            library="pandas + openpyxl",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error=str(e),
        )


def run_benchmark_pandas_xlsxwriter(df_pd, output_path: str, rows: int) -> BenchmarkResult:
    """Benchmark pandas with xlsxwriter engine."""
    try:
        start = time.perf_counter()
        df_pd.to_excel(output_path, index=False, engine="xlsxwriter")
        elapsed = time.perf_counter() - start
        size_mb = get_file_size_mb(output_path)
        return BenchmarkResult(
            library="pandas + xlsxwriter",
            time_seconds=elapsed,
            rows_per_second=rows / elapsed,
            file_size_mb=size_mb,
            success=True,
        )
    except ImportError:
        return BenchmarkResult(
            library="pandas + xlsxwriter",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error="xlsxwriter not installed",
        )
    except Exception as e:
        return BenchmarkResult(
            library="pandas + xlsxwriter",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error=str(e),
        )


def run_benchmark_polars(df_pd, output_path: str, rows: int, df_pl=None) -> BenchmarkResult:
    """Benchmark polars write_excel."""
    try:
        import polars as pl

        # Use pre-converted DataFrame if provided, otherwise convert (not timed)
        if df_pl is None:
            df_pl = pl.from_pandas(df_pd)
        start = time.perf_counter()
        df_pl.write_excel(output_path)
        elapsed = time.perf_counter() - start
        size_mb = get_file_size_mb(output_path)
        return BenchmarkResult(
            library="polars",
            time_seconds=elapsed,
            rows_per_second=rows / elapsed,
            file_size_mb=size_mb,
            success=True,
        )
    except ImportError:
        return BenchmarkResult(
            library="polars",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error="polars not installed",
        )
    except Exception as e:
        return BenchmarkResult(
            library="polars",
            time_seconds=0,
            rows_per_second=0,
            file_size_mb=0,
            success=False,
            error=str(e),
        )


BENCHMARK_FUNCS = [
    ("xlsxturbo", run_benchmark_xlsxturbo),
    ("pandas + openpyxl", run_benchmark_pandas_openpyxl),
    ("pandas + xlsxwriter", run_benchmark_pandas_xlsxwriter),
    ("polars", run_benchmark_polars),
]


def run_benchmarks(
    df_pd,
    rows: int,
    cols: int,
    runs: int = 3,
    warmup: bool = True,
    verbose: bool = True,
) -> dict[str, BenchmarkSummary]:
    """
    Run benchmarks for all libraries.

    Args:
        df_pd: pandas DataFrame to benchmark
        rows: Number of rows in the DataFrame
        cols: Number of columns
        runs: Number of benchmark runs per library
        warmup: Whether to do a warmup run (discarded)
        verbose: Whether to print progress

    Returns:
        Dictionary mapping library name to BenchmarkSummary
    """
    temp_dir = tempfile.mkdtemp(prefix="xlsxturbo_bench_")
    results: dict[str, list[BenchmarkResult]] = {name: [] for name, _ in BENCHMARK_FUNCS}

    # Pre-convert polars DataFrame once (outside timing)
    df_pl = None
    try:
        import polars as pl
        df_pl = pl.from_pandas(df_pd)
    except ImportError:
        pass  # polars not installed, will be skipped

    # Warmup run (discarded)
    if warmup and verbose:
        print("Warmup run...", flush=True)
        for name, func in BENCHMARK_FUNCS:
            output_path = os.path.join(temp_dir, f"warmup_{name.replace(' ', '_')}.xlsx")
            if name == "polars":
                func(df_pd, output_path, rows, df_pl=df_pl)
            else:
                func(df_pd, output_path, rows)
            gc.collect()
            if os.path.exists(output_path):
                os.unlink(output_path)

    # Main benchmark runs
    for run_num in range(1, runs + 1):
        if verbose:
            print(f"Run {run_num}/{runs}...", flush=True)

        for name, func in BENCHMARK_FUNCS:
            output_path = os.path.join(temp_dir, f"run{run_num}_{name.replace(' ', '_')}.xlsx")

            gc.collect()
            if name == "polars":
                result = func(df_pd, output_path, rows, df_pl=df_pl)
            else:
                result = func(df_pd, output_path, rows)
            results[name].append(result)

            if verbose and result.success:
                print(f"  {name}: {result.time_seconds:.2f}s", flush=True)
            elif verbose and not result.success:
                print(f"  {name}: SKIPPED ({result.error})", flush=True)

            # Clean up file
            if os.path.exists(output_path):
                os.unlink(output_path)

    # Clean up temp directory
    try:
        os.rmdir(temp_dir)
    except OSError:
        pass

    # Calculate summaries
    summaries = {}
    xlsxturbo_median = None

    for name, run_results in results.items():
        successful = [r for r in run_results if r.success]
        if not successful:
            continue

        times = [r.time_seconds for r in successful]
        median_time = statistics.median(times)
        median_rps = rows / median_time
        avg_size = statistics.mean([r.file_size_mb for r in successful])

        if name == "xlsxturbo":
            xlsxturbo_median = median_time

        summaries[name] = BenchmarkSummary(
            library=name,
            median_time=median_time,
            rows_per_second=median_rps,
            file_size_mb=avg_size,
            speedup_vs_xlsxturbo=1.0,  # Will update below
            all_times=times,
        )

    # Calculate speedup vs xlsxturbo
    if xlsxturbo_median:
        for name, summary in summaries.items():
            summary.speedup_vs_xlsxturbo = summary.median_time / xlsxturbo_median

    return summaries


def format_console_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict,
) -> str:
    """Format results for console output."""
    import xlsxturbo

    lines = []
    lines.append("")
    lines.append(f"xlsxturbo Benchmark Suite v{xlsxturbo.version()}")
    lines.append("=" * 75)
    lines.append(f"System: {system_info['platform']} {system_info['platform_release']}, "
                 f"Python {system_info['python_version']}, {system_info['cpu_count']} CPUs")
    lines.append("")
    lines.append(f"Dataset: {rows:,} rows x {cols} columns (mixed types)")
    lines.append(f"Runs: {runs} (median reported)")
    lines.append("")

    # Sort by time (fastest first)
    sorted_summaries = sorted(summaries.values(), key=lambda s: s.median_time)

    # Header
    lines.append(f"{'Library':<22} {'Time (s)':>10} {'Rows/sec':>12} {'Size (MB)':>10} {'vs xlsxturbo':>13}")
    lines.append("-" * 75)

    for summary in sorted_summaries:
        speedup_str = f"{summary.speedup_vs_xlsxturbo:.1f}x"
        if summary.library == "xlsxturbo":
            speedup_str = "1.0x (base)"

        lines.append(
            f"{summary.library:<22} "
            f"{summary.median_time:>10.2f} "
            f"{summary.rows_per_second:>12,.0f} "
            f"{summary.file_size_mb:>10.1f} "
            f"{speedup_str:>13}"
        )

    lines.append("")

    # Summary
    if len(sorted_summaries) >= 2:
        fastest = sorted_summaries[0]
        slowest = sorted_summaries[-1]
        max_speedup = slowest.median_time / fastest.median_time
        lines.append(f"Fastest: {fastest.library}")
        lines.append(f"Slowest: {slowest.library}")
        lines.append(f"Max speedup: {max_speedup:.1f}x")

    return "\n".join(lines)


def format_markdown_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict,
) -> str:
    """Format results as markdown table."""
    import xlsxturbo

    lines = []
    lines.append(f"## xlsxturbo Benchmark Results")
    lines.append("")
    lines.append(f"**System:** {system_info['platform']} {system_info['platform_release']}, "
                 f"Python {system_info['python_version']}, {system_info['cpu_count']} CPUs")
    lines.append(f"**xlsxturbo version:** {xlsxturbo.version()}")
    lines.append(f"**Dataset:** {rows:,} rows x {cols} columns (mixed types)")
    lines.append(f"**Runs:** {runs} (median reported)")
    lines.append("")

    # Sort by time (fastest first)
    sorted_summaries = sorted(summaries.values(), key=lambda s: s.median_time)

    lines.append("| Library | Time (s) | Rows/sec | Size (MB) | vs xlsxturbo |")
    lines.append("|---------|----------|----------|-----------|--------------|")

    for summary in sorted_summaries:
        speedup_str = f"{summary.speedup_vs_xlsxturbo:.1f}x"
        if summary.library == "xlsxturbo":
            speedup_str = "**1.0x**"

        name = summary.library
        if summary.library == "xlsxturbo":
            name = "**xlsxturbo**"

        lines.append(
            f"| {name} | {summary.median_time:.2f} | {summary.rows_per_second:,.0f} | "
            f"{summary.file_size_mb:.1f} | {speedup_str} |"
        )

    return "\n".join(lines)


def format_json_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict,
) -> str:
    """Format results as JSON for CI integration."""
    result = {
        "system": system_info,
        "benchmark": {
            "rows": rows,
            "cols": cols,
            "runs": runs,
            "data_types": {
                "integers": "25%",
                "floats": "25%",
                "strings": "25%",
                "dates": "12.5%",
                "booleans": "12.5%",
            },
        },
        "results": [
            {
                "library": s.library,
                "median_time_seconds": round(s.median_time, 3),
                "rows_per_second": round(s.rows_per_second, 0),
                "file_size_mb": round(s.file_size_mb, 2),
                "speedup_vs_xlsxturbo": round(s.speedup_vs_xlsxturbo, 2),
                "all_times": [round(t, 3) for t in s.all_times],
            }
            for s in sorted(summaries.values(), key=lambda s: s.median_time)
        ],
    }
    return json.dumps(result, indent=2)


# Predefined benchmark sizes
BENCHMARK_SIZES = {
    "small": (10_000, 20),
    "medium": (100_000, 50),
    "large": (500_000, 50),
}


def main():
    import xlsxturbo

    parser = argparse.ArgumentParser(
        description="xlsxturbo Benchmark Suite",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"xlsxturbo {xlsxturbo.version()}",
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help="Run full benchmark (small, medium, large sizes)",
    )
    parser.add_argument(
        "--rows",
        type=int,
        help="Custom number of rows (overrides --full)",
    )
    parser.add_argument(
        "--cols",
        type=int,
        help="Custom number of columns (overrides --full)",
    )
    parser.add_argument(
        "--runs",
        type=int,
        default=3,
        help="Number of benchmark runs per library (default: 3)",
    )
    parser.add_argument(
        "--markdown",
        action="store_true",
        help="Output as markdown table",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output as JSON for CI integration",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress progress output",
    )

    args = parser.parse_args()

    # Validate arguments
    if args.rows is not None and args.rows <= 0:
        parser.error("--rows must be a positive integer")
    if args.cols is not None and args.cols <= 0:
        parser.error("--cols must be a positive integer")
    if args.runs <= 0:
        parser.error("--runs must be a positive integer")

    verbose = not args.quiet and not args.json

    # Determine which sizes to benchmark
    if args.rows or args.cols:
        # Custom size
        sizes = [("custom", (args.rows or 100_000, args.cols or 50))]
    elif args.full:
        sizes = list(BENCHMARK_SIZES.items())
    else:
        # Default: medium only
        sizes = [("medium", BENCHMARK_SIZES["medium"])]

    # Collect system info
    system_info = get_system_info()

    all_outputs = []

    for size_name, (rows, cols) in sizes:
        if verbose:
            print(f"\n{'=' * 75}")
            print(f"Benchmark: {size_name} ({rows:,} rows x {cols} columns)")
            print(f"{'=' * 75}")
            print(f"Generating test data...", flush=True)

        # Generate test data
        df_pd = generate_test_data(rows, cols)

        if verbose:
            print(f"Data ready: {len(df_pd):,} rows x {len(df_pd.columns)} columns")
            print()

        # Run benchmarks
        summaries = run_benchmarks(
            df_pd,
            rows,
            cols,
            runs=args.runs,
            warmup=True,
            verbose=verbose,
        )

        if not summaries:
            print("No benchmarks completed successfully!", file=sys.stderr)
            continue

        # Format output
        if args.json:
            output = format_json_output(summaries, rows, cols, args.runs, system_info)
        elif args.markdown:
            output = format_markdown_output(summaries, rows, cols, args.runs, system_info)
        else:
            output = format_console_output(summaries, rows, cols, args.runs, system_info)

        all_outputs.append(output)

        # Clear DataFrame to free memory
        del df_pd
        gc.collect()

    # Print all outputs
    if args.json and len(all_outputs) > 1:
        # For JSON with multiple sizes, wrap in array
        print("[" + ",\n".join(all_outputs) + "]")
    else:
        print("\n\n".join(all_outputs))

    return 0


if __name__ == "__main__":
    sys.exit(main())
