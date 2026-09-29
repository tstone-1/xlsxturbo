#!/usr/bin/env python3
"""xlsxturbo Benchmark Suite.

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
    python benchmarks/benchmark.py --shape strings   # numeric, strings or mixed data
    python benchmarks/benchmark.py --memory          # also measure peak memory

Every library writes the same DataFrame. Each output is compared with that frame
after timing, so a library that silently writes less cannot look fast. Failed runs
remain in the report and make the command exit unsuccessfully.

Examples:
    python benchmarks/benchmark.py --full --markdown > benchmark_results.md
    python benchmarks/benchmark.py --json > benchmark_results.json
"""

from __future__ import annotations

import argparse
import contextlib
import gc
import importlib.metadata
import json
import math
import os
import platform
import re
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from itertools import chain
from pathlib import Path
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:
    import pandas as pd

# Column type cycles for each data shape. "mixed" is the reference workload:
# 25% integers, 25% floats, 25% strings (5-20 chars), 12.5% dates, 12.5% booleans.
# "numeric" and "strings" bracket it, because the gap between libraries depends on
# the data: string cells go through a shared-string table in every writer, numbers
# do not.
SHAPES: dict[str, tuple[str, ...]] = {
    "mixed": ("int", "int", "float", "float", "str", "str", "date", "bool"),
    "numeric": ("int", "float"),
    "strings": ("str",),
}

# Distributions whose versions decide the numbers, recorded with every result.
VERSIONED_PACKAGES = ("xlsxturbo", "pandas", "polars", "numpy", "openpyxl", "xlsxwriter")


@dataclass
class BenchmarkResult:
    """Result from a single benchmark run."""
    library: str
    time_seconds: float
    rows_per_second: float
    file_size_mb: float
    success: bool
    error: str | None = None
    peak_memory_mb: float | None = None


@dataclass
class BenchmarkSummary:
    """Summary of multiple runs for a library."""
    library: str
    median_time: float | None
    stdev_time: float | None
    rows_per_second: float | None
    file_size_mb: float | None
    speedup_vs_xlsxturbo: float | None
    all_times: list[float]
    attempted_runs: int
    errors: list[str]
    peak_memory_mb: float | None = None


def get_system_info() -> dict[str, object]:
    """Collect system information for reproducibility."""
    import xlsxturbo

    info: dict[str, object] = {
        "python_version": platform.python_version(),
        "platform": platform.system(),
        "platform_release": platform.release(),
        "processor": platform.processor() or "Unknown",
        "xlsxturbo_version": xlsxturbo.version(),
    }

    # Try to get CPU count
    info["cpu_count"] = os.cpu_count() or "Unknown"
    info["machine"] = platform.machine()
    info["packages"] = package_versions()

    return info


def package_versions() -> dict[str, str]:
    """The installed version of every package that affects the numbers.

    Returns:
        Distribution name to version, or "not installed".
    """
    versions: dict[str, str] = {}
    for name in VERSIONED_PACKAGES:
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = "not installed"
    return versions


def generate_test_data(rows: int, cols: int, seed: int = 42, shape: str = "mixed") -> pd.DataFrame:
    """Generate a seeded test DataFrame.

    Columns cycle through ``SHAPES[shape]``. Dates are ``datetime64[ns]``, the fast
    path for pandas; strings are 5-20 random lowercase letters.

    Args:
        rows: Number of rows.
        cols: Number of columns.
        seed: Random seed, so every library and every run writes the same data.
        shape: A key of ``SHAPES``.

    Returns:
        The DataFrame.
    """
    import numpy as np
    import pandas as pd

    rng = np.random.default_rng(seed)
    data: dict[str, object] = {}
    base_date = np.datetime64("2020-01-01", "ns")

    cycle = SHAPES[shape]
    for i in range(cols):
        col_type = cycle[i % len(cycle)]

        if col_type == "int":
            data[f"int_{i}"] = rng.integers(0, 1_000_000, rows)
        elif col_type == "float":
            data[f"float_{i}"] = rng.random(rows) * 10000
        elif col_type == "str":
            lengths = rng.integers(5, 21, rows)
            alphabet = np.array(list("abcdefghijklmnopqrstuvwxyz"))
            data[f"str_{i}"] = [
                "".join(rng.choice(alphabet, length))
                for length in lengths
            ]
        elif col_type == "date":
            days_offset = rng.integers(0, 1000, rows).astype("timedelta64[D]")
            data[f"date_{i}"] = base_date + days_offset.astype("timedelta64[ns]")
        else:
            data[f"bool_{i}"] = rng.integers(0, 2, rows).astype(bool)

    return pd.DataFrame(data)


def get_file_size_mb(filepath: str) -> float:
    """Get file size in megabytes."""
    return Path(filepath).stat().st_size / (1024 * 1024)


def column_letters(index: int) -> str:
    """Excel column letters for a 1-based column number.

    Args:
        index: 1 for A, 27 for AA.

    Returns:
        The column letters.
    """
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def written_dimension(filepath: str) -> str | None:
    """The used range the first worksheet declares, e.g. ``A1:AX100001``.

    Read the ``<dimension>`` element for a quick range check before comparing
    the individual cells.

    Args:
        filepath: The workbook.

    Returns:
        The range, or None when the sheet declares none.
    """
    with zipfile.ZipFile(filepath) as archive:
        head = archive.open("xl/worksheets/sheet1.xml").read(4096).decode("utf-8", "replace")
    match = re.search(r'<dimension ref="([^"]+)"', head)
    return match.group(1) if match else None


def check_output(filepath: str, frame: pd.DataFrame) -> None:
    """Compare every generated cell with the input, outside the timed write.

    Args:
        filepath: The workbook.
        frame: The generated benchmark frame, including its column names.

    Raises:
        RuntimeError: When dimensions, populated cells or values differ.
    """
    from openpyxl import load_workbook

    rows, cols = frame.shape
    expected = f"A1:{column_letters(cols)}{rows + 1}"
    found = written_dimension(filepath)
    if found != expected:
        raise RuntimeError(f"output declares {found!r}, expected {expected!r}")
    workbook = load_workbook(filepath, read_only=True, data_only=False)
    try:
        worksheet = workbook.worksheets[0]
        expected_rows = chain([tuple(frame.columns)], frame.itertuples(index=False, name=None))
        for row, (actual, wanted) in enumerate(zip(worksheet.values, expected_rows, strict=True), start=1):
            for col, (value, reference) in enumerate(zip(actual, wanted, strict=True), start=1):
                if isinstance(reference, bool):
                    equal = isinstance(value, bool) and value == reference
                elif isinstance(reference, float) and isinstance(value, (int, float)):
                    # Excel writers differ in the last decimal digit they serialize.
                    equal = not isinstance(value, bool) and math.isclose(value, reference, rel_tol=1e-14)
                else:
                    equal = not isinstance(value, bool) and value == reference
                if not equal:
                    raise RuntimeError(f"output differs from the input at {column_letters(col)}{row}")
    finally:
        workbook.close()


def run_benchmark_xlsxturbo(
    df_pd: pd.DataFrame, output_path: str, rows: int, *, verify_output: bool = True,
) -> BenchmarkResult:
    """Benchmark xlsxturbo df_to_xlsx."""
    import xlsxturbo

    try:
        start = time.perf_counter()
        xlsxturbo.df_to_xlsx(df_pd, output_path)
        elapsed = time.perf_counter() - start
        if verify_output:
            check_output(output_path, df_pd)
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


def run_benchmark_pandas_openpyxl(
    df_pd: pd.DataFrame, output_path: str, rows: int, *, verify_output: bool = True,
) -> BenchmarkResult:
    """Benchmark pandas with openpyxl engine."""
    try:
        start = time.perf_counter()
        df_pd.to_excel(output_path, index=False, engine="openpyxl")
        elapsed = time.perf_counter() - start
        if verify_output:
            check_output(output_path, df_pd)
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


def run_benchmark_pandas_xlsxwriter(
    df_pd: pd.DataFrame, output_path: str, rows: int, *, verify_output: bool = True,
) -> BenchmarkResult:
    """Benchmark pandas with xlsxwriter engine."""
    try:
        start = time.perf_counter()
        df_pd.to_excel(output_path, index=False, engine="xlsxwriter")
        elapsed = time.perf_counter() - start
        if verify_output:
            check_output(output_path, df_pd)
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


def run_benchmark_polars(
    df_pd: pd.DataFrame,
    output_path: str,
    rows: int,
    df_pl: object | None = None,
    *,
    verify_output: bool = True,
) -> BenchmarkResult:
    """Benchmark polars write_excel."""
    try:
        import polars as pl

        # Use pre-converted DataFrame if provided, otherwise convert (not timed)
        raw_frame = pl.from_pandas(df_pd) if df_pl is None else df_pl
        frame = cast("pl.DataFrame", raw_frame)
        start = time.perf_counter()
        frame.write_excel(output_path)
        elapsed = time.perf_counter() - start
        if verify_output:
            check_output(output_path, df_pd)
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


BENCHMARK_FUNCS: list[tuple[str, Callable[..., BenchmarkResult]]] = [
    ("xlsxturbo", run_benchmark_xlsxturbo),
    ("pandas + openpyxl", run_benchmark_pandas_openpyxl),
    ("pandas + xlsxwriter", run_benchmark_pandas_xlsxwriter),
    ("polars", run_benchmark_polars),
]


def _max_rss_mb() -> float:
    """This process's peak resident memory so far, in MB.

    Returns:
        The high-water mark. ``ru_maxrss`` is bytes on macOS and kilobytes on Linux.
    """
    import resource

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak / (1024 * 1024) if sys.platform == "darwin" else peak / 1024


def run_memory_child(library: str, rows: int, cols: int, shape: str) -> None:
    """Measure one library's write in this (fresh) process and print the result as JSON.

    The frame is built first, and the peak resident memory at that point is the
    baseline. The number reported is how far one write raises the peak above it:
    the memory the export needs on top of holding the data. A write whose own peak
    stays below the baseline reports about 0, so this is a floor, not an exact figure.

    Args:
        library: A name from ``BENCHMARK_FUNCS``.
        rows: Number of rows.
        cols: Number of columns.
        shape: A key of ``SHAPES``.
    """
    func = dict(BENCHMARK_FUNCS)[library]
    df_pd = generate_test_data(rows, cols, shape=shape)
    df_pl: object | None = None
    if library == "polars":
        import polars as pl

        df_pl = pl.from_pandas(df_pd)
    gc.collect()
    baseline = _max_rss_mb()
    with tempfile.TemporaryDirectory(prefix="xlsxturbo_mem_") as temp_dir:
        output_path = str(Path(temp_dir) / "out.xlsx")
        kwargs = {"df_pl": df_pl} if library == "polars" else {}
        # Readback allocates its own buffers and must not inflate export memory.
        result = func(df_pd, output_path, rows, verify_output=False, **kwargs)
    print(json.dumps({"ok": result.success, "error": result.error, "increase_mb": _max_rss_mb() - baseline}))


def measure_peak_memory(library: str, rows: int, cols: int, shape: str) -> float | None:
    """Run one write in a fresh interpreter and return its peak memory increase in MB.

    A fresh process per library, because a peak only ever rises: measured in one
    process, every library after the hungriest one would report nothing.
    Python-level tracing (``tracemalloc``) cannot be used either, because it does
    not see the Rust allocator.

    Args:
        library: A name from ``BENCHMARK_FUNCS``.
        rows: Number of rows.
        cols: Number of columns.
        shape: A key of ``SHAPES``.

    Returns:
        The increase in MB, or None when the child failed.
    """
    completed = subprocess.run(
        [
            sys.executable,
            str(Path(__file__).resolve()),
            "--memory-child",
            library,
            "--rows",
            str(rows),
            "--cols",
            str(cols),
            "--shape",
            shape,
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode:
        print(f"  {library}: memory run failed: {completed.stderr.strip()[-300:]}", file=sys.stderr)
        return None
    try:
        report = json.loads(completed.stdout.strip().splitlines()[-1])
    except (IndexError, json.JSONDecodeError):
        print(f"  {library}: memory run failed: {completed.stderr.strip()[-300:]}", file=sys.stderr)
        return None
    if not report["ok"]:
        print(f"  {library}: memory run failed: {report['error']}", file=sys.stderr)
        return None
    return max(float(report["increase_mb"]), 0.0)


def run_benchmarks(
    df_pd: pd.DataFrame,
    rows: int,
    cols: int,
    runs: int = 3,
    warmup: bool = True,
    verbose: bool = True,
) -> dict[str, BenchmarkSummary]:
    """Run benchmarks for all libraries.

    Args:
        df_pd: pandas DataFrame to benchmark
        rows: Number of rows in the DataFrame
        cols: Number of columns
        runs: Number of benchmark runs per library
        warmup: Whether to do a warmup run (timing discarded, failures retained)
        verbose: Whether to print progress

    Returns:
        Dictionary mapping library name to BenchmarkSummary
    """
    temp_dir = Path(tempfile.mkdtemp(prefix="xlsxturbo_bench_"))
    results: dict[str, list[BenchmarkResult]] = {name: [] for name, _ in BENCHMARK_FUNCS}
    failures: dict[str, list[str]] = {name: [] for name, _ in BENCHMARK_FUNCS}

    # Pre-convert polars DataFrame once (outside timing)
    df_pl: object | None = None
    try:
        import polars as pl
        df_pl = pl.from_pandas(df_pd)
    except ImportError:
        pass  # polars not installed, will be skipped

    # Warmup run (discarded)
    if warmup:
        if verbose:
            print("Warmup run...", flush=True)
        for name, func in BENCHMARK_FUNCS:
            output_path = temp_dir / f"warmup_{name.replace(' ', '_')}.xlsx"
            if name == "polars":
                result = func(df_pd, str(output_path), rows, df_pl=df_pl)
            else:
                result = func(df_pd, str(output_path), rows)
            if not result.success:
                failures[name].append(f"warmup: {result.error or 'failed'}")
            gc.collect()
            output_path.unlink(missing_ok=True)

    # Main benchmark runs
    for run_num in range(1, runs + 1):
        if verbose:
            print(f"Run {run_num}/{runs}...", flush=True)

        for name, func in BENCHMARK_FUNCS:
            output_path = temp_dir / f"run{run_num}_{name.replace(' ', '_')}.xlsx"

            gc.collect()
            if name == "polars":
                result = func(df_pd, str(output_path), rows, df_pl=df_pl)
            else:
                result = func(df_pd, str(output_path), rows)
            results[name].append(result)
            if not result.success:
                failures[name].append(f"run {run_num}: {result.error or 'failed'}")

            if verbose and result.success:
                print(f"  {name}: {result.time_seconds:.2f}s", flush=True)
            elif verbose and not result.success:
                print(f"  {name}: FAILED ({result.error})", flush=True)

            # Clean up file
            output_path.unlink(missing_ok=True)

    # Clean up temp directory
    with contextlib.suppress(OSError):
        temp_dir.rmdir()

    # Keep failed libraries in the report; absence is not a successful comparison.
    summaries: dict[str, BenchmarkSummary] = {}
    for name, run_results in results.items():
        successful = [r for r in run_results if r.success]
        times = [r.time_seconds for r in successful]
        median_time = statistics.median(times) if times else None
        summaries[name] = BenchmarkSummary(
            library=name,
            median_time=median_time,
            stdev_time=(statistics.stdev(times) if len(times) > 1 else 0.0) if times else None,
            rows_per_second=rows / median_time if median_time else None,
            file_size_mb=statistics.mean(r.file_size_mb for r in successful) if successful else None,
            speedup_vs_xlsxturbo=None,
            all_times=times,
            attempted_runs=len(run_results),
            errors=failures[name],
        )

    baseline = summaries.get("xlsxturbo")
    if baseline and not baseline.errors and baseline.median_time:
        for summary in summaries.values():
            if summary.median_time is not None:
                summary.speedup_vs_xlsxturbo = summary.median_time / baseline.median_time

    return summaries


def format_console_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict[str, object],
    shape: str = "mixed",
) -> str:
    """Format results for console output."""
    import xlsxturbo

    lines: list[str] = []
    lines.append("")
    lines.append(f"xlsxturbo Benchmark Suite v{xlsxturbo.version()}")
    lines.append("=" * 75)
    lines.append(f"System: {system_info['platform']} {system_info['platform_release']}, "
                 f"Python {system_info['python_version']}, {system_info['cpu_count']} CPUs")
    lines.append("")
    lines.append(f"Dataset: {rows:,} rows x {cols} columns ({shape})")
    lines.append(f"Requested runs: {runs} (medians use successful runs; counts shown per library)")
    lines.append("")

    # Sort by time (fastest first)
    sorted_summaries = sorted(
        summaries.values(), key=lambda s: s.median_time if s.median_time is not None else math.inf,
    )

    # Header
    with_memory = any(s.peak_memory_mb is not None for s in sorted_summaries)
    lines.append(
        f"{'Library':<22} {'Time (s)':>10} {'Stdev':>8} "
        f"{'Rows/sec':>12} {'Size (MB)':>10} {'vs xlsxturbo':>13} {'Passed':>7}"
        + (f" {'Peak +MB':>9}" if with_memory else "")
    )
    lines.append("-" * (102 if with_memory else 92))

    for summary in sorted_summaries:
        speedup_str = "-" if summary.speedup_vs_xlsxturbo is None else f"{summary.speedup_vs_xlsxturbo:.1f}x"
        if summary.library == "xlsxturbo" and summary.speedup_vs_xlsxturbo is not None:
            speedup_str = "1.0x (base)"

        lines.append(
            f"{summary.library:<22} "
            f"{_measurement(summary.median_time, '10.2f')} "
            f"{_measurement(summary.stdev_time, '8.3f')} "
            f"{_measurement(summary.rows_per_second, '12,.0f')} "
            f"{_measurement(summary.file_size_mb, '10.1f')} "
            f"{speedup_str:>13} {len(summary.all_times):>3}/{summary.attempted_runs:<3}"
            + (f" {_memory_cell(summary):>9}" if with_memory else "")
        )

    lines.append("")

    for summary in sorted_summaries:
        lines.extend(f"FAILED {summary.library}: {error}" for error in summary.errors)

    return "\n".join(lines)


def _measurement(value: float | None, spec: str) -> str:
    """Format a measured value, or a dash when no run succeeded.

    Args:
        value: Measurement or None.
        spec: Numeric format specification.

    Returns:
        The formatted measurement.
    """
    return "-" if value is None else format(value, spec)


def _memory_cell(summary: BenchmarkSummary) -> str:
    """The peak-memory column's text for one library.

    Args:
        summary: The library's results.

    Returns:
        Whole megabytes, or a dash when it was not measured.
    """
    return "-" if summary.peak_memory_mb is None else f"{summary.peak_memory_mb:,.0f}"


def format_markdown_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict[str, object],
    shape: str = "mixed",
) -> str:
    """Format results as markdown table."""
    import xlsxturbo

    lines: list[str] = []
    lines.append("## xlsxturbo Benchmark Results")
    lines.append("")
    lines.append(f"**System:** {system_info['platform']} {system_info['platform_release']}, "
                 f"Python {system_info['python_version']}, {system_info['cpu_count']} CPUs")
    lines.append(f"**xlsxturbo version:** {xlsxturbo.version()}")
    lines.append(f"**Dataset:** {rows:,} rows x {cols} columns, `{shape}` data (`--shape {shape}`)")
    packages = cast("dict[str, str]", system_info["packages"])
    lines.append("**Packages:** " + ", ".join(f"{name} {ver}" for name, ver in packages.items()))
    lines.append(f"**Requested runs:** {runs} (medians use successful runs; counts shown per library)")
    lines.append("")

    # Sort by time (fastest first)
    sorted_summaries = sorted(
        summaries.values(), key=lambda s: s.median_time if s.median_time is not None else math.inf,
    )

    with_memory = any(s.peak_memory_mb is not None for s in sorted_summaries)
    if with_memory:
        lines.append(
            "| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | "
            "Passed/attempted | Peak memory (+MB) |"
        )
        lines.append("|---------|----------|-------|----------|-----------|--------------|------------------|-------------------|")
    else:
        lines.append("| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | Passed/attempted |")
        lines.append("|---------|----------|-------|----------|-----------|--------------|------------------|")

    max_stdev_pct = max(
        (s.stdev_time / s.median_time * 100 for s in sorted_summaries
         if s.median_time and s.stdev_time is not None), default=0.0,
    )

    for summary in sorted_summaries:
        speedup_str = "-" if summary.speedup_vs_xlsxturbo is None else f"{summary.speedup_vs_xlsxturbo:.1f}x"
        if summary.library == "xlsxturbo" and summary.speedup_vs_xlsxturbo is not None:
            speedup_str = "**1.0x**"

        name = summary.library
        if summary.library == "xlsxturbo":
            name = "**xlsxturbo**"

        lines.append(
            f"| {name} | {_measurement(summary.median_time, '.2f')} | {_measurement(summary.stdev_time, '.3f')} | "
            f"{_measurement(summary.rows_per_second, ',.0f')} | {_measurement(summary.file_size_mb, '.1f')} | "
            f"{speedup_str} | {len(summary.all_times)}/{summary.attempted_runs} |"
            + (f" {_memory_cell(summary)} |" if with_memory else "")
        )

    lines.append("")
    lines.append(
        "*Median of successful runs after warmup; "
        f"max stdev across libraries: {max_stdev_pct:.1f}% of median.*"
        + (
            " *Peak memory is how far one write, in a fresh process, raises peak resident "
            "memory above the peak after building the frame.*"
            if with_memory
            else ""
        )
    )

    for summary in sorted_summaries:
        for error in summary.errors:
            lines.append(f"- FAILED {summary.library}: {error}")

    return "\n".join(lines)


def format_json_output(
    summaries: dict[str, BenchmarkSummary],
    rows: int,
    cols: int,
    runs: int,
    system_info: dict[str, object],
    shape: str = "mixed",
) -> str:
    """Format results as JSON for CI integration."""
    result = {
        "system": system_info,
        "benchmark": {
            "rows": rows,
            "cols": cols,
            "runs": runs,
            "shape": shape,
            "column_cycle": list(SHAPES[shape]),
        },
        "results": [
            {
                "library": s.library,
                "attempted_runs": s.attempted_runs,
                "successful_runs": len(s.all_times),
                "errors": s.errors,
                "median_time_seconds": None if s.median_time is None else round(s.median_time, 3),
                "stdev_time_seconds": None if s.stdev_time is None else round(s.stdev_time, 3),
                "rows_per_second": None if s.rows_per_second is None else round(s.rows_per_second, 0),
                "file_size_mb": None if s.file_size_mb is None else round(s.file_size_mb, 2),
                "speedup_vs_xlsxturbo": None if s.speedup_vs_xlsxturbo is None else round(s.speedup_vs_xlsxturbo, 2),
                "all_times": [round(t, 3) for t in s.all_times],
                "peak_memory_mb": None if s.peak_memory_mb is None else round(s.peak_memory_mb, 1),
            }
            for s in sorted(summaries.values(), key=lambda s: s.median_time if s.median_time is not None else math.inf)
        ],
    }
    return json.dumps(result, indent=2)


# Predefined benchmark sizes
BENCHMARK_SIZES = {
    "tiny": (1_000, 10),
    "small": (10_000, 20),
    "medium": (100_000, 50),
    "large": (500_000, 50),
}


def main() -> int:
    """Parse arguments, run the configured benchmarks, and print results."""
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
        help="Run full benchmark (tiny, small, medium, large sizes)",
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
        "--shape",
        choices=sorted(SHAPES),
        default="mixed",
        help="Column types: mixed (the reference workload), numeric, or strings (default: mixed)",
    )
    parser.add_argument(
        "--memory",
        action="store_true",
        help="Also measure each library's peak memory, one fresh process per library (not on Windows)",
    )
    parser.add_argument("--memory-child", help=argparse.SUPPRESS)
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

    if args.memory_child:
        if args.rows is None or args.cols is None:
            parser.error("--memory-child needs --rows and --cols")
        run_memory_child(args.memory_child, args.rows, args.cols, args.shape)
        return 0
    if args.memory and sys.platform == "win32":
        parser.error("--memory needs the resource module, which Windows does not have")

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

    all_outputs: list[str] = []
    failed = False

    for size_name, (rows, cols) in sizes:
        if verbose:
            print(f"\n{'=' * 75}")
            print(f"Benchmark: {size_name} ({rows:,} rows x {cols} columns)")
            print(f"{'=' * 75}")
            print("Generating test data...", flush=True)

        # Generate test data
        df_pd = generate_test_data(rows, cols, shape=args.shape)

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

        failed = failed or not summaries or any(s.errors or not s.all_times for s in summaries.values())

        if args.memory:
            if verbose:
                print("Measuring peak memory (one process per library)...", flush=True)
            for name, summary in summaries.items():
                if not summary.all_times:
                    continue
                summary.peak_memory_mb = measure_peak_memory(name, rows, cols, args.shape)
                if summary.peak_memory_mb is None:
                    summary.errors.append("memory run failed; see stderr for details")
                    failed = True

        # Format output
        if args.json:
            output = format_json_output(summaries, rows, cols, args.runs, system_info, args.shape)
        elif args.markdown:
            output = format_markdown_output(summaries, rows, cols, args.runs, system_info, args.shape)
        else:
            output = format_console_output(summaries, rows, cols, args.runs, system_info, args.shape)

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

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
