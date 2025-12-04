#!/usr/bin/env python3
"""
Benchmark script comparing Excel writing performance across different methods.

Compares:
- fast_xlsx (Rust binary via subprocess)
- PyExcelerate (optimized pure Python)
- pandas + xlsxwriter
- pandas + openpyxl
- polars.write_excel

Usage:
    python benchmark.py [--rows N] [--cols N]

Examples:
    python benchmark.py                    # Default: 100,000 rows x 50 cols
    python benchmark.py --rows 500000      # 500K rows
    python benchmark.py --rows 10000 --cols 100
"""

import argparse
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# Path to fast_xlsx binary (same directory as this script)
FAST_XLSX_EXE = Path(__file__).parent / "target" / "release" / "fast_xlsx.exe"


def get_file_size_mb(filepath: str) -> float:
    """Get file size in megabytes."""
    return os.path.getsize(filepath) / (1024 * 1024)


def generate_test_data(rows: int, cols: int):
    """Generate test DataFrames for benchmarking."""
    import pandas as pd
    import numpy as np

    print(f"Generating test data: {rows:,} rows x {cols} columns...")

    # Mix of data types for realistic benchmark
    data = {}
    for i in range(cols):
        col_type = i % 4
        if col_type == 0:
            data[f"int_{i}"] = np.random.randint(0, 1000000, rows)
        elif col_type == 1:
            data[f"float_{i}"] = np.random.random(rows) * 1000
        elif col_type == 2:
            data[f"str_{i}"] = [f"value_{j}" for j in range(rows)]
        else:
            data[f"bool_{i}"] = np.random.choice([True, False], rows)

    df_pd = pd.DataFrame(data)
    return df_pd


def benchmark_fast_xlsx(df_pd, output_path: str) -> tuple[float, float]:
    """Benchmark fast_xlsx Rust binary."""
    if not FAST_XLSX_EXE.exists():
        raise FileNotFoundError(f"fast_xlsx.exe not found at {FAST_XLSX_EXE}")

    # Write to temp CSV first
    temp_csv = Path(output_path).with_suffix(".tmp.csv")

    start = time.perf_counter()
    df_pd.to_csv(temp_csv, index=False)
    csv_time = time.perf_counter() - start

    # Run fast_xlsx
    xlsx_start = time.perf_counter()
    result = subprocess.run(
        [str(FAST_XLSX_EXE), str(temp_csv), output_path],
        capture_output=True,
        text=True,
        check=True,
    )
    xlsx_time = time.perf_counter() - xlsx_start

    # Clean up temp CSV
    temp_csv.unlink()

    elapsed = time.perf_counter() - start
    size_mb = get_file_size_mb(output_path)
    return elapsed, size_mb


def benchmark_pandas_openpyxl(df_pd, output_path: str) -> tuple[float, float]:
    """Benchmark pandas with openpyxl engine."""
    start = time.perf_counter()
    df_pd.to_excel(output_path, index=False, engine="openpyxl")
    elapsed = time.perf_counter() - start
    size_mb = get_file_size_mb(output_path)
    return elapsed, size_mb


def benchmark_pandas_xlsxwriter(df_pd, output_path: str) -> tuple[float, float]:
    """Benchmark pandas with xlsxwriter engine."""
    start = time.perf_counter()
    df_pd.to_excel(output_path, index=False, engine="xlsxwriter")
    elapsed = time.perf_counter() - start
    size_mb = get_file_size_mb(output_path)
    return elapsed, size_mb


def benchmark_polars_write_excel(df_pd, output_path: str) -> tuple[float, float]:
    """Benchmark polars built-in write_excel."""
    import polars as pl

    df_pl = pl.from_pandas(df_pd)
    start = time.perf_counter()
    df_pl.write_excel(output_path)
    elapsed = time.perf_counter() - start
    size_mb = get_file_size_mb(output_path)
    return elapsed, size_mb


def benchmark_pyexcelerate(df_pd, output_path: str) -> tuple[float, float]:
    """Benchmark PyExcelerate."""
    from pyexcelerate import Workbook

    start = time.perf_counter()
    data = [df_pd.columns.tolist()] + df_pd.values.tolist()
    wb = Workbook()
    wb.new_sheet("Sheet1", data=data)
    wb.save(output_path)
    elapsed = time.perf_counter() - start
    size_mb = get_file_size_mb(output_path)
    return elapsed, size_mb


def run_benchmark(name: str, func, *args) -> dict | None:
    """Run a single benchmark with error handling."""
    try:
        elapsed, size_mb = func(*args)
        return {"name": name, "time": elapsed, "size_mb": size_mb}
    except ImportError as e:
        print(f"  {name}: SKIPPED - {e}")
        return None
    except FileNotFoundError as e:
        print(f"  {name}: SKIPPED - {e}")
        return None
    except Exception as e:
        print(f"  {name}: FAILED - {e}")
        return None


def main():
    parser = argparse.ArgumentParser(description="Benchmark Excel writing methods")
    parser.add_argument("--rows", type=int, default=100000, help="Number of rows (default: 100000)")
    parser.add_argument("--cols", type=int, default=50, help="Number of columns (default: 50)")
    args = parser.parse_args()

    print("=" * 70)
    print("EXCEL WRITER BENCHMARK")
    print("=" * 70)
    print()

    # Generate test data
    df_pd = generate_test_data(args.rows, args.cols)
    rows, cols = df_pd.shape
    print(f"Data ready: {rows:,} rows x {cols} columns")
    print()

    # Create temp directory
    temp_dir = tempfile.mkdtemp(prefix="excel_benchmark_")
    print(f"Output directory: {temp_dir}")
    print()

    # Define benchmarks (fast_xlsx first as it's expected to be fastest)
    benchmarks = [
        ("fast_xlsx (Rust)", benchmark_fast_xlsx, df_pd, os.path.join(temp_dir, "fast_xlsx.xlsx")),
        ("PyExcelerate", benchmark_pyexcelerate, df_pd, os.path.join(temp_dir, "pyexcelerate.xlsx")),
        ("pandas + xlsxwriter", benchmark_pandas_xlsxwriter, df_pd, os.path.join(temp_dir, "pandas_xlsxwriter.xlsx")),
        ("pandas + openpyxl", benchmark_pandas_openpyxl, df_pd, os.path.join(temp_dir, "pandas_openpyxl.xlsx")),
        ("polars.write_excel", benchmark_polars_write_excel, df_pd, os.path.join(temp_dir, "polars.xlsx")),
    ]

    # Run benchmarks
    print("Running benchmarks...")
    print("-" * 70)
    results = []
    for name, func, df, output_path in benchmarks:
        print(f"  Testing {name}...", end=" ", flush=True)
        result = run_benchmark(name, func, df, output_path)
        if result:
            results.append(result)
            print(f"{result['time']:.2f}s ({result['size_mb']:.1f} MB)")
        print()

    if not results:
        print("No benchmarks completed successfully!")
        return 1

    # Display results
    print()
    print("=" * 70)
    print("RESULTS")
    print("=" * 70)
    print(f"Data size: {rows:,} rows x {cols} columns")
    print()

    # Sort by time (fastest first)
    results.sort(key=lambda x: x["time"])
    fastest_time = results[0]["time"]

    print(f"{'Method':<25} {'Time (s)':>10} {'Size (MB)':>10} {'Speedup':>10}")
    print("-" * 60)

    for r in results:
        speedup = r["time"] / fastest_time
        marker = " *" if speedup == 1.0 else ""
        print(f"{r['name']:<25} {r['time']:>10.2f} {r['size_mb']:>10.1f} {speedup:>9.1f}x{marker}")

    print()
    print(f"* = fastest")
    print()

    # Summary
    fastest = results[0]["name"]
    slowest = results[-1]["name"]
    max_speedup = results[-1]["time"] / results[0]["time"]
    print(f"Fastest: {fastest}")
    print(f"Slowest: {slowest}")
    print(f"Max speedup: {max_speedup:.1f}x")
    print()
    print(f"Output files saved to: {temp_dir}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
