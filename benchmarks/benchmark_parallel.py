#!/usr/bin/env python3
"""
Benchmark script for xlsxturbo parallel processing.
Compares single-threaded vs multi-threaded CSV to XLSX conversion.
"""

import os
import statistics
import sys
import time
import tempfile
import random
import string
from datetime import date, datetime, timedelta

def generate_test_csv(filepath: str, rows: int, cols: int, seed: int = 42):
    """Generate a test CSV with mixed data types."""
    print(f"Generating test CSV: {rows:,} rows x {cols} columns...")
    random.seed(seed)

    start = time.perf_counter()
    with open(filepath, 'w', encoding='utf-8') as f:
        # Header
        headers = [f"col_{i}" for i in range(cols)]
        f.write(','.join(headers) + '\n')

        # Data rows with mixed types
        base_date = date(2020, 1, 1)
        base_datetime = datetime(2020, 1, 1, 0, 0, 0)

        for row in range(rows):
            values = []
            for col in range(cols):
                col_type = col % 6
                if col_type == 0:
                    # Integer
                    values.append(str(random.randint(-10000, 10000)))
                elif col_type == 1:
                    # Float
                    values.append(f"{random.uniform(-1000, 1000):.4f}")
                elif col_type == 2:
                    # Boolean
                    values.append(random.choice(['true', 'false', 'TRUE', 'FALSE']))
                elif col_type == 3:
                    # Date
                    d = base_date + timedelta(days=random.randint(0, 1000))
                    values.append(d.strftime('%Y-%m-%d'))
                elif col_type == 4:
                    # Datetime
                    dt = base_datetime + timedelta(
                        days=random.randint(0, 1000),
                        hours=random.randint(0, 23),
                        minutes=random.randint(0, 59)
                    )
                    values.append(dt.strftime('%Y-%m-%dT%H:%M:%S'))
                else:
                    # String
                    length = random.randint(5, 20)
                    values.append(''.join(random.choices(string.ascii_letters, k=length)))

            f.write(','.join(values) + '\n')

    elapsed = time.perf_counter() - start
    file_size = os.path.getsize(filepath) / (1024 * 1024)
    print(f"  Generated in {elapsed:.2f}s ({file_size:.1f} MB)")
    return filepath

def benchmark_conversion(csv_path: str, parallel: bool, runs: int = 3, warmup: bool = True):
    """Benchmark CSV to XLSX conversion."""
    import xlsxturbo

    mode = "parallel" if parallel else "single-threaded"
    times = []

    total_runs = runs + (1 if warmup else 0)
    for run in range(total_runs):
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            xlsx_path = tmp.name

        try:
            is_warmup = warmup and run == 0
            if is_warmup:
                print(f"  warmup ({mode})...", flush=True)

            start = time.perf_counter()
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, parallel=parallel)
            elapsed = time.perf_counter() - start

            if not is_warmup:
                times.append(elapsed)

            if run == 0:
                xlsx_size = os.path.getsize(xlsx_path) / (1024 * 1024)
                print(f"  {mode}: {rows:,} rows x {cols} cols -> {xlsx_size:.1f} MB")
        finally:
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    median_time = statistics.median(times)
    stdev_time = statistics.stdev(times) if len(times) > 1 else 0.0
    return median_time, stdev_time

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Benchmark xlsxturbo parallel processing')
    parser.add_argument('--rows', type=int, default=500000, help='Number of rows (default: 500000)')
    parser.add_argument('--cols', type=int, default=50, help='Number of columns (default: 50)')
    parser.add_argument('--runs', type=int, default=3, help='Number of benchmark runs (default: 3)')
    args = parser.parse_args()

    print("=" * 60)
    print("xlsxturbo Parallel Processing Benchmark")
    print("=" * 60)

    import xlsxturbo
    print(f"xlsxturbo version: {xlsxturbo.version()}")
    print(f"CPU cores: {os.cpu_count()}")
    print()

    # Generate test data
    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
        csv_path = tmp.name

    try:
        generate_test_csv(csv_path, args.rows, args.cols)
        print()

        # Benchmark single-threaded (warmup + runs)
        print(f"Benchmarking single-threaded ({args.runs} runs + warmup)...")
        single_med, single_std = benchmark_conversion(csv_path, parallel=False, runs=args.runs)
        print(f"  Median: {single_med:.2f}s (stdev {single_std:.2f}s)")
        print()

        # Benchmark parallel (warmup + runs)
        print(f"Benchmarking parallel ({args.runs} runs + warmup)...")
        parallel_med, parallel_std = benchmark_conversion(csv_path, parallel=True, runs=args.runs)
        print(f"  Median: {parallel_med:.2f}s (stdev {parallel_std:.2f}s)")
        print()

        # Results
        print("=" * 60)
        print("Results:")
        print("=" * 60)
        speedup = single_med / parallel_med
        print(f"Single-threaded: {single_med:.2f}s (stdev {single_std:.2f}s)")
        print(f"Parallel:        {parallel_med:.2f}s (stdev {parallel_std:.2f}s)")
        print(f"Speedup:         {speedup:.2f}x")

        if speedup > 1:
            print(f"\n[OK] Parallel processing is {speedup:.2f}x faster!")
        else:
            print(f"\n[INFO] Parallel processing is slower for this dataset size.")
            print("       Try with larger files (1M+ rows) for better results.")

    finally:
        if os.path.exists(csv_path):
            os.unlink(csv_path)

if __name__ == '__main__':
    main()
