#!/usr/bin/env python3
"""
Benchmark script for xlsxturbo parallel processing.
Compares single-threaded vs multi-threaded CSV to XLSX conversion.
"""

import os
import sys
import time
import tempfile
import random
import string
from datetime import date, datetime, timedelta

def generate_test_csv(filepath: str, rows: int, cols: int):
    """Generate a test CSV with mixed data types."""
    print(f"Generating test CSV: {rows:,} rows x {cols} columns...")

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

def benchmark_conversion(csv_path: str, parallel: bool, runs: int = 3):
    """Benchmark CSV to XLSX conversion."""
    import xlsxturbo

    mode = "parallel" if parallel else "single-threaded"
    times = []

    for run in range(runs):
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            xlsx_path = tmp.name

        try:
            start = time.perf_counter()
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, parallel=parallel)
            elapsed = time.perf_counter() - start
            times.append(elapsed)

            if run == 0:
                xlsx_size = os.path.getsize(xlsx_path) / (1024 * 1024)
                print(f"  {mode}: {rows:,} rows x {cols} cols -> {xlsx_size:.1f} MB")
        finally:
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    avg_time = sum(times) / len(times)
    min_time = min(times)
    return avg_time, min_time

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

        # Benchmark single-threaded
        print(f"Benchmarking single-threaded ({args.runs} runs)...")
        single_avg, single_min = benchmark_conversion(csv_path, parallel=False, runs=args.runs)
        print(f"  Average: {single_avg:.2f}s, Best: {single_min:.2f}s")
        print()

        # Benchmark parallel
        print(f"Benchmarking parallel ({args.runs} runs)...")
        parallel_avg, parallel_min = benchmark_conversion(csv_path, parallel=True, runs=args.runs)
        print(f"  Average: {parallel_avg:.2f}s, Best: {parallel_min:.2f}s")
        print()

        # Results
        print("=" * 60)
        print("Results:")
        print("=" * 60)
        speedup_avg = single_avg / parallel_avg
        speedup_min = single_min / parallel_min
        print(f"Single-threaded: {single_avg:.2f}s (avg), {single_min:.2f}s (best)")
        print(f"Parallel:        {parallel_avg:.2f}s (avg), {parallel_min:.2f}s (best)")
        print(f"Speedup:         {speedup_avg:.2f}x (avg), {speedup_min:.2f}x (best)")

        if speedup_avg > 1:
            print(f"\n[OK] Parallel processing is {speedup_avg:.2f}x faster!")
        else:
            print(f"\n[INFO] Parallel processing is slower for this dataset size.")
            print("       Try with larger files (1M+ rows) for better results.")

    finally:
        if os.path.exists(csv_path):
            os.unlink(csv_path)

if __name__ == '__main__':
    main()
