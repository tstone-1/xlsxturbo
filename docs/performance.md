# Performance

The numbers below are measured, machine-specific, and reproducible with the scripts in
`benchmarks/`. Treat them as a shape rather than a promise: the ratio between libraries
is stable across systems, the absolute times are not.

## Performance

*Reference benchmark on 100,000 rows x 50 columns with mixed data types. Your results will vary by system - run the benchmark yourself (see [Benchmarking](#benchmarking)).*

*All libraries use default settings; outputs differ in styling (e.g. polars auto-sizes columns and bolds headers by default, while xlsxturbo writes bare cells unless asked).*

### Historical Windows 11 / AMD Ryzen 9

*Historical result retained for reference. Dispersion and output-size measurements were not captured, so this table is not directly comparable to the current reproducible benchmark below.*

| Library | Time (s) | Rows/sec | vs xlsxturbo |
|---------|----------|----------|--------------|
| **xlsxturbo** | **4.76** | **21,010** | **1.0x** |
| polars | 18.33 | 5,455 | 3.9x |
| pandas + xlsxwriter | 27.66 | 3,615 | 5.8x |
| pandas + openpyxl | 35.36 | 2,828 | 7.4x |

*Test system: Windows 11, Python 3.14, AMD Ryzen 9 (32 threads). Median of 3 runs after warmup; standard deviation was not recorded.*

### macOS / MacBook

| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo |
|---------|----------|-------|----------|-----------|--------------|
| **xlsxturbo** | **3.19** | 0.028 | **31,396** | 47.6 | **1.0x** |
| polars | 14.61 | 0.069 | 6,845 | 48.4 | 4.6x |
| pandas + xlsxwriter | 22.41 | 0.134 | 4,463 | 50.0 | 7.0x |
| pandas + openpyxl | 29.54 | 0.267 | 3,386 | 50.3 | 9.3x |

*Test system: macOS (Darwin 25.5.0), Python 3.14.5, 10 CPUs. Median of 3 runs after warmup; max stdev across libraries: 0.9% of median. Re-run with `--markdown` to regenerate the current-system table.*

Benchmark scripts can also emit markdown or JSON, which makes it easy to attach benchmark output to issues, release notes, or CI artifacts.

## Threads

Exporting several workbooks at once from a `ThreadPoolExecutor` is worth doing: the GIL is
released while the archive is serialised and compressed, which is the larger half of a
`df_to_xlsx` call. Two threads finish a batch in about 55% of the time one thread takes and
four in about 43%, measured on 32 cores with 8000-row frames. Eight threads measured the
same as four — the gain plateaus at roughly 2.3x, and a smaller machine will reach that
ceiling no later.

```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=4) as pool:
    list(pool.map(lambda job: xlsxturbo.df_to_xlsx(job.frame, job.path), jobs))
```

The remaining half — reading values out of the DataFrame — holds the GIL, so the speedup
flattens out well short of the thread count. Threads also share one process's memory, which
is the reason to prefer them over processes here: a `ThreadPoolExecutor` does not copy the
frame, a `ProcessPoolExecutor` pickles it to every worker.

Each call writes its own file and shares nothing, so no locking is needed on your side. One
`DataFrame` may safely be read by several threads at once, provided nothing mutates it while
they run. `csv_to_xlsx` has released the GIL for its whole conversion since it was written,
and scales further because of it.

## Benchmarking

Run the included benchmark scripts:

```bash
# Compare xlsxturbo vs other libraries (100K rows default)
python benchmarks/benchmark.py

# Full benchmark: small, medium, large datasets
python benchmarks/benchmark.py --full

# Custom size
python benchmarks/benchmark.py --rows 500000 --cols 100

# Output formats for CI/documentation
python benchmarks/benchmark.py --markdown
python benchmarks/benchmark.py --json

# Test parallel vs single-threaded CSV conversion
python benchmarks/benchmark_parallel.py
```
