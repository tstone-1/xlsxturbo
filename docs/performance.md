# Performance

The numbers below are measured, machine-specific, and reproducible with the scripts in
`benchmarks/`. With each library's defaults, the ratios between libraries differed by
at most 20% between the two machines measured so far, but that is two machines and one
workload shape, not a guarantee for yours.

## Performance

*Reference benchmark on 100,000 rows x 50 columns with mixed data types. Your results will vary by system - run the benchmark yourself (see [Benchmarking](#benchmarking)).*

Two workloads, because the libraries' defaults do not produce the same file:

- **Styled** (`--styled`): every writer produces the same report, an Excel table in
  `TableStyleMedium2`, `#,##0` / `#,##0.00` / `yyyy-mm-dd` number formats by column
  type, and a width of 14 on every column. Each output is checked for all three.
- **Defaults**: each library's own default output. `polars.write_excel` wraps the data
  in an Excel table and gives numeric columns thousands-separator formats, while
  xlsxturbo and pandas write bare cells.

In both, every header and data cell is compared with the generated frame after each
timed write, so a library that silently writes less cannot look fast. Floating-point
comparisons allow the rounding used by the compared writers. The readback is outside
the timing and is excluded from the memory measurement.

### macOS / MacBook Pro, styled

| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | Peak memory (+MB) |
|---------|----------|-------|----------|-----------|--------------|-------------------|
| **xlsxturbo** | **3.73** | 0.051 | **26,810** | 48.4 | **1.0x** | 1,110 |
| polars | 13.31 | 0.205 | 7,516 | 48.4 | 3.6x | 1,529 |
| pandas + xlsxwriter | 22.22 | 0.184 | 4,501 | 50.2 | 6.0x | 1,637 |
| pandas + openpyxl | 32.07 | 0.358 | 3,118 | 51.3 | 8.6x | 2,213 |

### macOS / MacBook Pro, defaults

| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | Peak memory (+MB) |
|---------|----------|-------|----------|-----------|--------------|-------------------|
| **xlsxturbo** | **3.03** | 0.048 | **32,969** | 47.6 | **1.0x** | 920 |
| polars | 13.20 | 0.037 | 7,577 | 48.4 | 4.4x | 1,529 |
| pandas + xlsxwriter | 20.34 | 0.398 | 4,918 | 50.0 | 6.7x | 1,019 |
| pandas + openpyxl | 27.04 | 0.263 | 3,699 | 50.3 | 8.9x | 1,894 |

*Test system: MacBook Pro (Mac17,2, Apple M5, 10 CPUs), macOS (Darwin 27.0.0), Python 3.14.6; xlsxturbo 1.7.0, pandas 3.0.6, polars 1.44.2, numpy 2.5.3, openpyxl 3.1.5, xlsxwriter 3.2.9. Median of 3 runs after warmup; max stdev across libraries 1.5% (styled) and 2.0% (defaults) of median. All runs passed their output checks. Peak memory is how far one write, in a fresh process, raises peak resident memory above the peak after building the frame. Regenerate with `python benchmarks/benchmark.py --markdown --memory [--styled]`.*

The styling costs every library something, but not the same amount: formatting takes
xlsxturbo from 3.03 s to 3.73 s and pandas + openpyxl from 27.0 s to 32.1 s, where
openpyxl sets the number format cell by cell. Peak memory for pandas + xlsxwriter rises
most, from 1,019 MB to 1,637 MB.

### Historical Windows 11 / AMD Ryzen 9, defaults

*Historical result retained for reference. It predates the cell-by-cell output check, and dispersion, output size and memory were not captured, so it is not directly comparable to the tables above.*

| Library | Time (s) | Rows/sec | vs xlsxturbo |
|---------|----------|----------|--------------|
| **xlsxturbo** | **4.76** | **21,010** | **1.0x** |
| polars | 18.33 | 5,455 | 3.9x |
| pandas + xlsxwriter | 27.66 | 3,615 | 5.8x |
| pandas + openpyxl | 35.36 | 2,828 | 7.4x |

*Test system: Windows 11, Python 3.14, AMD Ryzen 9 (32 threads). Median of 3 runs after warmup; standard deviation was not recorded.*

Benchmark scripts can also emit markdown or JSON, which makes it easy to attach benchmark output to issues, release notes, or CI artifacts.

Reports include successful and attempted run counts and any failures. A failed run or
warmup makes the command exit with status 1. Failed libraries remain in the report with
unavailable measurements, and a failed xlsxturbo baseline produces no speedup ratios.

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

# Data shape: mixed (the reference), numeric, or strings
python benchmarks/benchmark.py --shape strings

# Every writer produces the same table, number formats and column widths
python benchmarks/benchmark.py --styled

# Also measure each library's peak memory, in a fresh process per library (macOS, Linux)
python benchmarks/benchmark.py --memory

# Output formats for CI/documentation
python benchmarks/benchmark.py --markdown
python benchmarks/benchmark.py --json

# Test parallel vs single-threaded CSV conversion
python benchmarks/benchmark_parallel.py
```
