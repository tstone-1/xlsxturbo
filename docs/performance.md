# Performance

The numbers below are measured, machine-specific, and reproducible with the scripts in
`benchmarks/`. Measured on four machines (two Apple Silicon Macs and two Windows PCs),
the highest and lowest value of each ratio between libraries differed by at most 18%.
That is four machines and one workload shape, not a guarantee for yours.

## Comparison with other libraries

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
the timing and is excluded from the memory measurement. The libraries run in an order
that rotates by one place per run, so none always runs first or last.

### macOS / MacBook Pro, styled

| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | Peak memory (+MB) |
|---------|----------|-------|----------|-----------|--------------|-------------------|
| **xlsxturbo** | **3.68** | 0.002 | **27,152** | 48.4 | **1.0x** | 1,133 |
| polars | 13.16 | 0.011 | 7,600 | 48.4 | 3.6x | 1,550 |
| pandas + xlsxwriter | 21.73 | 0.151 | 4,601 | 50.2 | 5.9x | 1,657 |
| pandas + openpyxl | 31.90 | 0.148 | 3,135 | 51.3 | 8.7x | 2,229 |

### macOS / MacBook Pro, defaults

| Library | Time (s) | Stdev | Rows/sec | Size (MB) | vs xlsxturbo | Peak memory (+MB) |
|---------|----------|-------|----------|-----------|--------------|-------------------|
| **xlsxturbo** | **3.02** | 0.020 | **33,062** | 47.6 | **1.0x** | 941 |
| polars | 13.38 | 0.107 | 7,476 | 48.4 | 4.4x | 1,550 |
| pandas + xlsxwriter | 20.12 | 0.117 | 4,970 | 50.0 | 6.7x | 1,039 |
| pandas + openpyxl | 26.81 | 0.172 | 3,730 | 50.3 | 8.9x | 1,910 |

*Test system: MacBook Pro (Mac17,2, Apple M5, 10 CPUs), macOS (Darwin 27.0.0), Python 3.14.6; xlsxturbo 1.12.2, pandas 3.1.0, polars 2.1.0, numpy 2.5.4, openpyxl 3.1.5, xlsxwriter 3.3.1. Median of 3 runs after warmup; max stdev across libraries 0.7% (styled) and 0.8% (defaults) of median. All runs passed their output checks. Peak memory is how far one write, in a fresh process, raises peak resident memory above the peak after building the frame. Regenerate with `python benchmarks/benchmark.py --markdown --memory [--styled]`.*

The styling costs every library something, but not the same amount: formatting takes
xlsxturbo from 3.02 s to 3.68 s and pandas + openpyxl from 26.8 s to 31.9 s, where
openpyxl sets the number format cell by cell. Peak memory for pandas + xlsxwriter rises
most, from 1,039 MB to 1,657 MB.

The same machine with the same Python measured xlsxturbo 1.7.0 at 3.03 s (defaults) and
3.73 s (styled), and polars 1.44.2 at 13.20 s with its defaults, so the ratios have not
moved between those releases.

### Four machines

The same two workloads on every machine. Only xlsxturbo's time is given; the other
columns are how many times longer each library took.

| Machine | Workload | xlsxturbo | polars | pandas + xlsxwriter | pandas + openpyxl | Max stdev |
|---------|----------|----------:|-------:|--------------------:|------------------:|----------:|
| MacBook Pro, Apple M5 | styled | 3.68 s | 3.6x | 5.9x | 8.7x | 0.7% |
| Mac mini, Apple M4 | styled | 4.21 s | 3.9x | 6.5x | 9.3x | 0.7% |
| Windows 11, Ryzen 5 8500GE (12 threads) | styled | 6.97 s | 3.5x | 6.0x | 8.9x | 0.9% |
| Windows 11, Ryzen 9 (32 threads) | styled | 6.13 s | 3.4x | 5.7x | 8.8x | 5.4% |
| MacBook Pro, Apple M5 | defaults | 3.02 s | 4.4x | 6.7x | 8.9x | 0.8% |
| Mac mini, Apple M4 | defaults | 3.44 s | 4.7x | 7.2x | 9.5x | 0.8% |
| Windows 11, Ryzen 5 8500GE (12 threads) | defaults | 5.75 s | 4.3x | 6.6x | 8.9x | 0.9% |
| Windows 11, Ryzen 9 (32 threads) | defaults | 5.05 s | 4.1x | 6.1x | 8.5x | 7.9% |

*xlsxturbo 1.12.2 and the package versions above on every machine; Python 3.14.6 (MacBook Pro), 3.14.9 (Mac mini), 3.14.7 (Ryzen 5) and 3.14.10 (Ryzen 9). Median of 3 runs after warmup, 5 on the Ryzen 9, which is a desktop in daily use and the noisiest of the four. All runs passed their output checks. Peak memory needs the `resource` module and was measured on the Macs only; the Mac mini's figures match the MacBook's to within 1 MB.*

## Threads

Exporting several workbooks at once from a `ThreadPoolExecutor` is worth doing: the GIL is
released while the archive is serialised and compressed, which is the larger half of a
`df_to_xlsx` call. `csv_to_xlsx` releases it for the whole conversion, so it scales further.

| Threads | `df_to_xlsx` batch | Speedup | `csv_to_xlsx` batch | Speedup |
|--------:|-------------------:|--------:|--------------------:|--------:|
| 1 | 1.19 s | 1.00x | 1.03 s | 1.00x |
| 2 | 0.62 s | 1.91x | 0.55 s | 1.87x |
| 4 | 0.46 s | 2.58x | 0.29 s | 3.57x |
| 8 | 0.46 s | 2.56x | 0.21 s | 4.99x |

*Test system: MacBook Pro (Apple M5, 10 CPUs), Python 3.14.6, xlsxturbo 1.12.2. 32 exports to `io.BytesIO` per batch, of one seeded 8,000-row x 8-column mixed frame and of the same frame saved as CSV; functions and thread counts interleaved, median of 7 runs after a warmup. Every export unpacks to the same workbook as a serial one of the same function. Regenerate with `python benchmarks/benchmark_threads.py --markdown`.*

On the other three machines `df_to_xlsx` gave 2.07x on four threads (Ryzen 5, 12 threads),
1.83x (Ryzen 9, 32 threads) and 2.36x (Mac mini), and eight threads were at most 1% faster
than four on any of them. Its gain stops at about four threads, at 1.8x to 2.6x, whatever
the core count. `csv_to_xlsx` was still gaining at eight threads on every machine: 4.72x
(Mac mini), 4.83x (Ryzen 5) and 5.67x (Ryzen 9).

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
they run.

## Compression level

Compressing the archive is a large part of every save. `compression_level` sets the zip
compression from 1 (fastest, largest file) to 9; the default is 6. It is accepted by
`df_to_xlsx`, `dfs_to_xlsx` and `csv_to_xlsx`, and by the command as `-c` /
`--compression-level`. It changes only how the file is packed: Excel reads every level the
same.

```python
xlsxturbo.df_to_xlsx(df, "report.xlsx", compression_level=1)
```

| Level | `df_to_xlsx` | `csv_to_xlsx` | File size |
|------:|-------------:|--------------:|----------:|
| 1 | 276 ms (-27%) | 214 ms (-32%) | +64% |
| 2 | 285 ms (-25%) | 222 ms (-30%) | +16% |
| 3 | 296 ms (-22%) | 235 ms (-26%) | +7% |
| 4 | 314 ms (-17%) | 252 ms (-20%) | +2% |
| 5 | 322 ms (-15%) | 260 ms (-18%) | +0.4% |
| 6 (default) | 379 ms | 316 ms | 5.1 MiB |
| 7 | 439 ms (+16%) | 377 ms (+19%) | -5% |
| 8 | 882 ms (+133%) | 818 ms (+159%) | -5% |
| 9 | 656 ms (+73%) | 593 ms (+88%) | +0.2% |

*Test system: MacBook Pro (Apple M5, 10 CPUs), Python 3.14.6, xlsxturbo 1.12.2. A seeded 100,000-row frame of seven columns (integers, two float columns, 20 repeated strings, unique ids, booleans, dates), written to `io.BytesIO`; levels interleaved, median of 9 runs after a warmup. File size is the `df_to_xlsx` output. Every level unpacks to the same workbook. Regenerate with `python benchmarks/benchmark_compression.py --markdown`.*

The other three machines measured the same pattern: level 1 saved 25-28% of the
`df_to_xlsx` time and 31-34% of the `csv_to_xlsx` time, level 5 saved 13-15% and 14-18%,
and the file sizes were identical to these. Level 5 is close to free: a file 0.4% larger,
about a seventh faster. Level 1 suits files that are read once and deleted, or that go
straight into another archive. Levels 7 and 8 write a file 5% smaller, level 7 for 14-16%
more time and level 8 for 94-133% more; level 8 was the slowest level on every machine.
Level 9 is faster than 8 but writes a file 0.2% larger than level 6, so it is never worth
choosing for data like this.

## Parallel CSV parsing

`csv_to_xlsx(..., parallel=True)` (the command's `-p` / `--parallel`) parses the CSV on
several cores. It helps, but by a few percent rather than a multiple:

| Mode | Time (s) | Stdev | Speedup |
|------|---------:|------:|--------:|
| single-threaded | 13.79 | 0.086 | 1.00x |
| `parallel=True` | 12.78 | 0.225 | 1.08x |

*Test system: MacBook Pro (Apple M5, 10 CPUs), Python 3.14.6, xlsxturbo 1.12.2. A seeded CSV of 500,000 rows x 50 columns (integers, floats, booleans, dates, datetimes, strings); modes paired and alternated, median of 3 runs after a warmup. Both modes write the same workbook. Regenerate with `python benchmarks/benchmark_parallel.py --markdown`.*

The other machines measured 1.08x (Mac mini), 1.06x (Ryzen 9, 32 threads) and 1.05x
(Ryzen 5, 12 threads), so more cores did not raise the gain.

## Python 3.15

The interpreter makes almost no difference to xlsxturbo, because the time is spent in Rust.
Two venvs on one machine, identical except for the interpreter (the same xlsxturbo wheel,
the same package versions), ran the styled comparison in three interleaved rounds:

| Library | Python 3.14.10 | Python 3.15.2 | Change |
|---------|---------------:|--------------:|-------:|
| **xlsxturbo** | 3.68 s | 3.64 s | -1.1% |
| polars | 13.16 s | 13.21 s | +0.4% |
| pandas + xlsxwriter | 21.54 s | 21.38 s | -0.7% |
| pandas + openpyxl | 31.62 s | 30.45 s | -3.7% |

*Test system: MacBook Pro (Apple M5, 10 CPUs); xlsxturbo 1.12.2, pandas 3.1.0, polars 2.1.0, numpy 2.5.4, openpyxl 3.1.5, xlsxwriter 3.3.1, pyarrow 26.0.0. 100,000 rows x 50 columns, `--styled`; median of three rounds, each the median of 3 runs after a warmup, with the interpreter that went first alternating between rounds. Every round's largest standard deviation was at most 1.2% of its median.*

Only pandas + openpyxl, the one pure-Python writer, got measurably faster on 3.15, and it
was faster in every round, so xlsxturbo's lead over it shrinks from 8.6x to 8.4x. Peak
memory was the same on both interpreters to within 4 MB for every library, and the thread
speedups of the [Threads](#threads) table were within 4% of each other.

## Benchmarking

Every table on this page is printed by one of the scripts below with `--markdown`.
Reports include successful and attempted run counts and any failures. A failed run, a
failed warmup or a failed output check makes the command exit with status 1. In
`benchmark.py`, failed libraries remain in the report with unavailable measurements, and
a failed xlsxturbo baseline produces no speedup ratios.

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

# Compression levels 1-9 for df_to_xlsx and csv_to_xlsx
python benchmarks/benchmark_compression.py

# A batch of exports on 1, 2, 4 and 8 threads
python benchmarks/benchmark_threads.py

# Parallel vs single-threaded CSV conversion
python benchmarks/benchmark_parallel.py
```
