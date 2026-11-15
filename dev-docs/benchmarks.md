# Benchmarks

Moved verbatim from `AGENTS.md` on 2026-11-15, which keeps one index line per section.

## Benchmarks

- The main comparison suite is `benchmarks/benchmark.py`; use `--markdown` to regenerate the tables on `docs/performance.md` and `--json` for machine-readable output.
- The parallel CSV conversion suite is `benchmarks/benchmark_parallel.py`; the compression-level and thread-scaling tables come from `benchmarks/benchmark_compression.py` and `benchmarks/benchmark_threads.py`. All three print the published table with `--markdown`, and each refuses an output that does not unpack to the same workbook as its reference (`archive_contents` in `benchmark.py`).
- The compared writers, levels and thread counts run in an order rotated by one place per run (`run_order` in `benchmark.py`), so none always takes the first or last slot.
- `--memory` needs the `resource` module, so the memory column comes from a macOS or Linux machine. Which machines the published figures come from, their clones and interpreters, and how to reach them are in the gitignored `AGENTS.local.md`, because machine names and personal workflow stay out of this public repo.
- Published performance numbers are system-specific and must identify the machine, OS, Python version, and run methodology.
- Keep comparisons reproducible and fair: seed generated data, use native-fast dtypes for every compared library, perform warmup runs, report medians and standard deviations, and keep both benchmark suites methodologically aligned.
- Prefer honest, reproducible results over flattering headline numbers, including when a fairer method reduces the reported speedup.
- Generate measured documentation claims (benchmark results, variance, counts, and similar values) from their source script when practical; avoid hand-maintained factoids that silently become stale.
