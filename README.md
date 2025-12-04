# fast_xlsx

High-performance CSV to XLSX converter written in Rust using `rust_xlsxwriter`.

## Performance

Benchmarked on 525,684 rows x 98 columns:

| Method | Time | Speedup |
|--------|------|---------|
| **fast_xlsx (Rust)** | 28.5s | **26.7x** |
| PyExcelerate | 107s | 7.1x |
| pandas + xlsxwriter | 374s | 2.0x |
| pandas + openpyxl | 762s | 1.0x |
| polars.write_excel | 1039s | 0.7x |

## Usage

```bash
fast_xlsx input.csv output.xlsx [--sheet-name "Sheet1"] [-v]
```

### Options

- `-s, --sheet-name`: Name of the Excel sheet (default: "Sheet1")
- `-v, --verbose`: Show progress information

### Example

```bash
fast_xlsx data.csv report.xlsx -v
```

## Building

Requires Rust toolchain:

```bash
cargo build --release
```

Binary will be at `target/release/fast_xlsx.exe`

## Python Integration

Use with the `ecoglobal_functions.fast_to_excel()` wrapper:

```python
import ecoglobal
import ecoglobal_functions as ef

df = ecoglobal.Report('mag').load()
ef.fast_to_excel(df, 'output.xlsx', verbose=True)
```

## Features

- Automatic type detection (numbers, booleans, strings)
- Handles NaN/Inf values gracefully
- 1MB read buffer for optimal performance
- Progress reporting with `-v` flag
- Optimized release build (LTO, stripped binary)

## License

MIT
