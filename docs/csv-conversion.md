# CSV conversion

`csv_to_xlsx` reads a CSV and writes an `.xlsx`, detecting each value's type from its
text. It is a separate entry point from the DataFrame functions and takes far fewer
options -- see the [capability matrix](capability-matrix.md) for exactly which. The
[`xlsxturbo` command](#cli-usage) documented at the bottom of this page is a thin wrapper over it.

## CSV Conversion

```python
import xlsxturbo

# Convert CSV to XLSX with automatic type detection
rows, cols = xlsxturbo.csv_to_xlsx("input.csv", "output.xlsx")
print(f"Converted {rows} rows and {cols} columns")

# Custom sheet name
xlsxturbo.csv_to_xlsx("data.csv", "report.xlsx", sheet_name="Sales Data")

# For large files (100K+ rows), use parallel processing
xlsxturbo.csv_to_xlsx("big_data.csv", "output.xlsx", parallel=True)

# Handle ambiguous dates (01-02-2024: is it Jan 2 or Feb 1?)
xlsxturbo.csv_to_xlsx("us_data.csv", "output.xlsx", date_order="us")   # January 2
xlsxturbo.csv_to_xlsx("eu_data.csv", "output.xlsx", date_order="eu")   # February 1

# date_order options:
# - "auto" (default): ISO first, then European (DMY), then US (MDY)
# - "mdy" or "us": US format (MM-DD-YYYY)
# - "dmy" or "eu": European format (DD-MM-YYYY)
```

## Formula Injection

CSV and DataFrame string values are always written as literal string cells, never as formulas. A value starting with `=`, `+`, `-`, or `@` is stored as-is and does not execute in Excel. The only ways to produce a live formula are the explicit `formula_columns` option and the `hyperlinks` option; nothing else in xlsxturbo interprets cell content as a formula.

## CLI Usage

`pip install xlsxturbo` installs an `xlsxturbo` command (since 1.7.0). To try it without
installing anything into your environment:

```bash
uvx xlsxturbo input.csv output.xlsx      # or: pipx run xlsxturbo input.csv output.xlsx
```

```bash
xlsxturbo input.csv output.xlsx [OPTIONS]
python -m xlsxturbo input.csv output.xlsx [OPTIONS]    # the same command
```

A source build also produces a standalone Rust binary with the same options and exit
codes, for machines without Python: `cargo build --release` writes
`target/release/xlsxturbo`.

### Options

- `-s, --sheet-name <NAME>`: Name of the Excel sheet (default: "Sheet1")
- `-d, --date-order <ORDER>`: Date parsing order for ambiguous dates (default: "auto")
  - `auto`: ISO first, then European, then US
  - `mdy` or `us`: US format (01-02-2024 = January 2)
  - `dmy` or `eu`: European format (01-02-2024 = February 1)
- `-p, --parallel`: Use multi-core CSV processing (faster for large files, uses more memory)
- `-v, --verbose`: Show progress information on stderr
- `-V, --version`: Print the version and exit

### Exit codes

- `0`: the conversion succeeded; `OK <rows> <cols>` is printed to stdout.
- `1`: the conversion was attempted and failed (an unreadable input, an output path that
  cannot be written, malformed CSV). The reason is printed to stderr.
- `2`: the command line itself was wrong, an invalid `--date-order` included. This is the
  same code clap uses for the usage errors it rejects itself, so a script can tell "you
  typed the command wrong" from "the conversion failed" without parsing stderr.

### Examples

```bash
# Basic conversion
xlsxturbo sales.csv report.xlsx

# With US date format
xlsxturbo sales.csv report.xlsx --date-order us

# With European date format and verbose output
xlsxturbo sales.csv report.xlsx -d eu -v --sheet-name "Q4 Sales"
```
