# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2024-12-04

### Added
- Initial release
- Python bindings via PyO3
- `csv_to_xlsx()` function for converting CSV files to Excel format
- Automatic type detection from CSV strings:
  - Integers and floats → Excel numbers
  - Booleans (`true`/`false`, case-insensitive) → Excel booleans
  - Dates (YYYY-MM-DD, DD/MM/YYYY, etc.) → Excel dates with formatting
  - Datetimes (ISO 8601) → Excel datetimes with formatting
  - NaN/Inf → Empty cells
  - Empty strings → Empty cells
- CLI tool for command-line usage
- Support for custom sheet names
- Verbose mode for progress reporting

[0.1.0]: https://github.com/tstone-1/xlsxturbo/releases/tag/v0.1.0
