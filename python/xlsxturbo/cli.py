"""The ``xlsxturbo`` command: CSV to XLSX from the shell.

Installed by ``pip install xlsxturbo`` as a console script, and runnable as
``python -m xlsxturbo``. It is a thin wrapper over :func:`xlsxturbo.csv_to_xlsx`
and mirrors the Rust binary in ``src/main.rs`` (``cargo build --release``) flag
for flag: the same options, the same ``OK <rows> <cols>`` line on stdout, and the
same exit codes. ``tests/test_cli.py`` compares the two option lists, because the
Rust binary is the one that exists without a Python install and the two must not
drift apart.

Usage::

    xlsxturbo sales.csv report.xlsx --date-order us --sheet-name "Q4 Sales"
"""

from __future__ import annotations

import argparse
import sys
import time
from typing import get_args

from .types import DateOrder
from .xlsxturbo import XlsxTurboError, csv_to_xlsx, version

__all__ = ["EXIT_FAILURE", "EXIT_USAGE", "main"]

EXIT_FAILURE = 1
"""The conversion was attempted and failed: unreadable input, unwritable output, malformed CSV."""

EXIT_USAGE = 2
"""The command line itself was wrong. argparse exits 2 for the errors it catches, too."""

DATE_ORDERS = get_args(DateOrder)


def _date_order(value: str) -> str:
    """Accept a date order in any case, as ``csv_to_xlsx`` does.

    Args:
        value: The ``--date-order`` argument.

    Returns:
        The value, lower-cased.

    Raises:
        argparse.ArgumentTypeError: For a value ``csv_to_xlsx`` would refuse, so
            it is reported as a usage error (exit 2) rather than a failed
            conversion (exit 1).
    """
    lowered = value.lower()
    if lowered not in DATE_ORDERS:
        raise argparse.ArgumentTypeError(
            f"invalid date order {value!r}; valid values: {', '.join(DATE_ORDERS)}"
        )
    return lowered


def build_parser() -> argparse.ArgumentParser:
    """Build the argument parser.

    Returns:
        The parser for the ``xlsxturbo`` command.
    """
    parser = argparse.ArgumentParser(
        prog="xlsxturbo",
        description="Fast CSV to XLSX converter with automatic type detection. "
        "Numbers, booleans, dates and ISO 8601 datetimes become native Excel values; "
        "NaN and Inf become empty cells; everything else is text.",
    )
    parser.add_argument("input", help="Input CSV file path")
    parser.add_argument("output", help="Output XLSX file path")
    parser.add_argument("-s", "--sheet-name", default="Sheet1", help='Sheet name (default: "Sheet1")')
    parser.add_argument(
        "-d",
        "--date-order",
        default="auto",
        type=_date_order,
        metavar="ORDER",
        help="Date order for ambiguous dates like 01-02-2024. auto: ISO, then European, "
        "then US; mdy/us: 01-02-2024 = January 2; dmy/eu/european: 01-02-2024 = February 1 "
        "(default: auto)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Show progress information")
    parser.add_argument(
        "-p",
        "--parallel",
        action="store_true",
        help="Use multi-core parallel processing (faster for large files, uses more memory)",
    )
    parser.add_argument("-V", "--version", action="version", version=f"xlsxturbo {version()}")
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the command.

    Args:
        argv: The arguments after the program name; ``sys.argv[1:]`` when omitted.

    Returns:
        The exit code: 0 on success, :data:`EXIT_FAILURE` when the conversion fails.
        A usage error exits with :data:`EXIT_USAGE` from inside argparse.
    """
    args = build_parser().parse_args(argv)

    if args.verbose:
        print("xlsxturbo - CSV to XLSX converter", file=sys.stderr)
        print(f"Input:  {args.input}", file=sys.stderr)
        print(f"Output: {args.output}", file=sys.stderr)
        print(f"Sheet:  {args.sheet_name}", file=sys.stderr)
        print(f"Dates:  {args.date_order}", file=sys.stderr)
        print(f"Parallel: {args.parallel}", file=sys.stderr)

    start = time.perf_counter()
    try:
        rows, cols = csv_to_xlsx(
            args.input,
            args.output,
            sheet_name=args.sheet_name,
            parallel=args.parallel,
            date_order=args.date_order,
        )
    except XlsxTurboError as error:
        print(f"Error: {error}", file=sys.stderr)
        return EXIT_FAILURE

    if args.verbose:
        secs = time.perf_counter() - start
        rate = f"{rows / secs:.0f}" if secs > 0 else "instant"
        print(f"Converted {rows} rows x {cols} cols in {secs:.2f}s ({rate} rows/sec)", file=sys.stderr)
    print(f"OK {rows} {cols}")
    return 0
