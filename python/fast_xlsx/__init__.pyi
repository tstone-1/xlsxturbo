"""Type stubs for fast_xlsx"""

def csv_to_xlsx(
    input_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
) -> tuple[int, int]:
    """
    Convert a CSV file to XLSX format with automatic type detection.

    Args:
        input_path: Path to the input CSV file
        output_path: Path for the output XLSX file
        sheet_name: Name of the worksheet (default: "Sheet1")

    Returns:
        Tuple of (rows, columns) written to the Excel file

    Raises:
        ValueError: If the conversion fails

    Example:
        >>> import fast_xlsx
        >>> rows, cols = fast_xlsx.csv_to_xlsx("data.csv", "output.xlsx")
        >>> print(f"Converted {rows} rows and {cols} columns")
    """
    ...

def version() -> str:
    """Get the version of the fast_xlsx library."""
    ...

__version__: str
