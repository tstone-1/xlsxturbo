# Re-export from the Rust extension
from .fastxlsx import csv_to_xlsx, version, __version__

__all__ = ["csv_to_xlsx", "version", "__version__"]
