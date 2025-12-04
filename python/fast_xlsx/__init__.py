# Re-export from the Rust extension
from .fast_xlsx import csv_to_xlsx, version, __version__

__all__ = ["csv_to_xlsx", "version", "__version__"]
