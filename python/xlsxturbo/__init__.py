"""High-performance Excel writer backed by a Rust extension.

This package re-exports the public API implemented in the compiled
``xlsxturbo`` extension module.
"""

from .xlsxturbo import __version__, csv_to_xlsx, df_to_xlsx, dfs_to_xlsx, version

__all__ = ["__version__", "csv_to_xlsx", "df_to_xlsx", "dfs_to_xlsx", "version"]
