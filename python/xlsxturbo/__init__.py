"""High-performance Excel writer backed by a Rust extension.

This package re-exports the public API implemented in the compiled
``xlsxturbo`` extension module: the conversion functions, the version
helpers, and the exception hierarchy.

Every failure raised by this library is an :class:`XlsxTurboError`. Each
subclass also inherits the builtin exception that the same failure raised
before the hierarchy existed, so ``except ValueError`` and ``except TypeError``
keep working unchanged -- see ``docs/errors.md``.

The option shapes for annotating calls (``HeaderFormat``, ``ChartOptions``,
``ValidationType``, ...) live in ``xlsxturbo.types``, which imports nothing
beyond the standard library.

:class:`ExportOptions` bundles the option keywords into one reusable, typed
object for callers who would rather not spell out a long argument list; every
keyword remains available directly and nothing is deprecated.
"""

from .options import ExportOptions
from .xlsxturbo import (
    ConfigurationError,
    ConfigurationTypeError,
    FileError,
    InputDataError,
    WorkbookValidationError,
    XlsxTurboError,
    __version__,
    csv_to_xlsx,
    df_to_xlsx,
    dfs_to_xlsx,
    version,
)

__all__ = [
    "ConfigurationError",
    "ConfigurationTypeError",
    "ExportOptions",
    "FileError",
    "InputDataError",
    "WorkbookValidationError",
    "XlsxTurboError",
    "__version__",
    "csv_to_xlsx",
    "df_to_xlsx",
    "dfs_to_xlsx",
    "version",
]
