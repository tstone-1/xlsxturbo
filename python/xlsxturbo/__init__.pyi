"""Public type stubs for the xlsxturbo package.

The runtime surface of this package is the compiled extension re-exported by
``__init__.py``: the conversion functions, ``version`` / ``__version__``, and
the exception hierarchy. This stub mirrors exactly that surface, so a type
checker never reports an import as valid that would raise ``ImportError`` at
runtime.

The option ``TypedDict`` / ``Literal`` helpers (``SparklineOptions``,
``ChartOptions``, ``ValidationType``, ...) live in ``xlsxturbo.types``. They are
real runtime objects, so no ``TYPE_CHECKING`` guard is needed::

    from xlsxturbo.types import HeaderFormat

They are deliberately *not* re-exported here, to keep the package's top level to
what the compiled extension provides. :class:`ExportOptions` is the exception: it
is constructed in user code rather than used as an annotation, so it is exported
at the top level where it is discoverable.
"""

from .options import ExportOptions as ExportOptions
from .xlsxturbo import (
    ConfigurationError as ConfigurationError,
    ConfigurationTypeError as ConfigurationTypeError,
    FileError as FileError,
    InputDataError as InputDataError,
    OptionError as OptionError,
    WorkbookValidationError as WorkbookValidationError,
    XlsxTurboError as XlsxTurboError,
    __version__ as __version__,
    csv_to_xlsx as csv_to_xlsx,
    df_to_xlsx as df_to_xlsx,
    dfs_to_xlsx as dfs_to_xlsx,
    version as version,
)

__all__ = [
    "ConfigurationError",
    "ConfigurationTypeError",
    "ExportOptions",
    "FileError",
    "InputDataError",
    "OptionError",
    "WorkbookValidationError",
    "XlsxTurboError",
    "__version__",
    "csv_to_xlsx",
    "df_to_xlsx",
    "dfs_to_xlsx",
    "version",
]
