"""Public type stubs for the xlsxturbo package.

The runtime surface of this package is the compiled extension re-exported by
``__init__.py``: the conversion functions plus ``version`` / ``__version__``.
This stub mirrors exactly that surface, so a type checker never reports an
import as valid that would raise ``ImportError`` at runtime.

The option ``TypedDict`` / ``Literal`` helpers (``SparklineOptions``,
``ChartOptions``, ``ValidationType``, ...) are stub-only types with no runtime
object. Import them from the ``xlsxturbo.xlsxturbo`` submodule inside a
``TYPE_CHECKING`` block when you want to annotate option dicts; the full type
surface lives in ``xlsxturbo.pyi``.
"""

from .xlsxturbo import (
    __version__ as __version__,
    csv_to_xlsx as csv_to_xlsx,
    df_to_xlsx as df_to_xlsx,
    dfs_to_xlsx as dfs_to_xlsx,
    version as version,
)

__all__ = [
    "__version__",
    "csv_to_xlsx",
    "df_to_xlsx",
    "dfs_to_xlsx",
    "version",
]
