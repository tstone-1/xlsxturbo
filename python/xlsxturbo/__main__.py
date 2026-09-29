"""``python -m xlsxturbo``: the same command as the ``xlsxturbo`` console script."""

import sys

from .cli import main

sys.exit(main())
