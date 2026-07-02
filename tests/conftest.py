"""Shared pytest fixtures for xlsxturbo tests."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from pathlib import Path

import pytest

from tests.helpers import get_temp_path


@pytest.fixture
def tmp_xlsx() -> Iterator[str]:
    """Yield a single temporary ``.xlsx`` path, removing it afterward.

    The path is created (and its handle closed) before the test runs and is
    unlinked with ``missing_ok=True`` during teardown, so tests that never
    end up writing the file (e.g. because an exception is raised first) do
    not fail cleanup.
    """
    path = get_temp_path()
    yield path
    Path(path).unlink(missing_ok=True)


@pytest.fixture
def tmp_xlsx_factory() -> Iterator[Callable[[str], str]]:
    """Yield a factory that produces temporary paths.

    Every path returned by the factory is tracked and unlinked (with
    ``missing_ok=True``) during teardown, for tests that need more than one
    temporary file (e.g. an ``.xlsx`` output alongside a derived ``.csv``
    input or a sibling ``.png`` asset). Pass ``suffix`` to control the
    extension; it defaults to ``.xlsx``.
    """
    paths: list[str] = []

    def _make_path(suffix: str = ".xlsx") -> str:
        """Create and track a new temporary path.

        Args:
            suffix: File extension (including the dot) for the temp path.

        Returns:
            The tracked temporary path as a string.
        """
        path = get_temp_path(suffix=suffix)
        paths.append(path)
        return path

    yield _make_path
    for path in paths:
        Path(path).unlink(missing_ok=True)
