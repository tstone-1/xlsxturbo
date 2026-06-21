"""Tests for the version() function and __version__ attribute."""

from __future__ import annotations

import importlib.metadata

import xlsxturbo


def test_version_returns_semver_string() -> None:
    """version() returns a non-empty dotted version string."""
    v = xlsxturbo.version()
    assert isinstance(v, str)
    assert v.count(".") >= 2  # MAJOR.MINOR.PATCH


def test_version_matches_package_metadata() -> None:
    """version() agrees with the installed package metadata."""
    assert xlsxturbo.version() == importlib.metadata.version("xlsxturbo")


def test_dunder_version_matches_function() -> None:
    """__version__ and version() report the same value."""
    assert xlsxturbo.__version__ == xlsxturbo.version()
