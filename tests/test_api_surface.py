"""Tests guarding the parity of the df_to_xlsx / dfs_to_xlsx public API surface."""

from __future__ import annotations

import inspect

import xlsxturbo


def test_inspect_signature_works_on_compiled_functions() -> None:
    """`inspect.signature` resolves real parameter names on the pyo3-compiled functions.

    pyo3 exposes a `__text_signature__` on wrapped functions, which
    `inspect.signature` knows how to parse; this pins that assumption so a
    future pyo3/build change that drops it fails loudly here rather than only
    inside `test_df_to_xlsx_dfs_to_xlsx_parameter_parity` below.
    """
    sig = inspect.signature(xlsxturbo.df_to_xlsx)
    assert "df" in sig.parameters
    assert "output_path" in sig.parameters
    assert len(sig.parameters) > 5


def test_df_to_xlsx_dfs_to_xlsx_parameter_parity() -> None:
    """Every write option on df_to_xlsx must also exist on dfs_to_xlsx, and vice versa.

    df_to_xlsx and dfs_to_xlsx are two separate pyo3-exposed functions with
    duplicated option lists; a parameter added to one and forgotten on the
    other is a silent feature gap (an option works on single-sheet writes but
    is rejected, or simply missing, on the multi-sheet path). The only
    expected differences are structural, not feature options: df_to_xlsx
    takes a single 'df' plus a top-level 'sheet_name', while dfs_to_xlsx takes
    a 'sheets' list of (df, sheet_name[, options]) tuples instead.
    """
    single_params = set(inspect.signature(xlsxturbo.df_to_xlsx).parameters.keys())
    multi_params = set(inspect.signature(xlsxturbo.dfs_to_xlsx).parameters.keys())

    known_single_only = {"df", "sheet_name"}
    known_multi_only = {"sheets"}

    assert single_params - known_single_only == multi_params - known_multi_only
