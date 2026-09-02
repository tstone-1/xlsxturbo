"""Tests guarding the parity of the df_to_xlsx / dfs_to_xlsx public API surface."""

from __future__ import annotations

import inspect
import re

import xlsxturbo

# A `key:` at the start of a Google-style ``Args:`` line, i.e. the parameter
# names a docstring documents. Continuation lines are indented further and have
# no such key, so they do not match.
_ARG_KEY = re.compile(r"^ {4}(\w+):")


def _documented_args(doc: str | None) -> list[str]:
    """The parameter names listed under a docstring's ``Args:`` heading.

    Args:
        doc: The function's ``__doc__``.

    Returns:
        The keys in the order the docstring lists them.
    """
    assert doc is not None
    lines = doc.splitlines()
    start = next(i for i, line in enumerate(lines) if line.strip() == "Args:")
    keys: list[str] = []
    for line in lines[start + 1 :]:
        if line.strip() and not line.startswith(" " * 4):
            break
        if line.strip() in {"Returns:", "Raises:", "Example:"}:
            break
        match = _ARG_KEY.match(line)
        if match:
            keys.append(match.group(1))
    return keys


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


def test_df_to_xlsx_dfs_to_xlsx_document_the_same_options() -> None:
    """The twin docstrings must list the same options, in the same order.

    The test above pins the *signatures*; nothing pinned the ~110-line
    ``Args:`` block each function carries, and they had drifted in three
    places -- ``table_name`` described by its sanitizer rules on one and by the
    workbook uniqueness rule on the other, and the ``cells`` entry missing the
    sentence saying when those writes happen. A docstring is what ``help()``
    shows, so a caller reading the multi-sheet one simply did not learn things
    the single-sheet one taught.

    Deliberately shallow: it compares the key lists and not the prose, because
    the prose legitimately differs (``df`` versus ``sheets``, per-sheet
    overrides) and a text comparison would be a guard nobody could keep green.
    Order is included because these are two hand-maintained lists and a
    reordering is how the next divergence starts.
    """
    single = _documented_args(xlsxturbo.df_to_xlsx.__doc__)
    multi = _documented_args(xlsxturbo.dfs_to_xlsx.__doc__)

    assert len(single) > 20, f"Args: parsing found only {single}"
    assert single[0] == "df"
    assert multi[0] == "sheets"
    # `sheet_name` is the single-sheet path's own parameter; every other key
    # after the first must match one for one.
    assert [key for key in single[1:] if key != "sheet_name"] == multi[1:]
