"""Tests for :class:`xlsxturbo.options.ExportOptions`.

The bundle is a second spelling of an existing surface, so almost everything here
guards against the two halves disagreeing rather than against the bundle being
wrong on its own. Both populations are derived -- from the real signature and
from what the library actually rejects -- because a hand-written expected list is
the thing that rots, and this module's whole job is to not rot relative to
another one.
"""

from __future__ import annotations

import inspect
import zipfile
from dataclasses import FrozenInstanceError, fields, replace
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import pandas as pd
import pytest
import xlsxturbo
from xlsxturbo.options import ExportOptions
from xlsxturbo.types import SheetOptions, SparklineOptions

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

# `df_to_xlsx` parameters that are not options: the data, where it goes, and the
# name of the single sheet it goes to.
NON_OPTION_PARAMS = frozenset({"df", "output_path", "sheet_name"})

# One valid value per option, used to drive both the round-trip check and the
# empirical per-sheet split below. Values are deliberately minimal -- this module
# tests the plumbing, not what each option renders.
SAMPLE_VALUES: dict[str, Any] = {
    "header": True,
    "autofit": True,
    "freeze_panes": True,
    "column_widths": {0: 30},
    "row_heights": {0: 22},
    "table_style": "Medium9",
    "table_name": "T1",
    "header_format": {"bold": True},
    "column_formats": {"Score": {"bold": True}},
    # A data bar deliberately, because the bundle below also sets `sparklines`:
    # that pair made rust_xlsxwriter emit malformed XML until 0.98.1, and it was
    # the full-bundle check here that found it. Keeping the two together is what
    # would notice a regression.
    "conditional_formats": {"Score": {"type": "data_bar"}},
    "cells": {"E1": "note"},
    "formula_columns": {"Double": "=B{row}*2"},
    "merged_ranges": [("F1:G1", "Merged")],
    "hyperlinks": [("H1", "https://example.com")],
    "comments": {"I1": "a note"},
    "rich_text": {"J1": [("Bold", {"bold": True}), " plain"]},
    "validations": {"Name": {"type": "text_length", "min": 1, "max": 50}},
    "images": {},
    "checkboxes": {"K1": True},
    "textboxes": {"L1": "a textbox"},
    "charts": {"M2": {"type": "bar", "data_range": "Sheet1!$B$2:$B$3"}},
    "sparklines": {"N2": {"range": "Sheet1!A2:B2"}},
    "constant_memory": False,
    "defined_names": {"MyName": "Sheet1!$A$1"},
}


# Typed rather than inferred: the entry points take `dict[str, SparklineOptions]`,
# and an inferred `dict[str, dict[str, str]]` fails type checking at every use.
SPARKLINE_SAMPLE: dict[str, SparklineOptions] = {"N2": {"range": "Sheet1!A2:B2"}}


def _frame() -> pd.DataFrame:
    """Return the small frame every check in this module writes."""
    return pd.DataFrame({"Name": ["a", "b"], "Score": [1, 2]})


def _option_names() -> set[str]:
    """The real option keywords, read from the compiled signature."""
    params = set(inspect.signature(xlsxturbo.df_to_xlsx).parameters)
    assert len(params) > 20, f"signature introspection returned only {len(params)}"
    return params - NON_OPTION_PARAMS


def _field_names() -> set[str]:
    """The names ``ExportOptions`` declares."""
    names = {f.name for f in fields(ExportOptions)}
    assert names, "ExportOptions declares no fields"
    return names


class TestCoverage:
    """The bundle and the keyword surface cannot drift apart."""

    def test_every_option_has_a_field(self) -> None:
        """A new keyword argument without a field fails here.

        This is the guard the phase exists to install. Adding option N+1 to the
        Rust signature and forgetting the dataclass would otherwise leave the two
        spellings quietly unequal, and only the keyword one would work.
        """
        missing = _option_names() - _field_names()
        assert not missing, f"df_to_xlsx option(s) with no ExportOptions field: {sorted(missing)}"

    def test_every_field_is_a_real_option(self) -> None:
        """A field naming nothing real fails here.

        The other direction, and the one that produces a confusing failure rather
        than a missing feature: a typo'd field would be silently accepted by the
        dataclass and then rejected by the extension as an unknown keyword.
        """
        extra = _field_names() - _option_names()
        assert not extra, f"ExportOptions field(s) that are not options: {sorted(extra)}"

    def test_sample_values_cover_every_field(self) -> None:
        """The table driving the other checks is itself complete.

        Without this, adding a field and no sample would silently shrink the
        round-trip and per-sheet checks below to a subset -- they would keep
        passing while covering less.
        """
        missing = _field_names() - set(SAMPLE_VALUES)
        assert not missing, f"no sample value for: {sorted(missing)}"


class TestLowering:
    """``as_kwargs`` and ``as_sheet_options`` produce what the API accepts."""

    def test_unset_options_are_omitted(self) -> None:
        """An untouched bundle lowers to nothing at all."""
        assert ExportOptions().as_kwargs() == {}

    def test_explicit_none_is_kept(self) -> None:
        """``None`` is a value, not an absence.

        Per-sheet, ``table_style=None`` means "no table on this sheet", shadowing
        a workbook default. Collapsing unset and ``None`` would make that
        unexpressible -- so this is a behaviour, not a nicety.
        """
        assert ExportOptions(table_style=None).as_kwargs() == {"table_style": None}

    def test_set_options_survive_lowering(self) -> None:
        """Values pass through unchanged."""
        opts = ExportOptions(freeze_panes=True, header_format={"bold": True})
        assert opts.as_kwargs() == {"freeze_panes": True, "header_format": {"bold": True}}

    def test_sheet_options_drop_exactly_the_rejected_options(
        self, tmp_path: Path
    ) -> None:
        """``as_sheet_options`` drops precisely what a per-sheet dict rejects.

        The set is derived by asking the library, not copied from the module. A
        hardcoded list here would agree with the hardcoded list there forever,
        including when both are wrong -- and this is exactly the kind of split
        that gets one side updated.

        The probe writes into ``tmp_path``. It used to write ``tests/_probe.xlsx``
        inside the checkout, cleaned in a ``finally``: an interrupted run left an
        untracked file behind (``.gitignore`` un-ignores ``tests/*.xlsx``), and two
        concurrent runs collided on the one path.
        """
        probe = tmp_path / "_probe.xlsx"
        rejected: set[str] = set()
        for name, value in sorted(SAMPLE_VALUES.items()):
            try:
                xlsxturbo.dfs_to_xlsx(
                    [(_frame(), "S", {name: value})],  # type: ignore[list-item]
                    probe,
                )
            except xlsxturbo.ConfigurationError as exc:
                if "Unknown sheet option" in str(exc):
                    rejected.add(name)
            except xlsxturbo.XlsxTurboError:
                pass  # a different complaint means the key itself was accepted

        assert rejected, "no option was rejected per-sheet; the probe is not working"
        full, sheet = set(ExportOptions(**SAMPLE_VALUES).as_kwargs()), set(
            ExportOptions(**SAMPLE_VALUES).as_sheet_options()
        )
        assert full - sheet == rejected, (
            f"as_sheet_options drops {sorted(full - sheet)}, "
            f"but a per-sheet dict rejects {sorted(rejected)}"
        )


@pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")
class TestEquivalence:
    """A bundle produces the same workbook as the keywords it lowers to."""

    def test_bundle_and_keywords_produce_identical_bytes(self, tmp_path: Path) -> None:
        """The two spellings are the same call.

        Byte equality over the archive members rather than the file, because a
        zip carries per-entry timestamps that say nothing about the content.
        """
        opts = ExportOptions(
            freeze_panes=True, autofit=True, header_format={"bold": True}, table_style="Medium9"
        )
        via_bundle, via_kwargs = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), via_bundle, **opts.as_kwargs())
        xlsxturbo.df_to_xlsx(
            _frame(),
            via_kwargs,
            freeze_panes=True,
            autofit=True,
            header_format={"bold": True},
            table_style="Medium9",
        )
        with zipfile.ZipFile(via_bundle) as a, zipfile.ZipFile(via_kwargs) as b:
            assert a.namelist() == b.namelist()
            for name in a.namelist():
                assert a.read(name) == b.read(name), f"{name} differs between the two spellings"

    def test_a_bundle_actually_applies_per_sheet(self, tmp_path: Path) -> None:
        """``as_sheet_options`` reaches every sheet it is given to."""
        opts = ExportOptions(freeze_panes=True, header_format={"bold": True})
        out = tmp_path / "multi.xlsx"
        xlsxturbo.dfs_to_xlsx(
            [(_frame(), "Q1", opts.as_sheet_options()), (_frame(), "Q2", opts.as_sheet_options())],
            out,
        )
        wb = load_workbook(out)
        for sheet in ("Q1", "Q2"):
            assert wb[sheet].freeze_panes == "A2"
            assert wb[sheet]["A1"].font.bold is True

    def test_every_option_lowers_to_an_accepted_call(self, tmp_path: Path) -> None:
        """A full bundle is accepted by the entry point, field by field.

        Catches a field whose name matches an option but whose lowering the
        extension will not take -- the failure mode the coverage tests above
        cannot see, because they compare names and never make a call.
        """
        out = tmp_path / "full.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), out, **ExportOptions(**SAMPLE_VALUES).as_kwargs())
        assert active_ws(load_workbook(out))["A1"].value == "Name"


class TestBundleSemantics:
    """The dataclass behaves as the docstring promises."""

    def test_bundle_is_frozen(self) -> None:
        """A shared constant cannot be mutated by one caller for everyone."""
        opts = ExportOptions(freeze_panes=True)
        with pytest.raises(FrozenInstanceError):
            opts.freeze_panes = False  # type: ignore[misc]

    def test_replace_derives_a_variant(self) -> None:
        """``dataclasses.replace`` works, and leaves the original alone."""
        base = ExportOptions(freeze_panes=True, autofit=True)
        derived = replace(base, freeze_panes=False)
        assert derived.as_kwargs() == {"freeze_panes": False, "autofit": True}
        assert base.as_kwargs() == {"freeze_panes": True, "autofit": True}

    def test_merge_takes_only_what_the_override_set(self) -> None:
        """A sparse override does not reset the base's other options.

        This is the property that makes a shared base bundle worth having; a
        naive merge over all fields would silently blank them.
        """
        base = ExportOptions(freeze_panes=True, autofit=True)
        merged = base.merged_with(ExportOptions(autofit=False, table_style="Medium2"))
        assert merged.as_kwargs() == {
            "freeze_panes": True,
            "autofit": False,
            "table_style": "Medium2",
        }

    def test_merge_does_not_mutate_either_side(self) -> None:
        """Both inputs are unchanged by a merge."""
        base, other = ExportOptions(autofit=True), ExportOptions(autofit=False)
        base.merged_with(other)
        assert base.as_kwargs() == {"autofit": True}
        assert other.as_kwargs() == {"autofit": False}

    def test_module_exports_only_the_bundle(self) -> None:
        """``options.__all__`` is the one public name, and it resolves."""
        from xlsxturbo import options as options_module

        assert options_module.__all__ == ["ExportOptions"]
        assert options_module.ExportOptions is ExportOptions

    def test_bundle_is_reachable_from_the_package_root(self) -> None:
        """``xlsxturbo.ExportOptions`` is the same object, and is declared."""
        assert xlsxturbo.ExportOptions is ExportOptions
        assert "ExportOptions" in xlsxturbo.__all__


class TestDataBarBesideSparklines:
    """The pair that once produced a corrupt workbook is written correctly.

    A ``data_bar`` conditional format beside a sparkline made rust_xlsxwriter
    emit unbalanced ``<ext>`` elements -- three opened, two closed -- so Excel
    reported the workbook as damaged. xlsxturbo refused the combination from
    1.0.0 until the upstream fix in rust_xlsxwriter 0.98.1
    (`jmcnamara/rust_xlsxwriter#185 <https://github.com/jmcnamara/rust_xlsxwriter/issues/185>`_).

    The refusal is gone, so what has to be checked now is the output rather than
    the error. Parsing the worksheet XML is the point: the archive wrote fine
    while it was broken, and ``openpyxl`` is not what the failure reached -- Excel
    was. A parser is the cheapest thing that fails on exactly what failed before.
    """

    def _sheet_xml(self, path: Path) -> bytes:
        """Return ``xl/worksheets/sheet1.xml`` from a written workbook.

        Args:
            path: The workbook to read.

        Returns:
            The raw worksheet XML.
        """
        with zipfile.ZipFile(path) as zf:
            return zf.read("xl/worksheets/sheet1.xml")

    def test_the_pair_produces_well_formed_xml(self, tmp_path: Path) -> None:
        """Both features on one sheet, and the result parses.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "pair.xlsx"
        xlsxturbo.df_to_xlsx(
            _frame(),
            out,
            conditional_formats={"Score": {"type": "data_bar"}},
            sparklines=SPARKLINE_SAMPLE,
        )
        xml = self._sheet_xml(out)
        ElementTree.fromstring(xml)  # noqa: S314 - our own output, not untrusted input
        assert xml.count(b"<ext ") == xml.count(b"</ext>"), (
            f"unbalanced <ext>: {xml.count(b'<ext ')} opened, {xml.count(b'</ext>')} closed"
        )

    def test_both_features_actually_reached_the_sheet(self, tmp_path: Path) -> None:
        """The control for the check above.

        A workbook with neither feature applied parses perfectly, so the
        well-formedness assertion would pass if a future change silently dropped
        one of them. This reads both back out of the XML.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "pair.xlsx"
        xlsxturbo.df_to_xlsx(
            _frame(),
            out,
            conditional_formats={"Score": {"type": "data_bar"}},
            sparklines=SPARKLINE_SAMPLE,
        )
        xml = self._sheet_xml(out)
        assert b"dataBar" in xml
        assert b"sparkline" in xml

    def test_the_pair_survives_a_multi_sheet_export(self, tmp_path: Path) -> None:
        """Per-sheet options carry the combination too.

        The two features can arrive from different places -- a workbook-wide
        default and a per-sheet override -- which is where the old guard had to
        look, and is still where a merge mistake would show.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "multi.xlsx"
        xlsxturbo.dfs_to_xlsx(
            [
                (_frame(), "A", {}),
                (_frame(), "B", SheetOptions(sparklines=SPARKLINE_SAMPLE)),
            ],
            out,
            conditional_formats={"Score": {"type": "data_bar"}},
        )
        with zipfile.ZipFile(out) as zf:
            for member in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
                ElementTree.fromstring(zf.read(member))  # noqa: S314 - our own output
