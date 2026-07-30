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
    # Deliberately a colour scale rather than a data bar: a data bar and a
    # sparkline on the same worksheet make rust_xlsxwriter 0.97.0 emit malformed
    # XML -- so xlsxturbo refuses that pair outright, which would make the
    # full-bundle check below fail for a reason having nothing to do with
    # ExportOptions. See TestDataBarSparklineGuard at the bottom of this module.
    "conditional_formats": {"Score": {"type": "2_color_scale"}},
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

    def test_sheet_options_drop_exactly_the_rejected_options(self) -> None:
        """``as_sheet_options`` drops precisely what a per-sheet dict rejects.

        The set is derived by asking the library, not copied from the module. A
        hardcoded list here would agree with the hardcoded list there forever,
        including when both are wrong -- and this is exactly the kind of split
        that gets one side updated.
        """
        rejected: set[str] = set()
        for name, value in sorted(SAMPLE_VALUES.items()):
            try:
                xlsxturbo.dfs_to_xlsx(
                    [(_frame(), "S", {name: value})],  # type: ignore[list-item]
                    Path(__file__).parent / "_probe.xlsx",
                )
            except xlsxturbo.ConfigurationError as exc:
                if "Unknown sheet option" in str(exc):
                    rejected.add(name)
            except xlsxturbo.XlsxTurboError:
                pass  # a different complaint means the key itself was accepted
            finally:
                (Path(__file__).parent / "_probe.xlsx").unlink(missing_ok=True)

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


class TestDataBarSparklineGuard:
    """The one option combination the writer corrupts is refused, not written.

    A ``data_bar`` conditional format beside a sparkline makes rust_xlsxwriter
    0.97.0 emit malformed worksheet XML, so xlsxturbo refuses the pair. The
    upstream defect itself is pinned in ``tests/upstream_defect.rs``, which uses
    rust_xlsxwriter directly -- it has to, because the guard means no Python call
    can produce the corrupt file any more, and something still has to notice the
    day upstream fixes it.
    """

    @pytest.mark.parametrize("spelling", ["data_bar", "databar"])
    def test_the_combination_is_refused(self, spelling: str, tmp_path: Path) -> None:
        """Both accepted spellings of the type are caught.

        ``databar`` is an accepted alias, so a guard matching only ``data_bar``
        would leave half the door open -- and the half nobody tests.

        Args:
            spelling: The conditional-format type name under test.
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.ConfigurationError, match="sparklines"):
            xlsxturbo.df_to_xlsx(
                _frame(),
                tmp_path / "out.xlsx",
                conditional_formats={"Score": {"type": spelling}},
                sparklines=SPARKLINE_SAMPLE,
            )

    def test_nothing_is_written_when_refused(self, tmp_path: Path) -> None:
        """The refusal leaves no file behind.

        The guard runs before any feature is applied precisely so the caller does
        not get a half-built workbook, which would be worse than the corrupt one
        it replaces.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "out.xlsx"
        with pytest.raises(xlsxturbo.ConfigurationError):
            xlsxturbo.df_to_xlsx(
                _frame(),
                out,
                conditional_formats={"Score": {"type": "data_bar"}},
                sparklines=SPARKLINE_SAMPLE,
            )
        assert not out.exists()

    def test_the_message_names_the_sheet_and_the_column(self, tmp_path: Path) -> None:
        """The error says where the problem is and what to do instead.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.ConfigurationError) as caught:
            xlsxturbo.df_to_xlsx(
                _frame(),
                tmp_path / "out.xlsx",
                conditional_formats={"Score": {"type": "data_bar"}},
                sparklines=SPARKLINE_SAMPLE,
            )
        message = str(caught.value)
        assert "Score" in message
        assert "Sheet1" in message
        assert "2_color_scale" in message, "the message should offer a workaround"

    @pytest.mark.parametrize(
        ("label", "kwargs"),
        [
            ("data bar alone", {"conditional_formats": {"Score": {"type": "data_bar"}}}),
            ("sparklines alone", {"sparklines": SPARKLINE_SAMPLE}),
            (
                "colour scale beside sparklines",
                {
                    "conditional_formats": {"Score": {"type": "2_color_scale"}},
                    "sparklines": SPARKLINE_SAMPLE,
                },
            ),
            (
                "icon set beside sparklines",
                {
                    "conditional_formats": {
                        "Score": {"type": "icon_set", "icon_type": "3_arrows"}
                    },
                    "sparklines": SPARKLINE_SAMPLE,
                },
            ),
            (
                "data bar with an empty sparklines dict",
                {
                    "conditional_formats": {"Score": {"type": "data_bar"}},
                    "sparklines": {},
                },
            ),
        ],
    )
    def test_the_guard_is_narrow(self, label: str, kwargs: dict[str, Any], tmp_path: Path) -> None:
        """Everything adjacent to the bad pair still works.

        The expensive failure mode for a guard like this is over-reach: refusing
        a combination that is fine costs users a feature and nobody notices,
        because the only symptom is an error they assume is their fault.

        Args:
            label: Human-readable name of the combination, for the failure message.
            kwargs: The option combination that must still be accepted.
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "ok.xlsx"
        xlsxturbo.df_to_xlsx(_frame(), out, **kwargs)
        assert out.exists(), label

    def test_the_guard_applies_per_sheet(self, tmp_path: Path) -> None:
        """A multi-sheet export is checked sheet by sheet, and names the sheet.

        The two features could arrive from different places -- a workbook-wide
        default and a per-sheet override -- so checking the merged options of each
        sheet is the only correct place to look.

        Args:
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.ConfigurationError, match="sheet 'B'"):
            xlsxturbo.dfs_to_xlsx(
                [
                    (_frame(), "A", {}),
                    (_frame(), "B", SheetOptions(sparklines=SPARKLINE_SAMPLE)),
                ],
                tmp_path / "multi.xlsx",
                conditional_formats={"Score": {"type": "data_bar"}},
            )
