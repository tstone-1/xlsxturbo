"""Tests for the public exception hierarchy.

Three properties, in order of how much they are worth:

1. **Every class is reachable.** A class no call site raises is dead public API
   that can never be removed. This is not hypothetical -- the planned
   ``UnsupportedFeatureError`` was dropped for exactly this reason (a
   ``constant_memory`` conflict is a warning, not an error), and
   ``InputDataError`` was unraisable through ``df_to_xlsx`` on the first
   attempt, because frame detection happens deep inside the write pipeline and
   surfaced as ``ConfigurationError``. ``TestReachability`` is what caught it.
2. **Every class still is its pre-0.19 builtin.** This is the proof that the
   hierarchy is additive rather than breaking, and it is deliberately asserted
   per class rather than trusted from the class definitions. The 93 existing
   ``pytest.raises(ValueError|TypeError)`` assertions elsewhere in this suite
   cover the same invariant from the other direction, on real call paths.
3. **The shape is what is documented.** ``docs/errors.md`` and the type stubs
   both state the bases; a drift between those and the runtime classes would
   otherwise only show up as a user's broken ``except`` clause.

See ``docs/roadmap-1.0.md`` decision D6 for why the hierarchy has these five
classes and not the ones originally planned.
"""

from __future__ import annotations

import errno
import inspect
import pickle
import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import (
    HAS_OPENPYXL,
    REPO_ROOT,
    active_ws,
    load_workbook,
    repo_checkout_available,
)

# The public names, and the builtin each one must remain an instance of. A class
# appearing here with the wrong builtin is a breaking change for anyone whose
# `except` clause names that builtin.
EXPECTED_BASES: dict[str, tuple[type[BaseException], ...]] = {
    "XlsxTurboError": (Exception,),
    "OptionError": (Exception,),
    "ConfigurationError": (Exception, ValueError),
    "ConfigurationTypeError": (Exception, TypeError),
    "InputDataError": (Exception, ValueError),
    "FileError": (Exception, OSError, ValueError),
    "WorkbookValidationError": (Exception, ValueError),
}

# Builtins a class must NOT be, so that "inherits everything" cannot pass the
# table above. Without this, `ConfigurationError(XlsxTurboError, ValueError,
# TypeError)` -- a shape that was considered and rejected -- would satisfy every
# other assertion in this file.
FORBIDDEN_BASES: dict[str, tuple[type[BaseException], ...]] = {
    "XlsxTurboError": (ValueError, TypeError, OSError),
    # `OptionError` deliberately adds no builtin of its own. If it took one, it
    # would push that builtin onto both of its children -- so a
    # `ConfigurationTypeError` would silently become a `ValueError` too, and the
    # value/type split the hierarchy exists to express would stop meaning
    # anything to `except`.
    "OptionError": (ValueError, TypeError, OSError),
    "ConfigurationError": (TypeError, OSError),
    "ConfigurationTypeError": (ValueError, OSError),
    "InputDataError": (TypeError, OSError),
    "FileError": (TypeError,),
    "WorkbookValidationError": (TypeError, OSError),
}


def _frame() -> pd.DataFrame:
    """Return the smallest DataFrame that still produces a table."""
    return pd.DataFrame({"a": [1, 2]})


def _trigger_configuration(out: Path) -> None:
    """Pass an unknown key inside an option dict."""
    xlsxturbo.df_to_xlsx(_frame(), out / "o.xlsx", header_format={"nope": 1})  # type: ignore[typeddict-unknown-key]


def _trigger_configuration_type(out: Path) -> None:
    """Pass a list where an option requires a dict."""
    xlsxturbo.df_to_xlsx(_frame(), out / "o.xlsx", column_widths=[1, 2])  # type: ignore[arg-type]


def _trigger_input_data(out: Path) -> None:
    """Pass something that is not a DataFrame at all."""
    xlsxturbo.df_to_xlsx([1, 2, 3], out / "o.xlsx")  # type: ignore[arg-type]


def _trigger_file(out: Path) -> None:
    """Write into a directory that does not exist."""
    xlsxturbo.df_to_xlsx(_frame(), out / "absent-dir" / "o.xlsx")


def _trigger_workbook_validation(out: Path) -> None:
    """Give two sheets the same Excel table name."""
    df = _frame()
    xlsxturbo.dfs_to_xlsx(
        [(df, "A"), (df, "B")],
        out / "o.xlsx",
        table_style="Medium2",
        table_name="T",
    )


# One real call per concrete class.
TRIGGERS: dict[str, Callable[[Path], None]] = {
    "ConfigurationError": _trigger_configuration,
    "ConfigurationTypeError": _trigger_configuration_type,
    "InputDataError": _trigger_input_data,
    "FileError": _trigger_file,
    "WorkbookValidationError": _trigger_workbook_validation,
}

# Classes that exist to be *caught*, never raised, mapped to exactly the
# triggered classes each one must catch.
#
# "Abstract" is otherwise an excuse: the rule that every exported class needs a
# working trigger is what kept a dead class from shipping in 0.19, and a class
# that simply declares itself abstract would walk straight past it. So each entry
# here carries its own obligation, checked in both directions -- an abstract
# class that stopped catching one of its children fails, and so does one that
# started catching something else.
ABSTRACT: dict[str, frozenset[str]] = {
    "XlsxTurboError": frozenset(TRIGGERS),
    "OptionError": frozenset(
        {"ConfigurationError", "ConfigurationTypeError", "WorkbookValidationError"}
    ),
}


class TestReachability:
    """Every concrete class is raised by a real call, and by the right one."""

    @pytest.mark.parametrize("name", sorted(TRIGGERS))
    def test_class_is_raised_by_a_real_call(self, name: str, tmp_path: Path) -> None:
        """The trigger raises exactly this class -- not a base, not a sibling."""
        expected = getattr(xlsxturbo, name)
        with pytest.raises(xlsxturbo.XlsxTurboError) as caught:
            TRIGGERS[name](tmp_path)
        # `type(...) is` rather than `isinstance`: a subclass slipping through
        # would mean the message is being classified more narrowly than intended,
        # and `WorkbookValidationError` is a subclass of `ConfigurationError`, so
        # `isinstance` cannot tell those two apart at all.
        assert type(caught.value) is expected, (
            f"{name} trigger raised {type(caught.value).__name__}: {caught.value}"
        )

    def test_unsupported_frame_in_a_sheet_names_the_sheet(self, tmp_path: Path) -> None:
        """``dfs_to_xlsx`` says *which* sheet held the unusable frame.

        Covers the second boundary check (`src/lib.rs`, the `dfs_to_xlsx` loop),
        which the `TRIGGERS` table above cannot reach -- that table maps one
        trigger per class and its `InputDataError` entry goes through
        `df_to_xlsx`. Without this, dropping the sheet name from the message, or
        the whole per-sheet check, breaks no test.

        The prefix is not cosmetic: on a multi-sheet export it is the only thing
        identifying which frame was wrong.
        """
        good = _frame()
        with pytest.raises(xlsxturbo.InputDataError, match=r"sheet 'B':"):
            xlsxturbo.dfs_to_xlsx([(good, "A"), ({}, "B")], tmp_path / "o.xlsx")  # type: ignore[list-item]

    def test_every_concrete_class_has_a_trigger(self) -> None:
        """No class is exported without a call path in this file that raises it.

        This is the guard that keeps a future class from being added to the
        hierarchy, exported, documented, and never raised.
        """
        exported = {
            name
            for name in xlsxturbo.__all__
            if isinstance(getattr(xlsxturbo, name), type)
            and issubclass(getattr(xlsxturbo, name), xlsxturbo.XlsxTurboError)
        }
        concrete = exported - set(ABSTRACT)
        assert concrete == set(TRIGGERS), (
            f"exported without a trigger: {sorted(concrete - set(TRIGGERS))}; "
            f"triggered but not exported: {sorted(set(TRIGGERS) - concrete)}"
        )

    @pytest.mark.parametrize("name", sorted(ABSTRACT))
    def test_abstract_class_catches_exactly_its_documented_children(self, name: str) -> None:
        """An abstract class earns its place by what it catches.

        Checked against the *triggered* classes rather than against the module,
        so a class nothing can raise contributes nothing here -- which is what
        stops ``ABSTRACT`` becoming a way around
        :meth:`test_every_concrete_class_has_a_trigger`.
        """
        abstract = getattr(xlsxturbo, name)
        catches = {t for t in TRIGGERS if issubclass(getattr(xlsxturbo, t), abstract)}
        assert catches == ABSTRACT[name], (
            f"{name} catches {sorted(catches)}, documented as {sorted(ABSTRACT[name])}"
        )
        assert catches, f"{name} catches nothing that can actually be raised"

    @pytest.mark.parametrize("name", sorted(ABSTRACT))
    def test_abstract_class_catches_a_real_failure(self, name: str, tmp_path: Path) -> None:
        """...and by catching one, at runtime, not by ``issubclass`` alone.

        ``issubclass`` answers a question about the class objects. This answers
        the question a caller actually asks -- whether writing ``except
        OptionError`` around a call catches the failure it is meant to.
        """
        abstract = getattr(xlsxturbo, name)
        for child in sorted(ABSTRACT[name]):
            with pytest.raises(abstract):
                TRIGGERS[child](tmp_path)


class TestLegacyBuiltins:
    """The hierarchy is additive: no path changed which builtin it raises."""

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_class_is_its_documented_builtins(self, name: str) -> None:
        """The class is a subclass of each builtin listed for it."""
        cls = getattr(xlsxturbo, name)
        for base in EXPECTED_BASES[name]:
            assert issubclass(cls, base), f"{name} is no longer a {base.__name__}"

    @pytest.mark.parametrize("name", sorted(FORBIDDEN_BASES))
    def test_class_is_not_a_builtin_it_should_not_be(self, name: str) -> None:
        """The class is *not* a subclass of the builtins it must stay clear of."""
        cls = getattr(xlsxturbo, name)
        for base in FORBIDDEN_BASES[name]:
            assert not issubclass(cls, base), (
                f"{name} became a {base.__name__}, which widens what "
                f"`except {base.__name__}` catches"
            )

    @pytest.mark.parametrize("name", sorted(TRIGGERS))
    def test_raised_instance_is_catchable_as_its_builtins(self, name: str, tmp_path: Path) -> None:
        """A real raise is catchable by each legacy builtin, not just declared so.

        Distinct from the ``issubclass`` checks above: those read the class
        object, this one goes through an actual ``except`` clause on an actual
        failure, which is what user code does.
        """
        for base in EXPECTED_BASES[name]:
            with pytest.raises(base):
                TRIGGERS[name](tmp_path)


class TestHierarchyShape:
    """The runtime classes match what the stubs and docs claim."""

    def test_all_classes_derive_from_the_base(self) -> None:
        """Every exception class in the hierarchy is an ``XlsxTurboError``."""
        for name in EXPECTED_BASES:
            assert issubclass(getattr(xlsxturbo, name), xlsxturbo.XlsxTurboError)

    def test_base_catches_everything_the_library_raises(self, tmp_path: Path) -> None:
        """``except XlsxTurboError`` is sufficient for every trigger.

        Written as try/except rather than ``pytest.raises`` so the failure names
        the trigger: a bare "DID NOT RAISE" from inside a loop over five triggers
        does not say which one stopped raising.
        """
        for name, trigger in sorted(TRIGGERS.items()):
            try:
                trigger(tmp_path)
            except xlsxturbo.XlsxTurboError:
                continue
            pytest.fail(f"{name} trigger raised nothing catchable as XlsxTurboError")

    def test_workbook_validation_is_a_configuration_error(self) -> None:
        """The one intentional parent/child pair inside the hierarchy holds."""
        assert issubclass(xlsxturbo.WorkbookValidationError, xlsxturbo.ConfigurationError)

    def test_file_error_carries_the_os_error_number(self, tmp_path: Path) -> None:
        """``errno`` is populated, so ``FileError`` is an honest ``OSError``.

        Writing into a directory that does not exist is ``ENOENT`` whichever
        platform reports it, which is what makes this assertable rather than
        merely "some number".
        """
        with pytest.raises(xlsxturbo.FileError) as caught:
            _trigger_file(tmp_path)
        assert caught.value.errno == errno.ENOENT

    def test_file_error_leaves_strerror_and_filename_unset(self, tmp_path: Path) -> None:
        """...and the other two ``OSError`` fields stay ``None``, deliberately.

        Not an oversight and not laziness: ``OSError.__str__`` switches to the
        ``[Errno n] strerror: 'filename'`` form the moment ``filename`` is set,
        **discarding the message entirely**. This library's message says more
        than that form can, and already contains the path.

        Pinned because populating them looks like an obvious improvement, and
        the resulting message would be strictly worse.
        """
        with pytest.raises(xlsxturbo.FileError) as caught:
            _trigger_file(tmp_path)
        assert caught.value.strerror is None
        assert caught.value.filename is None
        assert str(caught.value).startswith("Failed to save workbook to")

    def test_a_readable_directory_still_produces_no_errno_confusion(
        self, tmp_path: Path
    ) -> None:
        """The control: a successful export raises nothing at all.

        Without it, an implementation that raised ``FileError(ENOENT)``
        unconditionally would satisfy both assertions above.
        """
        xlsxturbo.df_to_xlsx(_frame(), tmp_path / "fine.xlsx")
        assert (tmp_path / "fine.xlsx").exists()

    def test_csv_input_failure_also_carries_an_errno(self, tmp_path: Path) -> None:
        """The read path is a different call site and needs its own case."""
        with pytest.raises(xlsxturbo.FileError) as caught:
            xlsxturbo.csv_to_xlsx(tmp_path / "absent.csv", tmp_path / "o.xlsx")
        assert caught.value.errno == errno.ENOENT
        assert "Failed to open input file" in str(caught.value)


class TestExports:
    """The names are importable from the package, and are one object each."""

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_name_is_exported_and_listed(self, name: str) -> None:
        """The class is reachable as ``xlsxturbo.<name>`` and is in ``__all__``."""
        assert hasattr(xlsxturbo, name)
        assert name in xlsxturbo.__all__

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_package_and_extension_expose_the_same_object(self, name: str) -> None:
        """The re-export is the same class object, so ``isinstance`` agrees.

        Two distinct class objects with the same name would make
        ``except xlsxturbo.FileError`` silently fail to catch an error raised as
        ``xlsxturbo.xlsxturbo.FileError``.
        """
        assert getattr(xlsxturbo, name) is getattr(xlsxturbo.xlsxturbo, name)

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_module_and_qualname_read_as_the_public_path(self, name: str) -> None:
        """``repr()`` names the path users import from."""
        cls = getattr(xlsxturbo, name)
        assert cls.__module__ == "xlsxturbo"
        assert cls.__qualname__ == name

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_instances_survive_a_pickle_round_trip(self, name: str) -> None:
        """The class pickles to the same class object, not a lookalike.

        ``__module__``/``__qualname__`` above are the *mechanism* that makes this
        work, not the property itself -- asserting the mechanism and calling it
        picklability is asserting a proxy. Pickle is how an exception crosses a
        process boundary, so a multiprocessing pool that re-raises one depends on
        this exactly.
        """
        cls = getattr(xlsxturbo, name)
        restored = pickle.loads(pickle.dumps(cls("boom")))  # noqa: S301 - our own object
        assert type(restored) is cls
        assert str(restored) == "boom"

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_class_has_a_docstring(self, name: str) -> None:
        """Each class documents itself; ``help()`` is the first thing users try."""
        doc = getattr(xlsxturbo, name).__doc__
        assert doc is not None
        assert len(doc.strip()) > 40


# Options whose value the PyO3 signature converts before xlsxturbo sees it. A
# wrong type here raises a plain `TypeError` by design -- see "What is not in the
# hierarchy" in docs/errors.md. They are listed so the completeness check below
# can subtract them and still fail on a genuinely new option.
SIGNATURE_CONVERTED = frozenset(
    {
        "df",
        "output_path",
        "sheet_name",
        "header",
        "autofit",
        "table_style",
        "freeze_panes",
        "table_name",
        "constant_memory",
        # Typed in the PyO3 signature as `HashMap<u32, f64>` / `HashMap<String,
        # String>`, so a wrong inner type is converted -- and rejected -- by the
        # binding before any xlsxturbo code runs. Verified, not assumed: they are
        # the only two dict options declared with a concrete Rust map type rather
        # than `&Bound<PyAny>`, which is why `column_widths` classifies and
        # `row_heights` does not. Recorded as a 1.0 consistency question in
        # docs/roadmap-1.0.md D8.
        "row_heights",
        "defined_names",
    }
)

# One deliberately wrong nested key or value per extractor family. Each must
# surface inside the hierarchy rather than as PyO3's own TypeError.
NESTED_TYPE_PROBES: dict[str, dict[str, Any]] = {
    "column_widths": {"column_widths": {0: "wide"}},
    "header_format": {"header_format": {1: True}},
    "column_formats": {"column_formats": {1: {"bold": True}}},
    "conditional_formats": {"conditional_formats": {1: {"type": "data_bar"}}},
    "formula_columns": {"formula_columns": {"X": 123}},
    "merged_ranges": {"merged_ranges": [(1, "text")]},
    "hyperlinks": {"hyperlinks": [(1, "https://example.com")]},
    "comments": {"comments": {"D1": 123}},
    "validations": {"validations": {1: {"type": "list"}}},
    "rich_text": {"rich_text": {1: ["a"]}},
    "images": {"images": {"D1": 123}},
    "checkboxes": {"checkboxes": {"D1": {"checked": "yes"}}},
    "textboxes": {"textboxes": {"D1": {"text": 123}}},
    "charts": {"charts": {1: {"type": "bar"}}},
    "sparklines": {"sparklines": {1: {"range": "Sheet1!A1:B1"}}},
    "cells": {"cells": {1: "value"}},
}


class TestNestedExtractionStaysInTheHierarchy:
    """A wrong type *inside* an option must still be an ``XlsxTurboError``.

    This is the guard that ``TestReachability`` structurally cannot be. That
    class exercises one trigger per exported class, which proves those five
    paths reach the hierarchy and says nothing about the population of
    extraction paths -- a consistency check proves agreement, never
    completeness.

    It was not hypothetical. Until 0.19.1 every custom extractor used a bare
    ``extract()?`` for its nested keys and values, so a non-string
    ``column_formats`` key, a non-string ``formula_columns`` value, a bad
    ``merged_ranges`` tuple element and a dozen more all propagated PyO3's plain
    ``TypeError``. ``docs/errors.md`` promised the opposite, and the whole suite
    was green. An independent review found it; nothing in this file could have.
    """

    @pytest.mark.parametrize("option", sorted(NESTED_TYPE_PROBES))
    def test_bad_nested_value_raises_in_the_hierarchy(
        self, option: str, tmp_path: Path
    ) -> None:
        """Each extractor family classifies its own nested conversion failures.

        Args:
            option: The option under test, used to look up its probe.
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.XlsxTurboError):
            xlsxturbo.df_to_xlsx(
                _frame(), tmp_path / "out.xlsx", **NESTED_TYPE_PROBES[option]
            )

    @pytest.mark.parametrize("option", sorted(NESTED_TYPE_PROBES))
    def test_the_message_names_the_option(self, option: str, tmp_path: Path) -> None:
        """The error says which option was wrong, not just that something was.

        A classified error carrying PyO3's bare ``'int' object cannot be
        converted`` would pass the test above while telling the caller nothing
        about where to look.

        Args:
            option: The option under test, used to look up its probe.
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.XlsxTurboError) as caught:
            xlsxturbo.df_to_xlsx(
                _frame(), tmp_path / "out.xlsx", **NESTED_TYPE_PROBES[option]
            )
        assert option in str(caught.value), (
            f"message does not name the option: {caught.value!r}"
        )

    def test_every_non_scalar_option_has_a_probe(self) -> None:
        """The probe table covers every option that has its own extractor.

        Derived from the real signature rather than hand-listed, so option N+1
        cannot be added with an unclassified extractor and no failing test. The
        same mechanism as ``tests/test_option_coverage.py``, aimed at error
        classification instead of at whether the option does anything.
        """
        params = set(inspect.signature(xlsxturbo.df_to_xlsx).parameters)
        assert len(params) > 20, f"signature introspection returned only {len(params)}"

        needs_probe = params - SIGNATURE_CONVERTED
        missing = needs_probe - set(NESTED_TYPE_PROBES)
        extra = set(NESTED_TYPE_PROBES) - needs_probe
        assert not missing, f"option(s) with no nested-type probe: {sorted(missing)}"
        assert not extra, f"probes for nonexistent option(s): {sorted(extra)}"

    def test_signature_converted_list_is_not_hiding_a_real_extractor(self) -> None:
        """Every name excused above really is converted by the binding.

        Otherwise the exclusion list becomes a place to quietly park an option
        that fails the guard -- the exclusion file that swallows its own
        coverage gap.
        """
        stale = SIGNATURE_CONVERTED - set(inspect.signature(xlsxturbo.df_to_xlsx).parameters)
        assert not stale, f"SIGNATURE_CONVERTED names non-parameters: {sorted(stale)}"


# One deliberately wrong value per optional sub-key of a ``cells`` entry dict,
# with the Python type name the message has to name. The probe table above
# reaches one layer less far: it drives a wrong nested *key* per extractor
# family, and a `cells` entry's own option dict is a level below that.
CELL_VALUE_PROBES: dict[str, tuple[Any, str]] = {
    "num_format": (42, "int"),
    "align_horizontal": (42, "int"),
    "align_vertical": (["top"], "list"),
    "wrap_text": ("yes", "str"),
}


class TestCellsDictValuesStayInTheHierarchy:
    """The ``cells`` dict form classifies its own sub-values, and takes ``None``.

    ``cells={"A1": {"value": 1, "num_format": 42}}`` raised PyO3's own
    ``TypeError`` -- outside the hierarchy, which ``docs/errors.md`` promises
    covers every failure the library raises -- and ``num_format=None`` raised
    that same error rather than reading as "not given". These keys are the ones
    a caller writes out by name, so a ``None`` from their own conditional
    expression has to be as passable as omitting the key.
    """

    @pytest.mark.parametrize("key", sorted(CELL_VALUE_PROBES))
    def test_a_wrong_typed_sub_value_is_classified(
        self, key: str, tmp_path: Path
    ) -> None:
        """The failure is a ``ConfigurationTypeError`` naming cell, key and type.

        Args:
            key: The sub-key under test, used to look up its probe.
            tmp_path: pytest's per-test temporary directory.
        """
        bad, type_name = CELL_VALUE_PROBES[key]
        cells: dict[str, Any] = {"C1": {"value": 1, key: bad}}
        with pytest.raises(xlsxturbo.ConfigurationTypeError) as caught:
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", cells=cells)
        message = str(caught.value)
        for expected in ("cells['C1']", key, type_name):
            assert expected in message, (
                f"message does not name {expected!r}: {message!r}"
            )

    @pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required to read the workbook back")
    @pytest.mark.parametrize("key", sorted(CELL_VALUE_PROBES))
    def test_an_explicit_none_reads_as_unset(self, key: str, tmp_path: Path) -> None:
        """A ``None`` sub-value writes the cell as if the key were absent.

        The control for the test above: classifying the type failure is only
        half of it, and the half a caller hits far more often is passing the key
        with nothing in it.

        Args:
            key: The sub-key under test.
            tmp_path: pytest's per-test temporary directory.
        """
        out = tmp_path / "out.xlsx"
        cells: dict[str, Any] = {"C1": {"value": "x", key: None}}
        xlsxturbo.df_to_xlsx(_frame(), out, cells=cells)
        wb = load_workbook(out)
        cell = active_ws(wb)["C1"]
        assert cell.value == "x"
        assert cell.number_format == "General"
        assert cell.alignment.horizontal is None
        assert cell.alignment.vertical is None
        assert not cell.alignment.wrapText
        wb.close()


# A chart series item whose option dict has a non-string key. The mistake is the
# same one `NESTED_TYPE_PROBES["charts"]` makes at the chart level, but this one
# is only reachable at write time, when the series list stored unparsed at
# extract time is finally read.
BAD_SERIES_KEY: dict[str, Any] = {
    "charts": {"D2": {"type": "column", "series": [{1: "x"}]}}
}

# A textbox font dict with the same non-string key. The other option whose
# nested dict is first read at write time, through the same seam.
BAD_TEXTBOX_FONT_KEY: dict[str, Any] = {
    "textboxes": {"D1": {"text": "note", "font": {1: "y"}}}
}

# One entry per apply-time re-parse site: the kwargs that reach it, and the
# fragments its message must carry.
APPLY_TIME_NESTED_DICT_CASES = [
    pytest.param(BAD_SERIES_KEY, ("charts['D2']", "series item 0", "int"), id="chart-series"),
    pytest.param(BAD_TEXTBOX_FONT_KEY, ("textboxes['D1']", "'font'", "int"), id="textbox-font"),
]


class TestApplyTimeNestedDictsReportFaithfully:
    """A nested dict re-read at write time reports its own failure, not a class name.

    The apply layer's error currency is ``str`` and the boundary turns that into
    a ``ConfigurationError``. A ``PyErr`` flattened into that channel with
    ``to_string()`` renders as ``"<class>: <message>"``, so the caller reads
    ``ConfigurationTypeError:`` at the head of an exception that is not one --
    the demoted class name surviving as message text.
    """

    @pytest.mark.parametrize(("kwargs", "anchors"), APPLY_TIME_NESTED_DICT_CASES)
    def test_the_message_carries_no_class_name_prefix(
        self, kwargs: dict[str, Any], anchors: tuple[str, ...], tmp_path: Path
    ) -> None:
        """No class name heads the message, which still names anchor and type.

        Args:
            kwargs: the option that reaches an apply-time nested-dict re-parse.
            anchors: fragments the message must carry.
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.XlsxTurboError) as caught:
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", **kwargs)
        message = str(caught.value)
        assert not re.match(r"^\w+Error: ", message), (
            f"message is prefixed with an exception class name: {message!r}"
        )
        for expected in anchors:
            assert expected in message, (
                f"message does not name {expected!r}: {message!r}"
            )

    @pytest.mark.parametrize(("kwargs", "anchors"), APPLY_TIME_NESTED_DICT_CASES)
    def test_the_class_is_the_one_the_stability_promise_pins(
        self, kwargs: dict[str, Any], anchors: tuple[str, ...], tmp_path: Path
    ) -> None:
        """An apply-time nested-key failure is a ``ConfigurationError`` today.

        The same mistake at extract time is a ``ConfigurationTypeError``, so the
        value/type split reads differently depending on which layer catches it.
        Correcting that changes the builtin base an existing ``except`` clause
        sees, which ``docs/stability.md`` makes a major-version event -- this
        pins the current class so that change is a deliberate red test rather
        than silent drift.

        Args:
            kwargs: the option that reaches an apply-time nested-dict re-parse.
            anchors: fragments the message must carry (unused here; keeps the
                two tests on one parameter table).
            tmp_path: pytest's per-test temporary directory.
        """
        with pytest.raises(xlsxturbo.ConfigurationError) as caught:
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "out.xlsx", **kwargs)
        assert not isinstance(caught.value, TypeError), (
            "an `except TypeError` handler now catches this failure, which is a "
            "change of builtin base -- see docs/stability.md"
        )


@pytest.mark.skipif(
    not repo_checkout_available(),
    reason="reads src/convert.rs, which a wheel install does not carry",
)
class TestConvertErrorHasNoDefaultCategory:
    """No blanket ``String -> ConvertError`` conversion exists, and none returns.

    ``ConvertError`` once had ``From<String>``, mapping every untagged pipeline
    failure to ``Config`` and thence to ``ConfigurationError``. It made ``?``
    compile everywhere, and it was wrong in one direction only: a filesystem
    failure added later blamed the caller's options, silently. Two such
    misclassifications were found before it was removed.

    Removing it makes the Rust compiler ask the question at every new failure
    site, which is a stronger guarantee than any runtime test can give -- but
    only for as long as nobody adds the impl back for convenience. That is what
    this class watches, and it is why it reads the source: the property is the
    *absence* of code, which nothing observable at runtime can distinguish from
    a codebase that simply never took the shortcut.
    """

    def test_no_blanket_conversion_into_convert_error(self) -> None:
        """``impl From<...> for ConvertError`` is absent from the source."""
        source = (REPO_ROOT / "src" / "convert.rs").read_text(encoding="utf-8")
        assert "enum ConvertError" in source, (
            "src/convert.rs no longer defines ConvertError; this guard is reading "
            "the wrong file and would pass for the wrong reason"
        )
        offenders = re.findall(r"impl\s+From\s*<[^>]*>\s+for\s+ConvertError", source)
        assert not offenders, (
            f"a blanket conversion into ConvertError is back: {offenders}. Every "
            "failure site must name Config or File explicitly; see the note above "
            "the enum in src/convert.rs."
        )

    def test_both_categories_are_still_constructed(self) -> None:
        """...and the alternative to a default is that both variants get used.

        The control. Deleting the ``File`` variant's every construction site
        would satisfy the check above perfectly while making every filesystem
        failure a ``ConfigurationError`` again -- the exact outcome it exists to
        prevent.
        """
        source = (REPO_ROOT / "src" / "convert.rs").read_text(encoding="utf-8")
        for variant in ("ConvertError::Config", "ConvertError::File"):
            assert source.count(variant) >= 2, (
                f"{variant} is constructed {source.count(variant)} time(s) in "
                "convert.rs; a category nothing produces is not a category"
            )
