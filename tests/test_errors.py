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

from collections.abc import Callable
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

# The public names, and the builtin each one must remain an instance of. A class
# appearing here with the wrong builtin is a breaking change for anyone whose
# `except` clause names that builtin.
EXPECTED_BASES: dict[str, tuple[type[BaseException], ...]] = {
    "XlsxTurboError": (Exception,),
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


# One real call per concrete class. `XlsxTurboError` is the abstract base and is
# covered by every entry here being an instance of it.
TRIGGERS: dict[str, Callable[[Path], None]] = {
    "ConfigurationError": _trigger_configuration,
    "ConfigurationTypeError": _trigger_configuration_type,
    "InputDataError": _trigger_input_data,
    "FileError": _trigger_file,
    "WorkbookValidationError": _trigger_workbook_validation,
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
        concrete = exported - {"XlsxTurboError"}
        assert concrete == set(TRIGGERS), (
            f"exported without a trigger: {sorted(concrete - set(TRIGGERS))}; "
            f"triggered but not exported: {sorted(set(TRIGGERS) - concrete)}"
        )


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
        """``except XlsxTurboError`` is sufficient for every trigger."""
        for name, trigger in sorted(TRIGGERS.items()):
            with pytest.raises(xlsxturbo.XlsxTurboError):
                trigger(tmp_path)
            del name

    def test_workbook_validation_is_a_configuration_error(self) -> None:
        """The one intentional parent/child pair inside the hierarchy holds."""
        assert issubclass(xlsxturbo.WorkbookValidationError, xlsxturbo.ConfigurationError)

    def test_file_error_leaves_the_oserror_fields_unset(self, tmp_path: Path) -> None:
        """``errno``/``strerror``/``filename`` are ``None``, as documented.

        ``FileError`` inherits ``OSError`` for its ``except`` behaviour, but it is
        constructed with a single message argument, so the structured fields are
        never populated. Documented in ``docs/errors.md``; pinned here because a
        two-argument construction would silently start producing
        ``[Errno x] y`` messages instead.
        """
        with pytest.raises(xlsxturbo.FileError) as caught:
            _trigger_file(tmp_path)
        assert caught.value.errno is None
        assert caught.value.strerror is None
        assert caught.value.filename is None
        assert "Failed to save workbook to" in str(caught.value)


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
        """``repr()`` names the path users import from, and pickling follows it."""
        cls = getattr(xlsxturbo, name)
        assert cls.__module__ == "xlsxturbo"
        assert cls.__qualname__ == name

    @pytest.mark.parametrize("name", sorted(EXPECTED_BASES))
    def test_class_has_a_docstring(self, name: str) -> None:
        """Each class documents itself; ``help()`` is the first thing users try."""
        doc = getattr(xlsxturbo, name).__doc__
        assert doc is not None
        assert len(doc.strip()) > 40
