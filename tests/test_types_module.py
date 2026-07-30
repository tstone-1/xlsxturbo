"""Tests coupling ``xlsxturbo.types`` to the stub that re-exports it.

``python/xlsxturbo/types.py`` is the authoritative home for the option shapes and
``python/xlsxturbo/xlsxturbo.pyi`` imports them (decision D1 in
``docs/roadmap-1.0.md``). Nothing in the type checker's world notices if those two
drift: the stub could re-export a name ``types.py`` no longer defines, and pyright
would report an error only for whoever imported it. These tests close that gap by
reading the stub as text and the module at runtime.

The population is read from the stub rather than hardcoded here, because a list
of expected names in a test is the thing that rots -- and a hardcoded list that
someone trims to make a test pass is how a shape silently stops being exported.
"""

from __future__ import annotations

import re
import typing
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo
from xlsxturbo import types as types_module


def _frame() -> pd.DataFrame:
    """Return the smallest DataFrame the option probes can be applied to."""
    return pd.DataFrame({"Name": ["a"], "Score": [1]})

PACKAGE_DIR = Path(str(types_module.__file__)).parent
STUB = PACKAGE_DIR / "xlsxturbo.pyi"
TYPES_SOURCE = PACKAGE_DIR / "types.py"


def _stub_reexports() -> list[str]:
    """Names the extension stub re-exports from ``xlsxturbo.types``.

    Parsed from the ``from xlsxturbo.types import (...)`` block, matching the
    redundant-alias form (``X as X``) that makes a stub import public.
    """
    text = STUB.read_text(encoding="utf-8")
    block = re.search(r"from xlsxturbo\.types import \(\n(.*?)\n\)", text, re.DOTALL)
    assert block is not None, "the stub no longer imports from xlsxturbo.types"
    names = re.findall(r"^\s+([A-Za-z_][A-Za-z0-9_]*) as \1,$", block.group(1), re.MULTILINE)
    # A parser that matches nothing reads exactly like a clean result. The bound
    # is deliberately loose: it guards against the regex breaking, not against a
    # name being removed. A tighter bound made this module fail to *import* when
    # one name was dropped, so the comparison it feeds never ran and the drop got
    # reported as a collection error instead of by the test that exists for it.
    assert len(names) >= 10, f"only parsed {len(names)} re-exports: {names}"
    return sorted(names)


def _runtime_shapes() -> list[str]:
    """Public shape names ``types.py`` declares in ``__all__``.

    Reads the declaration rather than ``dir()`` minus a list of typing helpers to
    ignore. The old spelling defined the public API as "everything except these
    implementation details", which is the wrong way round twice over: it made the
    test assert a cleaner namespace than a wildcard import actually produced, and
    it turned every new typing import into a test edit.
    """
    names = list(types_module.__all__)
    assert len(names) >= 10, f"only found {len(names)} names in __all__: {names}"
    return sorted(names)


REEXPORTS = _stub_reexports()

# Names `types.py` imports for its own use, which `dir()` therefore reports as
# public but `__all__` deliberately omits. Guarded below against naming an
# import that no longer exists.
TYPING_HELPERS = {"annotations", "Literal", "PathLike", "TypedDict"}


class TestStubAndRuntimeAgree:
    """The stub's re-export list and the runtime module define the same names."""

    def test_name_sets_are_identical(self) -> None:
        """Neither surface has a name the other lacks."""
        stub, runtime = set(REEXPORTS), set(_runtime_shapes())
        assert stub == runtime, (
            f"re-exported but not defined: {sorted(stub - runtime)}; "
            f"defined but not re-exported: {sorted(runtime - stub)}"
        )

    @pytest.mark.parametrize("name", REEXPORTS)
    def test_every_reexported_name_exists_at_runtime(self, name: str) -> None:
        """Each re-exported name is a real attribute, not just a stub promise."""
        assert hasattr(types_module, name)

    def test_shapes_are_not_declared_in_the_stub_as_well(self) -> None:
        """The stub imports the shapes; it does not keep a second copy.

        A re-declared ``class HeaderFormat(TypedDict)`` in the stub would shadow
        the import, and the two could then disagree indefinitely -- which is the
        duplication D1 exists to remove.
        """
        text = STUB.read_text(encoding="utf-8")
        redeclared = [
            name for name in REEXPORTS if re.search(rf"^class {name}\(TypedDict", text, re.MULTILINE)
        ]
        assert not redeclared, f"the stub re-declares: {redeclared}"


class TestRuntimeUsability:
    """The shapes work as runtime objects, which is the point of the module."""

    def test_typeddicts_expose_their_fields(self) -> None:
        """A representative ``TypedDict`` has the fields it documents."""
        assert "bold" in types_module.HeaderFormat.__annotations__
        assert "bg_color" in types_module.HeaderFormat.__annotations__

    def test_a_typeddict_accepts_a_conforming_dict(self) -> None:
        """Annotating a dict literal with a shape works without a guard."""
        header: types_module.HeaderFormat = {"bold": True, "font_size": 12.0}
        assert header["bold"] is True

    def test_module_imports_without_the_extension(self) -> None:
        """``types.py`` depends only on the standard library.

        This is what makes it safe to import from anywhere, including from code
        that runs before the extension is built. Asserted against the source
        because an accidental ``from .xlsxturbo import ...`` would work fine on a
        developer machine and only fail in a fresh checkout.
        """
        source = TYPES_SOURCE.read_text(encoding="utf-8")
        imports = re.findall(r"^(?:from|import)\s+(\S+)", source, re.MULTILINE)
        assert imports, "no imports parsed from types.py"
        for module in imports:
            assert not module.startswith("."), f"types.py imports {module} relatively"
            assert "xlsxturbo" not in module, f"types.py imports {module}"

    def test_future_annotations_is_enabled(self) -> None:
        """Field annotations stay unevaluated, which keeps import cost off the shapes.

        Load-bearing for a different reason since 1.1.0. While the floor was 3.9
        this import was what let a field be written ``bool | str`` at all -- the
        expression is a syntax 3.9 cannot evaluate, and a companion test pinned
        module-level aliases to ``Union[...]`` for the same reason. Both are gone
        with 3.9. What remains is that these annotations are never evaluated at
        class creation, so a forward reference costs nothing and the module stays
        importable before the extension is built.
        """
        source = TYPES_SOURCE.read_text(encoding="utf-8")
        # Anchored to column 0, and deliberately so: types.py *documents* this
        # import in its own docstring, so an unanchored substring check still
        # passes after the statement itself is deleted. That exact mutation
        # survived until this became a regex -- a checker scanning text its own
        # bookkeeping also lives in needs the anchor.
        assert re.search(r"^from __future__ import annotations$", source, re.MULTILINE), (
            "types.py has no `from __future__ import annotations` statement"
        )
        assert "bool | str" in source, "no PEP 604 field annotation left to protect"


class TestPublicNamespace:
    """``__all__`` is accurate, and is what a wildcard import actually gives."""

    def test_every_declared_name_exists(self) -> None:
        """``__all__`` names nothing that is not defined.

        A stale entry breaks ``from xlsxturbo.types import *`` with an
        ``AttributeError`` at import time, for every user at once.
        """
        missing = [n for n in types_module.__all__ if not hasattr(types_module, n)]
        assert not missing, f"__all__ names undefined attribute(s): {missing}"

    def test_no_public_name_is_left_out(self) -> None:
        """Every public module-level name is declared.

        Catches the opposite drift: a shape added to the module and forgotten in
        ``__all__`` is invisible to a wildcard import and to tooling that reads
        the declaration, while still being importable by name -- so nothing else
        in this suite would notice.
        """
        public = {
            n for n in dir(types_module) if not n.startswith("_") and n not in TYPING_HELPERS
        }
        undeclared = public - set(types_module.__all__)
        assert not undeclared, f"public name(s) missing from __all__: {sorted(undeclared)}"

    def test_the_helper_exclusions_are_all_real_imports(self) -> None:
        """The exclusion set does not outlive the imports it excludes.

        ``Optional`` and ``Union`` sat here after 1.1.0 stopped importing them,
        which is harmless and exactly the shape of rot this suite distrusts
        elsewhere: an exclusion naming something that no longer exists reads as
        a deliberate carve-out and is nothing of the kind.
        """
        stale = TYPING_HELPERS - set(dir(types_module))
        assert not stale, f"excluded names types.py no longer imports: {sorted(stale)}"

    def test_wildcard_import_leaks_no_typing_helpers(self) -> None:
        """``import *`` gives the shapes and nothing else.

        The property the two tests above only imply. Executed rather than
        reasoned about, because ``__all__`` is what makes it true and a test that
        re-derives the answer from ``__all__`` would pass even if the module
        stopped declaring one.
        """
        namespace: dict[str, object] = {}
        exec("from xlsxturbo.types import *", namespace)  # noqa: S102 - our own module
        imported = {n for n in namespace if not n.startswith("__")}
        assert imported == set(types_module.__all__), (
            f"unexpected: {sorted(imported - set(types_module.__all__))}; "
            f"absent: {sorted(set(types_module.__all__) - imported)}"
        )


# Shape -> (the field it must mark required, a call that omits it). Both halves
# are asserted: that the type says the field is required, and that the runtime
# agrees by rejecting a dict without it. Either alone is half a contract.
REQUIRED_FIELDS: dict[str, tuple[str, dict[str, object]]] = {
    "CommentOptions": ("text", {"comments": {"D1": {}}}),
    "ValidationOptions": ("type", {"validations": {"Name": {}}}),
    "ImageOptions": ("path", {"images": {"D1": {}}}),
    "CheckboxOptions": ("checked", {"checkboxes": {"D1": {}}}),
    "TextboxOptions": ("text", {"textboxes": {"D1": {}}}),
    "ConditionalFormat": ("type", {"conditional_formats": {"Score": {}}}),
    "ChartOptions": ("type", {"charts": {"D2": {}}}),
    "SparklineOptions": ("range", {"sparklines": {"D2": {}}}),
    "CellValueOptions": ("value", {"cells": {"D1": {}}}),
}

# Shapes where every field really is optional, so `total=False` throughout is
# correct. Listed rather than inferred, so that a shape gaining a required field
# and not gaining a required base fails the partition check below.
FULLY_OPTIONAL = frozenset(
    {"HeaderFormat", "ColumnFormat", "RichTextFormat", "TextboxFont", "SheetOptions"}
)


class TestRequiredFieldsAreTyped:
    """A field the runtime demands is marked required in the shape.

    Every one of these carried a docstring note saying it was "required at
    runtime but TypedDict doesn't enforce this". It can be enforced -- a required
    base plus a ``total=False`` subclass has worked on every version this package
    has ever supported -- so the note was describing a gap, not a limitation.
    Until 0.19.1 a type checker accepted ``images={"D1": {}}``, which raises.
    """

    @pytest.mark.parametrize("shape", sorted(REQUIRED_FIELDS))
    def test_shape_marks_the_field_required(self, shape: str) -> None:
        """``__required_keys__`` contains the field.

        Args:
            shape: The TypedDict under test.
        """
        field, _ = REQUIRED_FIELDS[shape]
        cls = getattr(types_module, shape)
        assert field in cls.__required_keys__, (
            f"{shape}.__required_keys__ is {set(cls.__required_keys__)}, missing {field!r}"
        )

    @pytest.mark.parametrize("shape", sorted(REQUIRED_FIELDS))
    def test_runtime_rejects_the_dict_without_it(self, shape: str, tmp_path: Path) -> None:
        """Omitting the field is a real error, so requiring it is not overreach.

        Without this half, the shapes could mark a field required that the
        runtime happily defaults, and every user would have to pass it for no
        reason.

        Args:
            shape: The TypedDict under test.
            tmp_path: pytest's per-test temporary directory.
        """
        _, call = REQUIRED_FIELDS[shape]
        with pytest.raises(xlsxturbo.XlsxTurboError, match=r"(?i)missing|require"):
            xlsxturbo.df_to_xlsx(_frame(), tmp_path / "o.xlsx", **call)  # type: ignore[arg-type]

    def test_every_shape_is_classified(self) -> None:
        """The two tables partition the shapes; none is silently unexamined.

        A new shape lands in neither table and fails here, rather than defaulting
        to unchecked. ``ChartSeriesOptions`` is the one deliberate exclusion: its
        requirement is a one-of across ``values_range``/``values``/``data_range``,
        which a TypedDict cannot express, and it is documented as such.
        """
        conditional = {"ChartSeriesOptions"}
        classified = set(REQUIRED_FIELDS) | FULLY_OPTIONAL | conditional
        declared = {
            name
            for name in types_module.__all__
            if isinstance(getattr(types_module, name), type)
        }
        assert declared == classified, (
            f"unclassified shape(s): {sorted(declared - classified)}; "
            f"classified but nonexistent: {sorted(classified - declared)}"
        )

    @pytest.mark.parametrize("shape", sorted(FULLY_OPTIONAL))
    def test_optional_shapes_require_nothing(self, shape: str) -> None:
        """A shape listed as fully optional really has no required key.

        Keeps ``FULLY_OPTIONAL`` from becoming a place to park a shape whose
        required field nobody wanted to add.

        Args:
            shape: The TypedDict under test.
        """
        cls = getattr(types_module, shape)
        assert not cls.__required_keys__, (
            f"{shape} requires {set(cls.__required_keys__)} but is listed as fully optional"
        )


class TestAnnotationsResolveOnEverySupportedPython:
    """``typing.get_type_hints`` works on the option shapes.

    Frameworks resolve annotations rather than reading them as strings --
    pydantic, FastAPI, attrs, and anything building a form or a schema from a
    type -- so a shape whose hints cannot be resolved is a typed API that cannot
    be read programmatically.

    This was the harder half of the problem while the floor was 3.9: PEP 604
    unions are a syntax 3.9 cannot evaluate, ``from __future__ import
    annotations`` hid that at import time, and the failure appeared only when
    something else resolved the hints. Raising the floor to 3.10 in 1.1.0
    removed the constraint -- ``str | None`` evaluates there -- so the source
    scan that enforced the ``Union[...]`` spelling is gone and this test is what
    is left.
    """

    def test_get_type_hints_resolves_every_shape(self) -> None:
        """Every exported TypedDict resolves without raising.

        Catches what a source scan cannot see: a forward reference to a name
        that no longer exists.
        """
        hints_by_shape = {}
        for name in sorted(_runtime_shapes()):
            shape = getattr(types_module, name)
            if not (isinstance(shape, type) and issubclass(shape, dict)):
                continue
            hints_by_shape[name] = typing.get_type_hints(shape)
        assert hints_by_shape, "no TypedDict shapes were resolved"
        assert typing.get_type_hints(types_module.SheetOptions)["table_style"] == (str | None)

    # `test_no_annotation_uses_a_pep_604_union` lived here until 1.1.0. It read
    # the source through `inspect.getsource` and failed on any `|` in an
    # annotation, because 3.9 could not evaluate one and the mistake was
    # invisible on a developer machine. Raising the floor to 3.10 made the
    # syntax legal everywhere the package runs, so it was deleted rather than
    # weakened -- the constraint is now enforced by the language version, which
    # is a stronger guarantee than a test.
