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
from pathlib import Path

import pytest
from xlsxturbo import types as types_module

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
    """Public shape names actually defined in ``types.py`` at runtime."""
    names = [
        name
        for name in dir(types_module)
        if not name.startswith("_") and name not in {"annotations", "Literal", "TypedDict", "Union", "PathLike"}
    ]
    assert len(names) >= 10, f"only found {len(names)} runtime shapes: {names}"
    return sorted(names)


REEXPORTS = _stub_reexports()


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

    def test_pathalias_is_not_a_pep604_union(self) -> None:
        """``PathArg`` is spelled with ``Union`` so it evaluates on Python 3.9.

        A ``.pyi`` is never executed, so ``str | PathLike[str]`` is free there and
        was what the stub used. In a runtime module on 3.9 that same expression
        raises ``TypeError`` at import. Verified against a real 3.9 interpreter
        when this module was written; pinned here as source text, because on 3.10+
        both spellings import and the difference is invisible at runtime.
        """
        source = TYPES_SOURCE.read_text(encoding="utf-8")
        assert "PathArg = Union[str, PathLike[str]]" in source
        assert "PathArg = str | PathLike[str]" not in source

    def test_future_annotations_is_enabled(self) -> None:
        """Field annotations stay unevaluated, which is what allows ``bool | str``.

        Without this import, every ``|`` in a field annotation would be evaluated
        at class-creation time and the module would not import on Python 3.9 --
        the failure the ``PathArg`` test above pins for module-level aliases.
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
