"""Tests that `ExportOptions` field types match the signature they mirror.

`ExportOptions` exists to make 27 keyword arguments discoverable, and it is
worth exactly as much as its types are. It is a second surface describing the
same options as `df_to_xlsx`, whose authoritative types live in
`python/xlsxturbo/xlsxturbo.pyi` -- and two surfaces describing one thing drift.
They had: four fields had decayed to `dict[Any, Any]`, `list[Any]` or
`dict[str, Any]` while the stub carried the precise type all along, and one had
drifted *narrower* than the stub, so a checker rejected a call the runtime
accepts.

`tests/test_options.py` already pins the field *set* against
`inspect.signature(df_to_xlsx)`. It cannot check types: the compiled function
carries no annotations, which is why the stub exists. This module closes that
half by comparing annotation source text.

Comparing text rather than resolved types is deliberate. `dict[int | str, ...]`
and `dict[Union[int, str], ...]` mean the same thing to a checker, and a reader
comparing the two files sees a difference; the text is what a maintainer edits
and what an IDE renders, so text is the thing to keep identical.
"""

from __future__ import annotations

import ast

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="these tests read repository sources, which a wheel install does not carry",
)

STUB = REPO_ROOT / "python" / "xlsxturbo" / "xlsxturbo.pyi"
OPTIONS = REPO_ROOT / "python" / "xlsxturbo" / "options.py"

# Fields whose annotation is allowed to differ from the stub, each with the
# reason. Empty, and meant to stay that way: an entry here is a documented
# divergence between what a user reads on the dataclass and what the function
# actually accepts, which is the whole defect this module exists to prevent.
ALLOWED_DIVERGENCE: dict[str, str] = {}


def _stub_parameter_annotations() -> dict[str, str]:
    """Parameter name -> annotation source, from `df_to_xlsx` in the stub."""
    tree = ast.parse(STUB.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name == "df_to_xlsx":
            args = node.args
            annotations = {
                argument.arg: ast.unparse(argument.annotation)
                for argument in (*args.posonlyargs, *args.args, *args.kwonlyargs)
                if argument.annotation is not None
            }
            # A parser that matches nothing reads exactly like full agreement.
            assert len(annotations) >= 20, (
                f"only parsed {len(annotations)} annotated parameters from df_to_xlsx; "
                f"the comparison below would be nearly empty"
            )
            return annotations
    raise AssertionError("df_to_xlsx not found in the stub")


def _dataclass_field_annotations() -> dict[str, str]:
    """Field name -> annotation source, from the `ExportOptions` dataclass."""
    tree = ast.parse(OPTIONS.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == "ExportOptions":
            annotations = {
                statement.target.id: ast.unparse(statement.annotation)
                for statement in node.body
                if isinstance(statement, ast.AnnAssign)
                and isinstance(statement.target, ast.Name)
                and statement.annotation is not None
            }
            assert len(annotations) >= 20, (
                f"only parsed {len(annotations)} fields from ExportOptions; the comparison "
                f"below would be nearly empty"
            )
            return annotations
    raise AssertionError("ExportOptions not found in options.py")


class TestExportOptionsMirrorsTheSignature:
    """Each field says the same thing the stub says about the same option."""

    def test_every_shared_option_has_the_same_annotation(self) -> None:
        """No field is looser, narrower, or otherwise different from the stub.

        Both directions matter and for different reasons. A **looser** field
        (`list[Any]`) silently drops the guidance the bundle exists to provide.
        A **narrower** one is worse: it makes a checker reject a call the
        runtime accepts, so the user's working code is reported as wrong.
        """
        stub = _stub_parameter_annotations()
        dataclass_fields = _dataclass_field_annotations()
        shared = sorted(set(stub) & set(dataclass_fields) - set(ALLOWED_DIVERGENCE))
        assert shared, "ExportOptions and df_to_xlsx share no option names at all"

        mismatched = [
            f"{name}: ExportOptions has {dataclass_fields[name]!r}, stub has {stub[name]!r}"
            for name in shared
            if dataclass_fields[name] != stub[name]
        ]
        assert not mismatched, (
            f"{len(mismatched)} field(s) disagree with the signature they mirror:\n  "
            + "\n  ".join(mismatched)
        )

    def test_no_field_annotation_is_a_bare_any_container(self) -> None:
        """No field falls back to `Any` inside a container.

        The test above already catches this while the stub is precise, and this
        one stays useful if both surfaces are ever made vague together -- which
        is the shape a drift fix takes when someone resolves a mismatch by
        widening the wrong side.
        """
        vague = [
            f"{name}: {annotation}"
            for name, annotation in _dataclass_field_annotations().items()
            if "Any" in annotation
        ]
        assert not vague, (
            f"these ExportOptions fields fall back to Any, so the bundle gives an IDE "
            f"nothing for exactly the options that need it most: {vague}"
        )

    def test_the_allowed_divergence_list_is_documented_and_real(self) -> None:
        """Any accepted divergence names a real field and carries a reason.

        Guards the escape hatch rather than the rule: an entry added to silence
        a failure, for a field that no longer exists or with an empty reason,
        turns this module off one option at a time.
        """
        dataclass_fields = _dataclass_field_annotations()
        unknown = sorted(set(ALLOWED_DIVERGENCE) - set(dataclass_fields))
        assert not unknown, f"ALLOWED_DIVERGENCE names fields that do not exist: {unknown}"
        unexplained = sorted(name for name, why in ALLOWED_DIVERGENCE.items() if not why.strip())
        assert not unexplained, f"ALLOWED_DIVERGENCE entries with no reason: {unexplained}"
