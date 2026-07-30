"""A structured, reusable bundle of export options.

``df_to_xlsx`` takes 27 parameters. That is the whole problem this module exists
to solve: the options are discoverable only by reading a long signature, they
cannot be built up in pieces, and a set of them cannot be named, stored, or
passed around::

    from xlsxturbo.options import ExportOptions

    REPORT = ExportOptions(
        freeze_panes=True,
        autofit=True,
        header_format={"bold": True, "bg_color": "#DDDDDD"},
    )

    xlsxturbo.df_to_xlsx(df, "out.xlsx", **REPORT.as_kwargs())

The same bundle lowers to a per-sheet dict for the multi-sheet entry point::

    xlsxturbo.dfs_to_xlsx(
        [(q1, "Q1", REPORT.as_sheet_options()), (q2, "Q2", REPORT.as_sheet_options())],
        "out.xlsx",
    )

Three design decisions are worth knowing, because each had a plausible
alternative.

**It lowers to keyword arguments; it does not wrap the entry points.** The
obvious spelling is ``df_to_xlsx(df, path, options=REPORT)``, which needs a
Python wrapper around the compiled function. That wrapper would have to either
repeat all 27 parameters -- a hand-maintained duplicate, and duplicated lists
are what this repository keeps getting caught by -- or collapse them to
``**kwargs``, which would replace an explicit, fully typed signature with an
opaque one. An IDE showing 27 named parameters today would show ``**kwargs``
instead. That is the opposite of the goal, so the compiled functions are left
exactly as they are and every existing call keeps working untouched.

**It is flat, not a tree of grouped objects.** ``LayoutOptions``,
``TableOptions`` and friends were considered and dropped: the natural groups are
two to five fields each, none of them earns a public class, and nesting would
add five names and a level of indirection for no capability. Flat also keeps the
mapping to keyword arguments one-to-one, which is what makes the coverage guard
in ``tests/test_options.py`` trivial enough to trust.

**Unset is not the same as ``None``.** Passing ``table_style=None`` in a
per-sheet dict means "no table on this sheet", deliberately shadowing a workbook
default -- verified behaviour, not a guess. So an option nobody touched has to be
distinguishable from one explicitly set to ``None``, and both lowerings omit only
the former.

Every option remains available as a plain keyword argument, indefinitely. This is
an addition, not a replacement, and nothing here is deprecated.
"""

from __future__ import annotations

from dataclasses import dataclass, fields
from typing import Any

from xlsxturbo.types import (
    CellValueOptions,
    ChartOptions,
    ColumnFormat,
    CommentOptions,
    ConditionalFormat,
    HeaderFormat,
    ImageOptions,
    RichTextFormat,
    SheetOptions,
    SparklineOptions,
    TextboxOptions,
    ValidationOptions,
)

__all__ = ["ExportOptions"]

# Sentinel for "this option was never set", distinct from an explicit `None`.
#
# Typed `Any` so each field can declare its real type while defaulting to this.
# The alternative -- widening every annotation to include a sentinel type --
# would put an implementation detail into the signature users read, which is the
# one thing this module exists to keep clean.
_UNSET: Any = object()

# Options the workbook accepts but a per-sheet dict does not. Kept as the
# difference from the field list rather than as a second copy of it, so
# `as_sheet_options` cannot drift from the dataclass.
_WORKBOOK_ONLY = frozenset({"constant_memory", "defined_names"})


@dataclass(frozen=True)
class ExportOptions:
    """A reusable bundle of export options, lowered to keyword arguments.

    Field names and meanings are identical to the keyword arguments of
    :func:`xlsxturbo.df_to_xlsx`; see the API reference for what each one does.
    Any field left unset is omitted when lowering, so the library's own defaults
    apply.

    Frozen, so a bundle can be shared as a module-level constant without a caller
    mutating it for everyone else. Use :func:`dataclasses.replace` to derive a
    variant::

        from dataclasses import replace

        DRAFT = replace(REPORT, freeze_panes=False)
    """

    # Layout
    header: bool = _UNSET
    autofit: bool = _UNSET
    freeze_panes: bool = _UNSET
    column_widths: dict[Any, Any] | None = _UNSET
    row_heights: dict[int, float] | None = _UNSET

    # Excel table
    table_style: str | None = _UNSET
    table_name: str | None = _UNSET

    # Formatting
    header_format: HeaderFormat | None = _UNSET
    column_formats: dict[str, ColumnFormat] | None = _UNSET
    conditional_formats: dict[str, ConditionalFormat | list[ConditionalFormat]] | None = _UNSET

    # Content
    cells: dict[str, str | int | float | bool | CellValueOptions] | None = _UNSET
    formula_columns: dict[str, str] | None = _UNSET
    merged_ranges: list[Any] | None = _UNSET
    hyperlinks: list[Any] | None = _UNSET
    comments: dict[str, str | CommentOptions] | None = _UNSET
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None = _UNSET
    validations: dict[str, ValidationOptions] | None = _UNSET

    # Media and charts
    images: dict[str, str | ImageOptions] | None = _UNSET
    checkboxes: dict[str, Any] | None = _UNSET
    textboxes: dict[str, str | TextboxOptions] | None = _UNSET
    charts: dict[str, ChartOptions] | None = _UNSET
    sparklines: dict[str, SparklineOptions] | None = _UNSET

    # Workbook-level: accepted by the entry points, not by a per-sheet dict.
    constant_memory: bool = _UNSET
    defined_names: dict[str, str] | None = _UNSET

    def as_kwargs(self) -> dict[str, Any]:
        """Return the set options as keyword arguments for the entry points.

        Returns:
            A mapping of option name to value, containing only options that were
            actually set. Splat it into :func:`xlsxturbo.df_to_xlsx`,
            :func:`xlsxturbo.dfs_to_xlsx`, or any call accepting those keywords.
        """
        return {
            f.name: getattr(self, f.name)
            for f in fields(self)
            if getattr(self, f.name) is not _UNSET
        }

    def as_sheet_options(self) -> SheetOptions:
        """Return the set options as a per-sheet dict for ``dfs_to_xlsx``.

        Drops the two workbook-level options, which a per-sheet dict rejects as
        unknown keys. Dropping them silently is the right call here: it lets one
        bundle serve both call shapes, which is the point of the type.

        Returns:
            A :class:`~xlsxturbo.types.SheetOptions` mapping, containing only
            options that were set and are valid per sheet.
        """
        kwargs = self.as_kwargs()
        sheet: Any = {k: v for k, v in kwargs.items() if k not in _WORKBOOK_ONLY}
        return sheet

    def merged_with(self, other: ExportOptions) -> ExportOptions:
        """Return a copy with ``other``'s set options layered on top of this one.

        Only options ``other`` actually set are taken, so a sparse override does
        not reset everything else to a default -- which is what makes a shared
        base bundle useful::

            SHEET = BASE.merged_with(ExportOptions(table_style="Medium2"))

        Args:
            other: The bundle whose set options win.

        Returns:
            A new bundle; neither input is modified.
        """
        return ExportOptions(**{**self.as_kwargs(), **other.as_kwargs()})
