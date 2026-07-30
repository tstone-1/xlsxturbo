# Reusable option bundles

`df_to_xlsx` takes 27 parameters. That is fine for a call with two options and
unpleasant for a call with ten — and it gives you nowhere to put a set of options
you want to use more than once.

`ExportOptions` is a typed, frozen bundle of exactly those options:

```python
import xlsxturbo
from xlsxturbo import ExportOptions

REPORT = ExportOptions(
    freeze_panes=True,
    autofit=True,
    header_format={"bold": True, "bg_color": "#DDDDDD"},
    table_style="Medium9",
)

xlsxturbo.df_to_xlsx(df, "out.xlsx", **REPORT.as_kwargs())
```

Every keyword argument still works exactly as before. This is an addition, and
nothing is deprecated — if you pass two options, keep passing two options.

## Why `**` rather than `options=`

`df_to_xlsx(df, path, options=REPORT)` would read better, and it would cost more
than it is worth. The compiled function would need a Python wrapper, and that
wrapper would either repeat all 27 parameters — a second copy to keep in step —
or collapse them to `**kwargs`, which is what your editor would then show you
instead of 27 named, typed parameters.

Since the point of this type is discoverability, trading the signature's
discoverability for it would be a bad deal. So the entry points are untouched and
the bundle lowers to what they already accept.

## Sharing a bundle across sheets

The same object produces a per-sheet dict for `dfs_to_xlsx`:

```python
xlsxturbo.dfs_to_xlsx(
    [
        (q1, "Q1", REPORT.as_sheet_options()),
        (q2, "Q2", REPORT.as_sheet_options()),
    ],
    "quarters.xlsx",
)
```

`as_sheet_options()` drops the two options a per-sheet dict does not accept —
`constant_memory` and `defined_names`, both of which apply to the workbook as a
whole. Everything else passes through.

## Deriving variants

The bundle is frozen, so it is safe as a module-level constant: no caller can
mutate it for everyone else. Build variants instead of editing.

`dataclasses.replace` overrides named fields:

```python
from dataclasses import replace

DRAFT = replace(REPORT, freeze_panes=False)
```

`merged_with` layers one bundle over another, taking only the options the second
one actually set:

```python
BASE = ExportOptions(autofit=True, header_format={"bold": True})

SUMMARY = BASE.merged_with(ExportOptions(table_style="Medium2"))
# autofit and header_format survive; table_style is added
```

That last property is what makes a shared base worth having. A merge that took
*every* field would silently reset the base's options to their defaults wherever
the override said nothing.

## Unset is not `None`

An option you never touched and an option you set to `None` are different, and
the bundle keeps them apart:

```python
ExportOptions().as_kwargs()                    # {}
ExportOptions(table_style=None).as_kwargs()    # {"table_style": None}
```

This matters on the multi-sheet path, where an explicit `None` means "not on this
sheet", deliberately overriding a workbook-wide default:

```python
xlsxturbo.dfs_to_xlsx(
    [
        (summary, "Summary", ExportOptions(table_style=None).as_sheet_options()),
        (detail, "Detail", {}),
    ],
    "report.xlsx",
    table_style="Medium9",   # applies to Detail, not to Summary
)
```

Collapsing the two would make that impossible to express, so unset options are
omitted from both lowerings and `None` is passed through.

## What is not in the bundle

`df`, `output_path` and `sheet_name`. They identify *this* call rather than
configuring it, so they stay positional arguments where they belong.

The per-feature shapes — `ChartOptions`, `ValidationOptions`, `HeaderFormat` and
the rest — are `TypedDict`s in [`xlsxturbo.types`](api-reference.md), not
dataclasses. They already describe themselves to a type checker, and a second
representation of the same thing would only be another surface to keep in step.
