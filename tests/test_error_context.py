"""Every option failure names the option, and the key inside it where there is one.

``AGENTS.md`` states the convention for the apply layer as
``<feature>['<key>']: <message>``. The parsers under ``src/parse/`` are pure --
``parse_color`` is handed ``"nope"`` and cannot know which option it came from
-- so meeting that convention is the caller's job, and most callers were not
doing it. Measured against 1.3.0, six of six sampled failures arrived with no
option and no key at all::

    conditional_formats={"a": {...}, "score": {"criteria": "bogus"}}
      -> Unknown criteria 'bogus'. Valid: ...        (which of the two entries?)
    header_format={"bg_color": "nope"}   -> Unknown color: nope
    charts={"D2": {"type": "bogus"}}     -> Unknown chart type 'bogus'. ...

This module is the guard for that. It is deliberately about *text*, which
``docs/stability.md`` does not cover -- what it pins is that a message can be
traced back to the dict entry that caused it, not the wording. Each case
therefore asserts the substrings that identify the entry and leaves the rest of
the sentence alone.

Two things a reader should know before adding a case:

- **The prefix is additive.** The pre-existing sentence is untouched, including
  its capitalisation, so the ``pytest.raises(match=...)`` assertions elsewhere
  in the suite -- a case-sensitive ``re.search`` -- keep matching.
- **Where the key adds nothing, there is no key.** A bad cell reference is
  reported as ``charts['ZZ']: ...`` and not ``charts['ZZ']: 'ZZ': ...``; the
  key and the failing value are the same string.
"""

from __future__ import annotations

from typing import Any

import pandas as pd
import pytest
import xlsxturbo

# Two columns so a pattern-keyed option has something to be ambiguous about:
# with one column, "which entry?" has only one possible answer and the test
# could pass on a message that named nothing.
FRAME = pd.DataFrame({"a": [1, 2], "score": [3, 4]})

# (label, kwargs, substrings that must all appear in str(exc))
#
# The first element of each expectation is the option context. Where a second
# is given it is the key inside that option, which is what tells apart the
# three colour fields of one format dict.
CASES: list[tuple[str, dict[str, Any], list[str]]] = [
    # --- conditional_formats: five distinct parsers, one context ------------
    (
        "cf_criteria",
        {
            "conditional_formats": {
                "a": {"type": "cell", "criteria": "gt", "value": 1},
                "score": {"type": "cell", "criteria": "bogus", "value": 1},
            }
        },
        ["conditional_formats['score']", "'criteria'", "Unknown criteria 'bogus'"],
    ),
    (
        "cf_type",
        {"conditional_formats": {"score": {"type": "bogus"}}},
        ["conditional_formats['score']", "'type'", "Unknown conditional format type"],
    ),
    (
        "cf_min_color",
        {"conditional_formats": {"score": {"type": "2_color_scale", "min_color": "#12"}}},
        ["conditional_formats['score']", "'min_color'", "Invalid hex color"],
    ),
    (
        "cf_max_color",
        {"conditional_formats": {"score": {"type": "2_color_scale", "max_color": "nope"}}},
        ["conditional_formats['score']", "'max_color'", "Unknown color"],
    ),
    (
        "cf_icon_type",
        {"conditional_formats": {"score": {"type": "icon_set", "icon_type": "bogus"}}},
        ["conditional_formats['score']", "'icon_type'", "Unknown icon_type"],
    ),
    (
        "cf_direction",
        {"conditional_formats": {"score": {"type": "data_bar", "direction": "bogus"}}},
        ["conditional_formats['score']", "'direction'", "Unknown direction"],
    ),
    (
        "cf_nested_format",
        {
            "conditional_formats": {
                "score": {
                    "type": "cell",
                    "criteria": "gt",
                    "value": 1,
                    "format": {"bg_color": "nope"},
                }
            }
        },
        ["conditional_formats['score']", "'bg_color'", "Unknown color"],
    ),
    # --- charts, validations ------------------------------------------------
    (
        "chart_type",
        {"charts": {"D2": {"type": "bogus"}}},
        ["charts['D2']", "'type'", "Unknown chart type"],
    ),
    (
        "chart_cell_ref",
        {"charts": {"": {"type": "column", "values": "A2:A3"}}},
        ["charts['']", "Empty cell reference"],
    ),
    (
        "validation_type",
        {"validations": {"a": {"type": "bogus"}}},
        ["validations['a']", "'type'", "Unknown validation type"],
    ),
    # --- format dicts: the three colour keys must be told apart --------------
    (
        "header_bg_color",
        {"header_format": {"bg_color": "nope"}},
        ["header_format", "'bg_color'", "Unknown color"],
    ),
    (
        "header_font_color",
        {"header_format": {"font_color": "nope"}},
        ["header_format", "'font_color'", "Unknown color"],
    ),
    (
        "header_border_color",
        {"header_format": {"border_color": "nope"}},
        ["header_format", "'border_color'", "Unknown color"],
    ),
    (
        "header_border_style",
        {"header_format": {"border": "bogus"}},
        ["header_format", "'border'", "Unknown border style"],
    ),
    (
        "header_border_side_style",
        {"header_format": {"border_top": "bogus"}},
        ["header_format", "'border_top'", "Unknown border style"],
    ),
    (
        "header_align_horizontal",
        {"header_format": {"align_horizontal": "bogus"}},
        ["header_format", "'align_horizontal'", "Unknown horizontal alignment"],
    ),
    (
        "header_align_vertical",
        {"header_format": {"align_vertical": "bogus"}},
        ["header_format", "'align_vertical'", "Unknown vertical alignment"],
    ),
    (
        "column_format_align",
        {"column_formats": {"a": {"align_horizontal": "bogus"}}},
        ["column_formats['a']", "'align_horizontal'", "Unknown horizontal alignment"],
    ),
    # --- cell-ref keyed options: the ref alone does not say which option ----
    (
        "rich_text_format",
        {"rich_text": {"A1": [("x", {"font_color": "nope"})]}},
        ["rich_text['A1']", "'font_color'", "Unknown color"],
    ),
    (
        "rich_text_cell_ref",
        {"rich_text": {"": [("x", None)]}},
        ["rich_text['']", "Empty cell reference"],
    ),
    (
        "comment_cell_ref",
        {"comments": {"": "hi"}},
        ["comments['']", "Empty cell reference"],
    ),
    (
        "hyperlink_cell_ref",
        {"hyperlinks": [("", "https://example.com")]},
        ["hyperlinks['']", "Empty cell reference"],
    ),
    (
        "merged_range_ref",
        {"merged_ranges": [("nope", "T", None)]},
        ["merged_ranges['nope']", "Invalid cell range"],
    ),
    (
        "merged_range_format",
        {"merged_ranges": [("A1:B1", "T", {"bg_color": "nope"})]},
        ["merged_ranges['A1:B1']", "'bg_color'", "Unknown color"],
    ),
    (
        "image_cell_ref",
        {"images": {"": "/nonexistent.png"}},
        ["images['']", "Empty cell reference"],
    ),
    (
        "checkbox_cell_ref",
        {"checkboxes": {"": True}},
        ["checkboxes['']", "Empty cell reference"],
    ),
    (
        "textbox_cell_ref",
        {"textboxes": {"": "T"}},
        ["textboxes['']", "Empty cell reference"],
    ),
    (
        "textbox_font_not_a_dict",
        {"textboxes": {"B2": {"text": "T", "font": 5}}},
        ["textboxes['B2']", "'font'", "must be a dict", "got int"],
    ),
    (
        "textbox_font_color",
        {"textboxes": {"B2": {"text": "T", "font": {"color": "nope"}}}},
        ["textboxes['B2']", "'font'", "'color'", "Unknown color"],
    ),
    (
        "sparkline_cell_ref",
        {"sparklines": {"": {"range": "Sheet1!A2:A3"}}},
        ["sparklines['']", "Empty cell reference"],
    ),
    (
        "sparkline_range",
        {"sparklines": {"nope:": {"range": "Sheet1!A2:A3"}}},
        ["sparklines['nope:']", "Invalid cell reference"],
    ),
    # --- cells: parsed at extract time, so the prefix comes from extract.rs ---
    (
        "cells_align_horizontal",
        {"cells": {"C1": {"value": 1}, "D1": {"value": 1, "align_horizontal": "bogus"}}},
        ["cells['D1']", "'align_horizontal'", "Unknown horizontal alignment 'bogus'"],
    ),
    (
        "cells_align_vertical",
        {"cells": {"D1": {"value": 1, "align_vertical": "bogus"}}},
        ["cells['D1']", "'align_vertical'", "Unknown vertical alignment 'bogus'"],
    ),
]


class TestFailuresNameTheOptionThatCausedThem:
    """Each parametrized case is one call site that used to answer anonymously."""

    @pytest.mark.parametrize(
        ("kwargs", "expected"),
        [pytest.param(k, e, id=label) for label, k, e in CASES],
    )
    def test_the_message_identifies_the_dict_entry(
        self, tmp_xlsx: str, kwargs: dict[str, Any], expected: list[str]
    ) -> None:
        """The option -- and the key inside it -- appear in the message."""
        with pytest.raises(xlsxturbo.XlsxTurboError) as exc_info:
            xlsxturbo.df_to_xlsx(FRAME, tmp_xlsx, **kwargs)
        message = str(exc_info.value)
        missing = [part for part in expected if part not in message]
        assert not missing, f"message {message!r} is missing {missing}"

    def test_a_valid_call_still_writes(self, tmp_xlsx: str) -> None:
        """The control.

        Every case above is a failure, so a screen that rejected *everything*
        would pass all of them. This is the case that says the options in that
        table are otherwise well formed and the calls do go through.
        """
        xlsxturbo.df_to_xlsx(
            FRAME,
            tmp_xlsx,
            header_format={
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": "thin",
                "align_horizontal": "center",
            },
            conditional_formats={
                "score": {"type": "2_color_scale", "min_color": "#FFFFFF"}
            },
            charts={"D2": {"type": "column", "values": "Sheet1!A2:A3"}},
            validations={"a": {"type": "whole_number", "min": 0, "max": 9}},
            rich_text={"F1": [("x", {"font_color": "red"})]},
            comments={"G1": "hi"},
            merged_ranges=[("H1:I1", "T", {"bg_color": "yellow"})],
            textboxes={"J2": {"text": "T", "font": {"color": "red"}}},
        )


class TestColumnWidthFailuresNameTheColumn:
    """``column_widths`` shared one message across five call sites.

    Not reachable from Python today: ``extract_column_widths`` validates every
    key against Excel's column range first, so ``set_column_width`` is only ever
    handed a value it accepts. The five sites still had to be told apart, and
    the coverage that exists for them is structural rather than behavioural --
    said here rather than left for a reader to discover by mutating one and
    finding nothing goes red.
    """

    def test_a_width_out_of_the_data_range_is_still_applied(
        self, tmp_xlsx: str
    ) -> None:
        """The control for the helper the five sites now share.

        It exercises the out-of-range branch (`"9"` names a column past the two
        the frame has), which is the one whose loop the other two skip.
        """
        xlsxturbo.df_to_xlsx(FRAME, tmp_xlsx, column_widths={"0": 30.0, "9": 12.0})
