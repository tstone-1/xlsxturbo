//! Cell annotations, hyperlinks, and merged ranges.

use crate::parse::{parse_cell_range, parse_cell_ref, parse_header_format, WithOptionContext};
use crate::types::{Comment, Hyperlink, MergedRange};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Note, Worksheet};

/// Apply merged ranges to worksheet
pub(crate) fn apply_merged_ranges(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    merged_ranges: &[MergedRange],
) -> Result<(), String> {
    for (range_str, text, format_dict) in merged_ranges {
        let context = format!("merged_ranges['{}']", range_str);
        let (first_row, first_col, last_row, last_col) =
            parse_cell_range(range_str).in_option(&context)?;

        // Center-aligned unless the caller gave a format. Resolved to one value
        // before the merge rather than in two arms of an if/else, which differed
        // only in which format they passed.
        let format = match format_dict {
            Some(fmt_dict) => parse_header_format(py, fmt_dict, &context)?,
            None => Format::new().set_align(rust_xlsxwriter::FormatAlign::Center),
        };

        worksheet
            .merge_range(first_row, first_col, last_row, last_col, text, &format)
            .map_err(|e| format!("Failed to merge range '{}': {}", range_str, e))?;
    }

    Ok(())
}

/// Apply hyperlinks to worksheet
pub(crate) fn apply_hyperlinks(
    worksheet: &mut Worksheet,
    hyperlinks: &[Hyperlink],
) -> Result<(), String> {
    for (cell_ref, url, display_text) in hyperlinks {
        let (row, col) =
            parse_cell_ref(cell_ref).in_option(&format!("hyperlinks['{}']", cell_ref))?;

        if let Some(text) = display_text {
            worksheet
                .write_url_with_text(row, col, url.as_str(), text.as_str())
                .map_err(|e| format!("Failed to write hyperlink at '{}': {}", cell_ref, e))?;
        } else {
            worksheet
                .write_url(row, col, url.as_str())
                .map_err(|e| format!("Failed to write hyperlink at '{}': {}", cell_ref, e))?;
        }
    }

    Ok(())
}

/// Apply comments/notes to worksheet
pub(crate) fn apply_comments(
    worksheet: &mut Worksheet,
    comments: &IndexMap<String, Comment>,
) -> Result<(), String> {
    for (cell_ref, (text, author)) in comments {
        let (row, col) =
            parse_cell_ref(cell_ref).in_option(&format!("comments['{}']", cell_ref))?;

        let mut note = Note::new(text);
        if let Some(auth) = author {
            note = note.set_author(auth);
        }

        worksheet
            .insert_note(row, col, &note)
            .map_err(|e| format!("Failed to insert note at '{}': {}", cell_ref, e))?;
    }

    Ok(())
}
