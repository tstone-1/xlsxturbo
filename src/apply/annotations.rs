//! Cell annotations, hyperlinks, and merged ranges.

use crate::parse::{parse_cell_range, parse_cell_ref, parse_header_format};
use crate::types::{Comment, Hyperlink, MergedRange};
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Note, Worksheet};
use std::collections::HashMap;

/// Apply merged ranges to worksheet
pub(crate) fn apply_merged_ranges(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    merged_ranges: &[MergedRange],
) -> Result<(), String> {
    for (range_str, text, format_dict) in merged_ranges {
        let (first_row, first_col, last_row, last_col) = parse_cell_range(range_str)?;

        // Build format if provided
        let format = if let Some(fmt_dict) = format_dict {
            let parsed = parse_header_format(py, fmt_dict)?;
            Some(parsed)
        } else {
            None
        };

        // Apply merge with or without format
        if let Some(ref fmt) = format {
            worksheet
                .merge_range(first_row, first_col, last_row, last_col, text, fmt)
                .map_err(|e| format!("Failed to merge range '{}': {}", range_str, e))?;
        } else {
            // Create default center-aligned format for merged cells
            let default_fmt = Format::new().set_align(rust_xlsxwriter::FormatAlign::Center);
            worksheet
                .merge_range(first_row, first_col, last_row, last_col, text, &default_fmt)
                .map_err(|e| format!("Failed to merge range '{}': {}", range_str, e))?;
        }
    }

    Ok(())
}

/// Apply hyperlinks to worksheet
pub(crate) fn apply_hyperlinks(
    worksheet: &mut Worksheet,
    hyperlinks: &[Hyperlink],
) -> Result<(), String> {
    for (cell_ref, url, display_text) in hyperlinks {
        let (row, col) = parse_cell_ref(cell_ref)?;

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
    comments: &HashMap<String, Comment>,
) -> Result<(), String> {
    for (cell_ref, (text, author)) in comments {
        let (row, col) = parse_cell_ref(cell_ref)?;

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
