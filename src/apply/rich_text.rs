//! Rich text application helpers.

use crate::parse::{parse_cell_ref, parse_rich_text_format};
use crate::types::RichTextSegment;
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Worksheet};

/// Apply rich text to worksheet
pub(crate) fn apply_rich_text(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    rich_text: &IndexMap<String, Vec<RichTextSegment>>,
) -> Result<(), String> {
    for (cell_ref, segments) in rich_text {
        let (row, col) = parse_cell_ref(cell_ref)?;
        let context = format!("rich_text['{}']", cell_ref);

        // Build formats and strings separately
        let mut formats: Vec<Format> = Vec::new();
        let mut texts: Vec<String> = Vec::new();

        for (text, format_dict) in segments {
            if let Some(fmt_dict) = format_dict {
                let format = parse_rich_text_format(py, fmt_dict, &context)?;
                formats.push(format);
            } else {
                formats.push(Format::new());
            }
            texts.push(text.clone());
        }

        // Create the segments as tuples of (&Format, &str)
        let rich_segments: Vec<(&Format, &str)> = formats
            .iter()
            .zip(texts.iter())
            .map(|(f, t)| (f, t.as_str()))
            .collect();

        if !rich_segments.is_empty() {
            worksheet
                .write_rich_string(row, col, &rich_segments)
                .map_err(|e| format!("Failed to write rich text at '{}': {}", cell_ref, e))?;
        }
    }

    Ok(())
}
