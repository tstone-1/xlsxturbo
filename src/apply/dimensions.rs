//! Column width application helpers.

use rust_xlsxwriter::Worksheet;
use std::collections::HashMap;

/// Apply column widths to worksheet, supporting '_all' global cap
pub(crate) fn apply_column_widths(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
) -> Result<(), String> {
    let global_width = widths.get("_all").copied();

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        // Specific column overrides '_all'
        if let Some(width) = widths.get(&col_key) {
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else if let Some(width) = global_width {
            worksheet
                .set_column_width(col_idx, width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        }
    }
    Ok(())
}

/// Apply column widths with autofit and cap: autofit each column to content, then cap at '_all'.
/// Uses pre-computed content widths to apply min(autofit, cap) per column.
///
/// Caller must ensure this is NOT called in constant_memory mode (autofit is unsupported).
pub(crate) fn apply_column_widths_with_autofit_cap(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
    content_widths: &[f64],
) -> Result<(), String> {
    let global_cap = widths.get("_all").copied().unwrap_or(f64::MAX);

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        if let Some(width) = widths.get(&col_key) {
            // Specific width overrides autofit and cap
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else {
            // Autofit capped at '_all'
            let autofit_width = content_widths
                .get(col_idx as usize)
                .copied()
                .unwrap_or(8.43); // Excel default
            let capped = autofit_width.min(global_cap);
            worksheet
                .set_column_width(col_idx, capped)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        }
    }
    Ok(())
}
