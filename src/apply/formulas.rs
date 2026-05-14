//! Formula column application helpers.

use indexmap::IndexMap;
use rust_xlsxwriter::{Format, Worksheet};

/// Apply formula columns to worksheet
/// Formula templates can use {row} which is replaced with the actual row number (1-based)
pub(crate) fn apply_formula_columns(
    worksheet: &mut Worksheet,
    formula_columns: &IndexMap<String, String>,
    start_col: u16,
    data_start_row: u32,
    data_end_row: u32,
    include_header: bool,
    header_format: Option<&Format>,
) -> Result<u16, String> {
    let mut col_offset = 0u16;

    for (col_name, formula_template) in formula_columns {
        let col_idx = start_col
            .checked_add(col_offset)
            .ok_or("Formula column index exceeds u16 limit")?;

        // Write header for formula column (only when headers are enabled)
        if include_header {
            if let Some(fmt) = header_format {
                worksheet
                    .write_string_with_format(0, col_idx, col_name, fmt)
                    .map_err(|e| format!("Failed to write formula column header: {}", e))?;
            } else {
                worksheet
                    .write_string(0, col_idx, col_name)
                    .map_err(|e| format!("Failed to write formula column header: {}", e))?;
            }
        }

        // Write formula for each data row
        for row in data_start_row..=data_end_row {
            // Replace {row} with actual row number (Excel is 1-based)
            let excel_row = row + 1; // Convert 0-based to 1-based
            let formula = formula_template.replace("{row}", &excel_row.to_string());

            worksheet
                .write_formula(row, col_idx, formula.as_str())
                .map_err(|e| format!("Failed to write formula at row {}: {}", row, e))?;
        }

        col_offset = col_offset
            .checked_add(1)
            .ok_or("Formula column count exceeds u16 limit")?;
    }

    Ok(col_offset)
}
