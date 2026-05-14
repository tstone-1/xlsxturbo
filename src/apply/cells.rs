//! Arbitrary cell write application helpers.

use crate::convert::{write_py_value_with_format, DATETIME_NUM_FORMAT, DATE_NUM_FORMAT};
use crate::parse::{parse_horizontal_alignment, parse_vertical_alignment};
use crate::types::CellWrite;
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Worksheet};

/// Apply arbitrary cell writes to a worksheet
pub(crate) fn apply_cells(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    cells: &[CellWrite],
) -> Result<(), String> {
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    for cell in cells {
        let value = cell.value.bind(py);
        let has_formatting = cell.num_format.is_some()
            || cell.align_horizontal.is_some()
            || cell.align_vertical.is_some()
            || cell.wrap_text;
        let fmt = if has_formatting {
            let mut f = Format::new();
            if let Some(nf) = &cell.num_format {
                f = f.set_num_format(nf);
            }
            if let Some(ah) = &cell.align_horizontal {
                f = f.set_align(parse_horizontal_alignment(ah)?);
            }
            if let Some(av) = &cell.align_vertical {
                f = f.set_align(parse_vertical_alignment(av)?);
            }
            if cell.wrap_text {
                f = f.set_text_wrap();
            }
            Some(f)
        } else {
            None
        };
        write_py_value_with_format(
            worksheet,
            cell.row,
            cell.col,
            value,
            &date_format,
            &datetime_format,
            fmt.as_ref(),
        )?;
    }
    Ok(())
}
