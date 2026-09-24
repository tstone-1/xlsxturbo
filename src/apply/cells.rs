//! Arbitrary cell write application helpers.

use crate::parse::parse_column_format;
use crate::types::CellWrite;
use crate::write::{write_py_value_with_format, CellFormat, DATETIME_NUM_FORMAT, DATE_NUM_FORMAT};
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
        let fmt = cell
            .format
            .as_ref()
            .map(|format| parse_column_format(py, format, &cell.context).map(CellFormat::new))
            .transpose()?;
        write_py_value_with_format(
            worksheet,
            cell.row,
            cell.col,
            cell.value.bind(py),
            &date_format,
            &datetime_format,
            fmt.as_ref(),
        )?;
    }
    Ok(())
}
