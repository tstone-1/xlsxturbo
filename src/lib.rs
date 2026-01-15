//! xlsxturbo - High-performance Excel writer with automatic type detection
//!
//! This library provides fast DataFrame and CSV to Excel conversion:
//! - Integers and floats → Excel numbers
//! - Booleans (true/false) → Excel booleans
//! - Dates → Excel dates
//! - Datetimes → Excel datetimes
//! - NaN/Inf/None → Empty cells
//! - Everything else → Strings
//!
//! Supports pandas DataFrames, polars DataFrames, and CSV files.

use chrono::Timelike;
use csv::ReaderBuilder;
use indexmap::IndexMap;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyString};
use rayon::prelude::*;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatDataBar,
    ConditionalFormatDataBarDirection, ConditionalFormatIconSet, ConditionalFormatIconType, Format,
    Table, TableStyle, Workbook, Worksheet, XlsxError,
};
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

/// Date formats by locale/order preference
/// ISO formats (YYYY-MM-DD) are always tried first as they're unambiguous
const DATE_PATTERNS_ISO: &[&str] = &[
    "%Y-%m-%d", // 2024-01-15
    "%Y/%m/%d", // 2024/01/15
];

/// European date formats (day first): DD-MM-YYYY
const DATE_PATTERNS_DMY: &[&str] = &[
    "%d-%m-%Y", // 15-01-2024
    "%d/%m/%Y", // 15/01/2024
];

/// US date formats (month first): MM-DD-YYYY
const DATE_PATTERNS_MDY: &[&str] = &[
    "%m-%d-%Y", // 01-15-2024
    "%m/%d/%Y", // 01/15/2024
];

/// Date order preference for ambiguous dates like 01-02-2024
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum DateOrder {
    /// Year-Month-Day first, then Day-Month-Year, then Month-Day-Year (default)
    #[default]
    Auto,
    /// US format: Month-Day-Year (01-02-2024 = January 2)
    MDY,
    /// European format: Day-Month-Year (01-02-2024 = February 1)
    DMY,
}

impl DateOrder {
    /// Parse from string, returns None for invalid input
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "auto" => Some(DateOrder::Auto),
            "mdy" | "us" => Some(DateOrder::MDY),
            "dmy" | "eu" | "european" => Some(DateOrder::DMY),
            _ => None,
        }
    }

    /// Get date patterns in order of preference
    fn patterns(&self) -> Vec<&'static str> {
        let mut patterns = Vec::with_capacity(6);
        // ISO formats are always first (unambiguous)
        patterns.extend_from_slice(DATE_PATTERNS_ISO);
        match self {
            DateOrder::Auto | DateOrder::DMY => {
                patterns.extend_from_slice(DATE_PATTERNS_DMY);
                patterns.extend_from_slice(DATE_PATTERNS_MDY);
            }
            DateOrder::MDY => {
                patterns.extend_from_slice(DATE_PATTERNS_MDY);
                patterns.extend_from_slice(DATE_PATTERNS_DMY);
            }
        }
        patterns
    }
}

/// Datetime formats we recognize
const DATETIME_PATTERNS: &[&str] = &[
    "%Y-%m-%dT%H:%M:%S",    // ISO 8601
    "%Y-%m-%d %H:%M:%S",    // Common format
    "%Y-%m-%dT%H:%M:%S%.f", // ISO 8601 with fractional seconds
    "%Y-%m-%d %H:%M:%S%.f", // With fractional seconds
];

/// Represents the detected type of a cell value
#[derive(Debug, Clone)]
enum CellValue {
    Empty,
    Integer(i64),
    Float(f64),
    Boolean(bool),
    Date(f64),     // Excel serial date
    DateTime(f64), // Excel serial datetime
    String(String),
}

/// Type alias for merged range tuple: (range_str, text, optional format_dict)
type MergedRange = (String, String, Option<HashMap<String, Py<PyAny>>>);

/// Type alias for hyperlink tuple: (cell_ref, url, optional display_text)
type Hyperlink = (String, String, Option<String>);

/// Per-sheet configuration options (all optional, defaults to global settings)
#[derive(Debug, Default)]
struct SheetConfig {
    header: Option<bool>,
    autofit: Option<bool>,
    table_style: Option<Option<String>>, // None = use default, Some(None) = explicitly no style
    freeze_panes: Option<bool>,
    column_widths: Option<HashMap<String, f64>>, // Keys: "0", "1", "_all" for global cap
    table_name: Option<String>,
    header_format: Option<HashMap<String, Py<PyAny>>>,
    row_heights: Option<HashMap<u32, f64>>,
    column_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>, // Pattern -> format dict (ordered)
    conditional_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>, // Column/pattern -> conditional format config (ordered)
    formula_columns: Option<IndexMap<String, String>>, // Column name -> formula template (ordered)
    merged_ranges: Option<Vec<MergedRange>>,           // (range, text, format)
    hyperlinks: Option<Vec<Hyperlink>>,                // (cell, url, optional display_text)
}

/// Extract sheet info from a Python tuple (supports both 2-tuple and 3-tuple formats)
/// 2-tuple: (df, sheet_name)
/// 3-tuple: (df, sheet_name, options_dict)
fn extract_sheet_info<'py>(
    sheet_tuple: &Bound<'py, PyAny>,
) -> PyResult<(Bound<'py, PyAny>, String, SheetConfig)> {
    let len: usize = sheet_tuple.len()?;

    if len < 2 {
        return Err(pyo3::exceptions::PyValueError::new_err(
            "Sheet tuple must have at least 2 elements: (df, sheet_name)",
        ));
    }

    let df = sheet_tuple.get_item(0)?;
    let sheet_name: String = sheet_tuple.get_item(1)?.extract()?;

    let config = if len >= 3 {
        let opts = sheet_tuple.get_item(2)?;
        let mut config = SheetConfig::default();

        // Extract optional fields from the dict
        if let Ok(val) = opts.get_item("header") {
            if !val.is_none() {
                config.header = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("autofit") {
            if !val.is_none() {
                config.autofit = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("table_style") {
            // Handle both None and string values
            if val.is_none() {
                config.table_style = Some(None); // Explicitly no style
            } else {
                config.table_style = Some(Some(val.extract()?));
            }
        }
        if let Ok(val) = opts.get_item("freeze_panes") {
            if !val.is_none() {
                config.freeze_panes = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("column_widths") {
            if !val.is_none() {
                // Support both integer keys {0: 20} and string keys {"_all": 50}
                let mut widths: HashMap<String, f64> = HashMap::new();
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    for (k, v) in dict.iter() {
                        let key_str = if let Ok(i) = k.extract::<i64>() {
                            i.to_string()
                        } else {
                            k.extract::<String>()?
                        };
                        widths.insert(key_str, v.extract()?);
                    }
                }
                if !widths.is_empty() {
                    config.column_widths = Some(widths);
                }
            }
        }
        if let Ok(val) = opts.get_item("row_heights") {
            if !val.is_none() {
                config.row_heights = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("table_name") {
            if !val.is_none() {
                config.table_name = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("header_format") {
            if !val.is_none() {
                let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    for (k, v) in dict.iter() {
                        fmt.insert(k.extract()?, v.unbind());
                    }
                }
                if !fmt.is_empty() {
                    config.header_format = Some(fmt);
                }
            }
        }
        if let Ok(val) = opts.get_item("column_formats") {
            if !val.is_none() {
                if let Ok(outer_dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> =
                        IndexMap::new();
                    for (pattern, fmt_dict) in outer_dict.iter() {
                        let pattern_str: String = pattern.extract()?;
                        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
                            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                            for (k, v) in inner_dict.iter() {
                                fmt.insert(k.extract()?, v.unbind());
                            }
                            col_fmts.insert(pattern_str, fmt);
                        }
                    }
                    if !col_fmts.is_empty() {
                        config.column_formats = Some(col_fmts);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("conditional_formats") {
            if !val.is_none() {
                if let Ok(outer_dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut cond_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> =
                        IndexMap::new();
                    for (col_name, fmt_dict) in outer_dict.iter() {
                        let col_str: String = col_name.extract()?;
                        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
                            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                            for (k, v) in inner_dict.iter() {
                                fmt.insert(k.extract()?, v.unbind());
                            }
                            cond_fmts.insert(col_str, fmt);
                        }
                    }
                    if !cond_fmts.is_empty() {
                        config.conditional_formats = Some(cond_fmts);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("formula_columns") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut formulas: IndexMap<String, String> = IndexMap::new();
                    for (col_name, formula) in dict.iter() {
                        let col_str: String = col_name.extract()?;
                        let formula_str: String = formula.extract()?;
                        formulas.insert(col_str, formula_str);
                    }
                    if !formulas.is_empty() {
                        config.formula_columns = Some(formulas);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("merged_ranges") {
            if !val.is_none() {
                if let Ok(list) = val.cast::<pyo3::types::PyList>() {
                    let extracted = extract_merged_ranges(list)?;
                    if !extracted.is_empty() {
                        config.merged_ranges = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("hyperlinks") {
            if !val.is_none() {
                if let Ok(list) = val.cast::<pyo3::types::PyList>() {
                    let extracted = extract_hyperlinks(list)?;
                    if !extracted.is_empty() {
                        config.hyperlinks = Some(extracted);
                    }
                }
            }
        }

        config
    } else {
        SheetConfig::default()
    };

    Ok((df, sheet_name, config))
}

/// Parse a table style string to TableStyle enum.
/// Supports: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
fn parse_table_style(style: &str) -> Result<TableStyle, String> {
    match style {
        "None" => Ok(TableStyle::None),
        "Light1" => Ok(TableStyle::Light1),
        "Light2" => Ok(TableStyle::Light2),
        "Light3" => Ok(TableStyle::Light3),
        "Light4" => Ok(TableStyle::Light4),
        "Light5" => Ok(TableStyle::Light5),
        "Light6" => Ok(TableStyle::Light6),
        "Light7" => Ok(TableStyle::Light7),
        "Light8" => Ok(TableStyle::Light8),
        "Light9" => Ok(TableStyle::Light9),
        "Light10" => Ok(TableStyle::Light10),
        "Light11" => Ok(TableStyle::Light11),
        "Light12" => Ok(TableStyle::Light12),
        "Light13" => Ok(TableStyle::Light13),
        "Light14" => Ok(TableStyle::Light14),
        "Light15" => Ok(TableStyle::Light15),
        "Light16" => Ok(TableStyle::Light16),
        "Light17" => Ok(TableStyle::Light17),
        "Light18" => Ok(TableStyle::Light18),
        "Light19" => Ok(TableStyle::Light19),
        "Light20" => Ok(TableStyle::Light20),
        "Light21" => Ok(TableStyle::Light21),
        "Medium1" => Ok(TableStyle::Medium1),
        "Medium2" => Ok(TableStyle::Medium2),
        "Medium3" => Ok(TableStyle::Medium3),
        "Medium4" => Ok(TableStyle::Medium4),
        "Medium5" => Ok(TableStyle::Medium5),
        "Medium6" => Ok(TableStyle::Medium6),
        "Medium7" => Ok(TableStyle::Medium7),
        "Medium8" => Ok(TableStyle::Medium8),
        "Medium9" => Ok(TableStyle::Medium9),
        "Medium10" => Ok(TableStyle::Medium10),
        "Medium11" => Ok(TableStyle::Medium11),
        "Medium12" => Ok(TableStyle::Medium12),
        "Medium13" => Ok(TableStyle::Medium13),
        "Medium14" => Ok(TableStyle::Medium14),
        "Medium15" => Ok(TableStyle::Medium15),
        "Medium16" => Ok(TableStyle::Medium16),
        "Medium17" => Ok(TableStyle::Medium17),
        "Medium18" => Ok(TableStyle::Medium18),
        "Medium19" => Ok(TableStyle::Medium19),
        "Medium20" => Ok(TableStyle::Medium20),
        "Medium21" => Ok(TableStyle::Medium21),
        "Medium22" => Ok(TableStyle::Medium22),
        "Medium23" => Ok(TableStyle::Medium23),
        "Medium24" => Ok(TableStyle::Medium24),
        "Medium25" => Ok(TableStyle::Medium25),
        "Medium26" => Ok(TableStyle::Medium26),
        "Medium27" => Ok(TableStyle::Medium27),
        "Medium28" => Ok(TableStyle::Medium28),
        "Dark1" => Ok(TableStyle::Dark1),
        "Dark2" => Ok(TableStyle::Dark2),
        "Dark3" => Ok(TableStyle::Dark3),
        "Dark4" => Ok(TableStyle::Dark4),
        "Dark5" => Ok(TableStyle::Dark5),
        "Dark6" => Ok(TableStyle::Dark6),
        "Dark7" => Ok(TableStyle::Dark7),
        "Dark8" => Ok(TableStyle::Dark8),
        "Dark9" => Ok(TableStyle::Dark9),
        "Dark10" => Ok(TableStyle::Dark10),
        "Dark11" => Ok(TableStyle::Dark11),
        _ => Err(format!(
            "Unknown table_style '{}'. Valid styles: Light1-Light21, Medium1-Medium28, Dark1-Dark11, None",
            style
        )),
    }
}

/// Apply column widths to worksheet, supporting '_all' global cap
fn apply_column_widths(
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

/// Apply column widths with autofit cap: autofit first, then cap columns at '_all' width
fn apply_column_widths_with_autofit_cap(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
    constant_memory: bool,
) -> Result<(), String> {
    // First autofit
    if !constant_memory {
        worksheet.autofit();
    }

    // Then apply specific widths and cap at '_all' if specified
    let global_cap = widths.get("_all").copied();

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        if let Some(width) = widths.get(&col_key) {
            // Specific width overrides autofit completely
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else if let Some(cap) = global_cap {
            // '_all' acts as a cap - only set if current width exceeds cap
            // Since we can't read current width, just set the cap
            worksheet
                .set_column_width(col_idx, cap)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        }
    }
    Ok(())
}

/// Extract column_widths from Python dict, supporting both integer and string keys
fn extract_column_widths(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, f64>> {
    let mut widths: HashMap<String, f64> = HashMap::new();
    for (k, v) in py_dict.iter() {
        let key_str = if let Ok(i) = k.extract::<i64>() {
            i.to_string()
        } else {
            k.extract::<String>()?
        };
        widths.insert(key_str, v.extract()?);
    }
    Ok(widths)
}

/// Extract header_format from Python dict
fn extract_header_format(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
    for (k, v) in py_dict.iter() {
        fmt.insert(k.extract()?, v.unbind());
    }
    Ok(fmt)
}

/// Extract column_formats from Python dict (pattern -> format dict)
/// Uses IndexMap to preserve insertion order from Python dict
fn extract_column_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (pattern, fmt_dict) in py_dict.iter() {
        let pattern_str: String = pattern.extract()?;
        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
            for (k, v) in inner_dict.iter() {
                fmt.insert(k.extract()?, v.unbind());
            }
            col_fmts.insert(pattern_str, fmt);
        }
    }
    Ok(col_fmts)
}

/// Extract conditional_formats from Python dict (column/pattern -> config dict)
/// Uses IndexMap to preserve insertion order for pattern matching (first match wins)
fn extract_conditional_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut cond_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (col_name, fmt_dict) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
            for (k, v) in inner_dict.iter() {
                fmt.insert(k.extract()?, v.unbind());
            }
            cond_fmts.insert(col_str, fmt);
        }
    }
    Ok(cond_fmts)
}

/// Extract formula_columns from Python dict (column name -> formula template)
/// Uses IndexMap to preserve column order
fn extract_formula_columns(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, String>> {
    let mut formulas: IndexMap<String, String> = IndexMap::new();
    for (col_name, formula) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        let formula_str: String = formula.extract()?;
        formulas.insert(col_str, formula_str);
    }
    Ok(formulas)
}

/// Apply formula columns to worksheet
/// Formula templates can use {row} which is replaced with the actual row number (1-based)
fn apply_formula_columns(
    worksheet: &mut Worksheet,
    formula_columns: &IndexMap<String, String>,
    start_col: u16,
    data_start_row: u32,
    data_end_row: u32,
    header_format: Option<&Format>,
) -> Result<u16, String> {
    let mut col_offset = 0u16;

    for (col_name, formula_template) in formula_columns {
        let col_idx = start_col + col_offset;

        // Write header for formula column
        if let Some(fmt) = header_format {
            worksheet
                .write_string_with_format(0, col_idx, col_name, fmt)
                .map_err(|e| format!("Failed to write formula column header: {}", e))?;
        } else {
            worksheet
                .write_string(0, col_idx, col_name)
                .map_err(|e| format!("Failed to write formula column header: {}", e))?;
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

        col_offset += 1;
    }

    Ok(col_offset)
}

/// Parse a cell reference like "A1" into (row, col) - 0-based
fn parse_cell_ref(cell_ref: &str) -> Result<(u32, u16), String> {
    let cell_ref = cell_ref.trim().to_uppercase();
    if cell_ref.is_empty() {
        return Err("Empty cell reference".to_string());
    }

    // Find where letters end and numbers begin
    let col_end = cell_ref
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .count();
    if col_end == 0 {
        return Err(format!(
            "Invalid cell reference '{}': no column letters",
            cell_ref
        ));
    }

    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    if row_str.is_empty() {
        return Err(format!(
            "Invalid cell reference '{}': no row number",
            cell_ref
        ));
    }

    // Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, etc.)
    let col: u16 = col_str
        .chars()
        .fold(0u16, |acc, c| acc * 26 + (c as u16 - 'A' as u16 + 1))
        .saturating_sub(1);

    // Parse row number (Excel rows are 1-based, so must be >= 1)
    let row_1based: u32 = row_str
        .parse::<u32>()
        .map_err(|_| format!("Invalid row number in cell reference '{}'", cell_ref))?;

    if row_1based == 0 {
        return Err(format!(
            "Invalid cell reference '{}': row number must be >= 1 (Excel rows are 1-based)",
            cell_ref
        ));
    }

    // Convert to 0-based index
    let row = row_1based - 1;

    Ok((row, col))
}

/// Parse a cell range like "A1:D1" into (first_row, first_col, last_row, last_col) - 0-based
fn parse_cell_range(range_str: &str) -> Result<(u32, u16, u32, u16), String> {
    let parts: Vec<&str> = range_str.split(':').collect();
    if parts.len() != 2 {
        return Err(format!(
            "Invalid cell range '{}': expected format 'A1:B2'",
            range_str
        ));
    }

    let (first_row, first_col) = parse_cell_ref(parts[0])?;
    let (last_row, last_col) = parse_cell_ref(parts[1])?;

    Ok((first_row, first_col, last_row, last_col))
}

/// Extract merged_ranges from Python list of tuples
/// Each tuple: (range_str, text) or (range_str, text, format_dict)
fn extract_merged_ranges(py_list: &Bound<'_, pyo3::types::PyList>) -> PyResult<Vec<MergedRange>> {
    let mut ranges = Vec::new();

    for item in py_list.iter() {
        let tuple_len = item.len()?;
        if tuple_len < 2 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "merged_ranges tuple must have at least 2 elements: (range, text)",
            ));
        }

        let range_str: String = item.get_item(0)?.extract()?;
        let text: String = item.get_item(1)?.extract()?;

        let format_dict = if tuple_len >= 3 {
            let fmt_item = item.get_item(2)?;
            if !fmt_item.is_none() {
                if let Ok(dict) = fmt_item.cast::<pyo3::types::PyDict>() {
                    let mut fmt_map: HashMap<String, Py<PyAny>> = HashMap::new();
                    for (k, v) in dict.iter() {
                        let key: String = k.extract()?;
                        fmt_map.insert(key, v.unbind());
                    }
                    Some(fmt_map)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        ranges.push((range_str, text, format_dict));
    }

    Ok(ranges)
}

/// Apply merged ranges to worksheet
fn apply_merged_ranges(
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

/// Extract hyperlinks from Python list of tuples
/// Each tuple: (cell_ref, url) or (cell_ref, url, display_text)
fn extract_hyperlinks(py_list: &Bound<'_, pyo3::types::PyList>) -> PyResult<Vec<Hyperlink>> {
    let mut links = Vec::new();

    for item in py_list.iter() {
        let tuple_len = item.len()?;
        if tuple_len < 2 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "hyperlinks tuple must have at least 2 elements: (cell_ref, url)",
            ));
        }

        let cell_ref: String = item.get_item(0)?.extract()?;
        let url: String = item.get_item(1)?.extract()?;

        let display_text = if tuple_len >= 3 {
            let text_item = item.get_item(2)?;
            if !text_item.is_none() {
                Some(text_item.extract()?)
            } else {
                None
            }
        } else {
            None
        };

        links.push((cell_ref, url, display_text));
    }

    Ok(links)
}

/// Apply hyperlinks to worksheet
fn apply_hyperlinks(worksheet: &mut Worksheet, hyperlinks: &[Hyperlink]) -> Result<(), String> {
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

/// Parse icon type string to ConditionalFormatIconType
fn parse_icon_type(icon_type: &str) -> Result<ConditionalFormatIconType, String> {
    match icon_type.to_lowercase().as_str() {
        "3_arrows" | "3arrows" => Ok(ConditionalFormatIconType::ThreeArrows),
        "3_arrows_gray" | "3arrowsgray" => Ok(ConditionalFormatIconType::ThreeArrowsGray),
        "3_flags" | "3flags" => Ok(ConditionalFormatIconType::ThreeFlags),
        "3_traffic_lights" | "3trafficlights" | "traffic_lights" => {
            Ok(ConditionalFormatIconType::ThreeTrafficLights)
        }
        "3_traffic_lights_rimmed" | "3trafficlightsrimmed" => {
            Ok(ConditionalFormatIconType::ThreeTrafficLightsWithRim)
        }
        "3_signs" | "3signs" => Ok(ConditionalFormatIconType::ThreeSigns),
        "3_symbols" | "3symbols" => Ok(ConditionalFormatIconType::ThreeSymbolsCircled),
        "3_symbols_uncircled" | "3symbolsuncircled" => {
            Ok(ConditionalFormatIconType::ThreeSymbols)
        }
        "4_arrows" | "4arrows" => Ok(ConditionalFormatIconType::FourArrows),
        "4_arrows_gray" | "4arrowsgray" => Ok(ConditionalFormatIconType::FourArrowsGray),
        "4_rating" | "4rating" => Ok(ConditionalFormatIconType::FourHistograms),
        "4_traffic_lights" | "4trafficlights" => {
            Ok(ConditionalFormatIconType::FourTrafficLights)
        }
        "5_arrows" | "5arrows" => Ok(ConditionalFormatIconType::FiveArrows),
        "5_arrows_gray" | "5arrowsgray" => Ok(ConditionalFormatIconType::FiveArrowsGray),
        "5_rating" | "5rating" => Ok(ConditionalFormatIconType::FiveHistograms),
        "5_quarters" | "5quarters" => Ok(ConditionalFormatIconType::FiveQuadrants),
        _ => Err(format!(
            "Unknown icon_type '{}'. Valid types: 3_arrows, 3_arrows_gray, 3_flags, 3_traffic_lights, 3_traffic_lights_rimmed, 3_signs, 3_symbols, 3_symbols_uncircled, 4_arrows, 4_arrows_gray, 4_rating, 4_traffic_lights, 5_arrows, 5_arrows_gray, 5_quarters, 5_rating",
            icon_type
        )),
    }
}

/// Apply conditional formats to a worksheet
/// Supports: 2_color_scale, 3_color_scale, data_bar, icon_set
/// Uses IndexMap to preserve pattern order (first match wins for overlapping patterns)
fn apply_conditional_formats(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_start_row: u32,
    data_end_row: u32,
    cond_formats: &IndexMap<String, HashMap<String, Py<PyAny>>>,
) -> Result<(), String> {
    for (col_pattern, config) in cond_formats {
        // Find column index by name (supports exact match or pattern)
        let col_indices: Vec<u16> = columns
            .iter()
            .enumerate()
            .filter(|(_, name)| matches_pattern(name, col_pattern))
            .map(|(idx, _)| idx as u16)
            .collect();

        if col_indices.is_empty() {
            continue; // Skip if no matching columns
        }

        // Get the format type
        let format_type: String = config
            .get("type")
            .ok_or_else(|| format!("conditional_formats['{}']: missing 'type' key", col_pattern))?
            .bind(py)
            .extract()
            .map_err(|e| {
                format!(
                    "conditional_formats['{}']: invalid 'type': {}",
                    col_pattern, e
                )
            })?;

        for col_idx in col_indices {
            match format_type.to_lowercase().as_str() {
                "2_color_scale" | "2colorscale" | "two_color_scale" => {
                    let mut cf = ConditionalFormat2ColorScale::new();

                    // Parse min_color
                    if let Some(min_color_obj) = config.get("min_color") {
                        if let Ok(color_str) = min_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_minimum_color(color);
                        }
                    }

                    // Parse max_color
                    if let Some(max_color_obj) = config.get("max_color") {
                        if let Ok(color_str) = max_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_maximum_color(color);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add 2_color_scale: {}", e))?;
                }

                "3_color_scale" | "3colorscale" | "three_color_scale" => {
                    let mut cf = ConditionalFormat3ColorScale::new();

                    // Parse min_color
                    if let Some(min_color_obj) = config.get("min_color") {
                        if let Ok(color_str) = min_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_minimum_color(color);
                        }
                    }

                    // Parse mid_color
                    if let Some(mid_color_obj) = config.get("mid_color") {
                        if let Ok(color_str) = mid_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_midpoint_color(color);
                        }
                    }

                    // Parse max_color
                    if let Some(max_color_obj) = config.get("max_color") {
                        if let Ok(color_str) = max_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_maximum_color(color);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add 3_color_scale: {}", e))?;
                }

                "data_bar" | "databar" => {
                    let mut cf = ConditionalFormatDataBar::new();

                    // Parse bar_color (fill color)
                    if let Some(color_obj) = config.get("bar_color") {
                        if let Ok(color_str) = color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_fill_color(color);
                        }
                    }

                    // Parse border_color
                    if let Some(color_obj) = config.get("border_color") {
                        if let Ok(color_str) = color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_border_color(color);
                        }
                    }

                    // Parse solid (vs gradient)
                    if let Some(solid_obj) = config.get("solid") {
                        if let Ok(solid) = solid_obj.bind(py).extract::<bool>() {
                            if solid {
                                cf = cf.set_solid_fill(true);
                            }
                        }
                    }

                    // Parse direction
                    if let Some(dir_obj) = config.get("direction") {
                        if let Ok(dir_str) = dir_obj.bind(py).extract::<String>() {
                            let direction = match dir_str.to_lowercase().as_str() {
                                "left_to_right" | "ltr" => {
                                    ConditionalFormatDataBarDirection::LeftToRight
                                }
                                "right_to_left" | "rtl" => {
                                    ConditionalFormatDataBarDirection::RightToLeft
                                }
                                "context" | "" => ConditionalFormatDataBarDirection::Context,
                                _ => {
                                    return Err(format!(
                                        "Unknown direction '{}'. Valid values: left_to_right, right_to_left, context",
                                        dir_str
                                    ));
                                }
                            };
                            cf = cf.set_direction(direction);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add data_bar: {}", e))?;
                }

                "icon_set" | "iconset" => {
                    let mut cf = ConditionalFormatIconSet::new();

                    // Parse icon_type
                    if let Some(icon_obj) = config.get("icon_type") {
                        if let Ok(icon_str) = icon_obj.bind(py).extract::<String>() {
                            let icon_type = parse_icon_type(&icon_str)?;
                            cf = cf.set_icon_type(icon_type);
                        }
                    }

                    // Parse reverse
                    if let Some(rev_obj) = config.get("reverse") {
                        if let Ok(reverse) = rev_obj.bind(py).extract::<bool>() {
                            if reverse {
                                cf = cf.reverse_icons(true);
                            }
                        }
                    }

                    // Parse icons_only (hide numbers, show only icons)
                    if let Some(icons_only_obj) = config.get("icons_only") {
                        if let Ok(icons_only) = icons_only_obj.bind(py).extract::<bool>() {
                            if icons_only {
                                cf = cf.show_icons_only(true);
                            }
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add icon_set: {}", e))?;
                }

                _ => {
                    return Err(format!(
                        "Unknown conditional format type '{}'. Valid types: 2_color_scale, 3_color_scale, data_bar, icon_set",
                        format_type
                    ));
                }
            }
        }
    }

    Ok(())
}

/// Sanitize table name for Excel (alphanumeric + underscore, must start with letter/underscore)
fn sanitize_table_name(name: &str) -> String {
    let mut sanitized: String = name
        .chars()
        .map(|c| {
            if c.is_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect();

    // Must start with letter or underscore
    if sanitized.chars().next().is_none_or(|c| c.is_ascii_digit()) {
        sanitized = format!("_{}", sanitized);
    }

    // Max 255 chars
    sanitized.truncate(255);
    sanitized
}

/// Parse color string (hex #RRGGBB or named color) to u32
fn parse_color(color_str: &str) -> Result<u32, String> {
    let color = color_str.trim();
    if let Some(hex) = color.strip_prefix('#') {
        if hex.len() != 6 {
            return Err(format!(
                "Invalid hex color '{}': expected 6 characters after #, got {}",
                color,
                hex.len()
            ));
        }
        u32::from_str_radix(hex, 16).map_err(|_| format!("Invalid hex color: {}", color))
    } else {
        match color.to_lowercase().as_str() {
            "white" => Ok(0xFFFFFF),
            "black" => Ok(0x000000),
            "red" => Ok(0xFF0000),
            "green" => Ok(0x00FF00),
            "blue" => Ok(0x0000FF),
            "yellow" => Ok(0xFFFF00),
            "cyan" => Ok(0x00FFFF),
            "magenta" => Ok(0xFF00FF),
            "gray" | "grey" => Ok(0x808080),
            "silver" => Ok(0xC0C0C0),
            "orange" => Ok(0xFFA500),
            "purple" => Ok(0x800080),
            "navy" => Ok(0x000080),
            "teal" => Ok(0x008080),
            "maroon" => Ok(0x800000),
            _ => Err(format!("Unknown color: {}", color)),
        }
    }
}

/// Parse header format dictionary into rust_xlsxwriter Format
fn parse_header_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    let mut format = Format::new();

    if let Some(bold_obj) = fmt_dict.get("bold") {
        let bold: bool = bold_obj.bind(py).extract().unwrap_or(false);
        if bold {
            format = format.set_bold();
        }
    }

    if let Some(italic_obj) = fmt_dict.get("italic") {
        let italic: bool = italic_obj.bind(py).extract().unwrap_or(false);
        if italic {
            format = format.set_italic();
        }
    }

    if let Some(bg_obj) = fmt_dict.get("bg_color") {
        if let Ok(color_str) = bg_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_background_color(color);
        }
    }

    if let Some(font_obj) = fmt_dict.get("font_color") {
        if let Ok(color_str) = font_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_font_color(color);
        }
    }

    if let Some(size_obj) = fmt_dict.get("font_size") {
        if let Ok(size) = size_obj.bind(py).extract::<f64>() {
            format = format.set_font_size(size);
        }
    }

    if let Some(underline_obj) = fmt_dict.get("underline") {
        let underline: bool = underline_obj.bind(py).extract().unwrap_or(false);
        if underline {
            format = format.set_underline(rust_xlsxwriter::FormatUnderline::Single);
        }
    }

    Ok(format)
}

/// Check if a column name matches a wildcard pattern.
/// Supports: "prefix*", "*suffix", "*contains*", or exact match
fn matches_pattern(column_name: &str, pattern: &str) -> bool {
    let starts_with_star = pattern.starts_with('*');
    let ends_with_star = pattern.ends_with('*');

    match (starts_with_star, ends_with_star) {
        (true, true) => {
            // *contains* - match substring
            let inner = &pattern[1..pattern.len() - 1];
            column_name.contains(inner)
        }
        (true, false) => {
            // *suffix - match ending
            let suffix = &pattern[1..];
            column_name.ends_with(suffix)
        }
        (false, true) => {
            // prefix* - match beginning
            let prefix = &pattern[..pattern.len() - 1];
            column_name.starts_with(prefix)
        }
        (false, false) => {
            // Exact match
            column_name == pattern
        }
    }
}

/// Parse column format dictionary into rust_xlsxwriter Format
/// Similar to parse_header_format but also supports num_format
fn parse_column_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    let mut format = Format::new();

    if let Some(bold_obj) = fmt_dict.get("bold") {
        let bold: bool = bold_obj.bind(py).extract().unwrap_or(false);
        if bold {
            format = format.set_bold();
        }
    }

    if let Some(italic_obj) = fmt_dict.get("italic") {
        let italic: bool = italic_obj.bind(py).extract().unwrap_or(false);
        if italic {
            format = format.set_italic();
        }
    }

    if let Some(bg_obj) = fmt_dict.get("bg_color") {
        if let Ok(color_str) = bg_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_background_color(color);
        }
    }

    if let Some(font_obj) = fmt_dict.get("font_color") {
        if let Ok(color_str) = font_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_font_color(color);
        }
    }

    if let Some(size_obj) = fmt_dict.get("font_size") {
        if let Ok(size) = size_obj.bind(py).extract::<f64>() {
            format = format.set_font_size(size);
        }
    }

    if let Some(underline_obj) = fmt_dict.get("underline") {
        let underline: bool = underline_obj.bind(py).extract().unwrap_or(false);
        if underline {
            format = format.set_underline(rust_xlsxwriter::FormatUnderline::Single);
        }
    }

    // Support num_format for number formatting (e.g., "0.00000", "#,##0")
    if let Some(num_fmt_obj) = fmt_dict.get("num_format") {
        if let Ok(num_fmt_str) = num_fmt_obj.bind(py).extract::<String>() {
            format = format.set_num_format(&num_fmt_str);
        }
    }

    // Support border (adds thin border around cell)
    if let Some(border_obj) = fmt_dict.get("border") {
        let border: bool = border_obj.bind(py).extract().unwrap_or(false);
        if border {
            format = format.set_border(rust_xlsxwriter::FormatBorder::Thin);
        }
    }

    Ok(format)
}

/// Build a vector of column formats, one for each column.
/// Returns None for columns with no matching pattern.
/// Uses IndexMap to preserve pattern order - first matching pattern wins.
fn build_column_formats(
    py: Python<'_>,
    columns: &[String],
    column_formats: &IndexMap<String, HashMap<String, Py<PyAny>>>,
) -> Result<Vec<Option<Format>>, String> {
    let mut formats = Vec::with_capacity(columns.len());

    for col_name in columns {
        // Find the first matching pattern (order preserved by IndexMap)
        let mut matched_format: Option<Format> = None;
        for (pattern, fmt_dict) in column_formats {
            if matches_pattern(col_name, pattern) {
                matched_format = Some(parse_column_format(py, fmt_dict)?);
                break;
            }
        }
        formats.push(matched_format);
    }

    Ok(formats)
}

/// Parse a string value and detect its type
fn parse_value(value: &str, date_order: DateOrder) -> CellValue {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return CellValue::Empty;
    }

    // Try integer
    if let Ok(int_val) = trimmed.parse::<i64>() {
        return CellValue::Integer(int_val);
    }

    // Try float
    if let Ok(float_val) = trimmed.parse::<f64>() {
        if float_val.is_nan() || float_val.is_infinite() {
            return CellValue::Empty;
        }
        return CellValue::Float(float_val);
    }

    // Try boolean
    if trimmed.eq_ignore_ascii_case("true") {
        return CellValue::Boolean(true);
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return CellValue::Boolean(false);
    }

    // Try datetime (before date, as datetime patterns are more specific)
    for pattern in DATETIME_PATTERNS {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(trimmed, pattern) {
            let excel_date = naive_datetime_to_excel(dt);
            return CellValue::DateTime(excel_date);
        }
    }

    // Try date with locale-aware ordering
    for pattern in date_order.patterns() {
        if let Ok(date) = chrono::NaiveDate::parse_from_str(trimmed, pattern) {
            let excel_date = naive_date_to_excel(date);
            return CellValue::Date(excel_date);
        }
    }

    // Default to string
    CellValue::String(trimmed.to_string())
}

/// Convert NaiveDate to Excel serial date number
fn naive_date_to_excel(date: chrono::NaiveDate) -> f64 {
    // Excel epoch is December 30, 1899 (accounting for the 1900 leap year bug)
    let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let duration = date.signed_duration_since(excel_epoch);
    duration.num_days() as f64
}

/// Convert NaiveDateTime to Excel serial datetime number
fn naive_datetime_to_excel(dt: chrono::NaiveDateTime) -> f64 {
    let date_part = naive_date_to_excel(dt.date());
    let time = dt.time();
    let time_fraction = (time.num_seconds_from_midnight() as f64) / 86400.0;
    date_part + time_fraction
}

/// Write a cell value to the worksheet with appropriate formatting
fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: CellValue,
    date_format: &Format,
    datetime_format: &Format,
) -> Result<(), XlsxError> {
    match value {
        CellValue::Empty => {
            worksheet.write_string(row, col, "")?;
        }
        CellValue::Integer(v) => {
            worksheet.write_number(row, col, v as f64)?;
        }
        CellValue::Float(v) => {
            worksheet.write_number(row, col, v)?;
        }
        CellValue::Boolean(v) => {
            worksheet.write_boolean(row, col, v)?;
        }
        CellValue::Date(v) => {
            worksheet.write_number_with_format(row, col, v, date_format)?;
        }
        CellValue::DateTime(v) => {
            worksheet.write_number_with_format(row, col, v, datetime_format)?;
        }
        CellValue::String(v) => {
            worksheet.write_string(row, col, &v)?;
        }
    }
    Ok(())
}

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// # Arguments
/// * `input_path` - Path to the input CSV file
/// * `output_path` - Path for the output XLSX file
/// * `sheet_name` - Name of the worksheet (default: "Sheet1")
/// * `date_order` - Date parsing order for ambiguous dates (default: Auto)
///
/// # Returns
/// * `Ok((rows, cols))` - Number of rows and columns written
/// * `Err(message)` - Error description if conversion fails
pub fn convert_csv_to_xlsx(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    date_order: DateOrder,
) -> Result<(u32, u16), String> {
    // Open CSV file
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let reader = BufReader::with_capacity(1024 * 1024, file);
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(reader);

    // Create workbook and worksheet
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    let mut row_count: u32 = 0;
    let mut col_count: u16 = 0;

    // Process records
    for result in csv_reader.records() {
        let record = result.map_err(|e| format!("CSV parse error at row {}: {}", row_count, e))?;
        let num_cols = record.len() as u16;
        if num_cols > col_count {
            col_count = num_cols;
        }

        for (col_idx, value) in record.iter().enumerate() {
            let cell_value = parse_value(value, date_order);
            write_cell(
                worksheet,
                row_count,
                col_idx as u16,
                cell_value,
                &date_format,
                &datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_count, col_idx, e))?;
        }

        row_count += 1;
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

/// Convert a CSV file to XLSX format using parallel processing.
///
/// This version reads all records into memory, parses them in parallel,
/// then writes sequentially. Best for large files with complex type detection.
pub fn convert_csv_to_xlsx_parallel(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    date_order: DateOrder,
) -> Result<(u32, u16), String> {
    // Open CSV file
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let reader = BufReader::with_capacity(1024 * 1024, file);
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(reader);

    // Read all records into memory
    let records: Vec<Vec<String>> = csv_reader
        .records()
        .enumerate()
        .map(|(row_idx, result)| {
            result
                .map(|record| record.iter().map(|s| s.to_string()).collect())
                .map_err(|e| format!("CSV parse error at row {}: {}", row_idx, e))
        })
        .collect::<Result<Vec<_>, _>>()?;

    let row_count = records.len() as u32;
    let col_count = records.iter().map(|r| r.len()).max().unwrap_or(0) as u16;

    // Parse all values in parallel
    let parsed_rows: Vec<Vec<CellValue>> = records
        .par_iter()
        .map(|row| {
            row.iter()
                .map(|value| parse_value(value, date_order))
                .collect()
        })
        .collect();

    // Create workbook and worksheet
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Write parsed values sequentially
    for (row_idx, row) in parsed_rows.into_iter().enumerate() {
        for (col_idx, cell_value) in row.into_iter().enumerate() {
            write_cell(
                worksheet,
                row_idx as u32,
                col_idx as u16,
                cell_value,
                &date_format,
                &datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_idx, col_idx, e))?;
        }
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

// ============================================================================
// DataFrame support
// ============================================================================

/// Write a Python value to the worksheet with optional column format
fn write_py_value_with_format(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'_, PyAny>,
    date_format: &Format,
    datetime_format: &Format,
    column_format: Option<&Format>,
) -> Result<(), String> {
    // Check for None first
    if value.is_none() {
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, "", fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, "")
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Check for pandas NA/NaT
    let type_name = value
        .get_type()
        .name()
        .map_err(|e| e.to_string())?
        .to_string();
    if type_name == "NAType" || type_name == "NaTType" {
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, "", fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, "")
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try boolean first (before int, since bool is subclass of int in Python)
    if let Ok(b) = value.cast::<PyBool>() {
        worksheet
            .write_boolean(row, col, b.is_true())
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Try datetime (before date, since datetime is subclass of date)
    if type_name == "datetime" || type_name == "Timestamp" {
        let year: i32 = value
            .getattr("year")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1900);
        let month: u32 = value
            .getattr("month")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let day: u32 = value
            .getattr("day")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let hour: u32 = value
            .getattr("hour")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);
        let minute: u32 = value
            .getattr("minute")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);
        let second: u32 = value
            .getattr("second")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            if let Some(time) = chrono::NaiveTime::from_hms_opt(hour, minute, second) {
                let dt = chrono::NaiveDateTime::new(date, time);
                let excel_dt = naive_datetime_to_excel(dt);
                // For datetime, use column format if provided, otherwise datetime_format
                let fmt = column_format.unwrap_or(datetime_format);
                worksheet
                    .write_number_with_format(row, col, excel_dt, fmt)
                    .map_err(|e| e.to_string())?;
                return Ok(());
            }
        }
    }

    // Try date
    if type_name == "date" {
        let year: i32 = value
            .getattr("year")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1900);
        let month: u32 = value
            .getattr("month")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let day: u32 = value
            .getattr("day")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            let excel_date = naive_date_to_excel(date);
            // For date, use column format if provided, otherwise date_format
            let fmt = column_format.unwrap_or(date_format);
            worksheet
                .write_number_with_format(row, col, excel_date, fmt)
                .map_err(|e| e.to_string())?;
            return Ok(());
        }
    }

    // Try integer
    if let Ok(i) = value.cast::<PyInt>() {
        if let Ok(val) = i.extract::<i64>() {
            if let Some(fmt) = column_format {
                worksheet
                    .write_number_with_format(row, col, val as f64, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_number(row, col, val as f64)
                    .map_err(|e| e.to_string())?;
            }
            return Ok(());
        }
    }

    // Try float
    if let Ok(f) = value.cast::<PyFloat>() {
        if let Ok(val) = f.extract::<f64>() {
            if val.is_nan() || val.is_infinite() {
                if let Some(fmt) = column_format {
                    worksheet
                        .write_string_with_format(row, col, "", fmt)
                        .map_err(|e| e.to_string())?;
                } else {
                    worksheet
                        .write_string(row, col, "")
                        .map_err(|e| e.to_string())?;
                }
            } else if let Some(fmt) = column_format {
                worksheet
                    .write_number_with_format(row, col, val, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_number(row, col, val)
                    .map_err(|e| e.to_string())?;
            }
            return Ok(());
        }
    }

    // Try to extract as f64 (covers numpy types)
    if let Ok(val) = value.extract::<f64>() {
        if val.is_nan() || val.is_infinite() {
            if let Some(fmt) = column_format {
                worksheet
                    .write_string_with_format(row, col, "", fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(row, col, "")
                    .map_err(|e| e.to_string())?;
            }
        } else if let Some(fmt) = column_format {
            worksheet
                .write_number_with_format(row, col, val, fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_number(row, col, val)
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try to extract as i64 (covers numpy int types)
    if let Ok(val) = value.extract::<i64>() {
        if let Some(fmt) = column_format {
            worksheet
                .write_number_with_format(row, col, val as f64, fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_number(row, col, val as f64)
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try to extract as bool
    if let Ok(val) = value.extract::<bool>() {
        worksheet
            .write_boolean(row, col, val)
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Try string
    if let Ok(s) = value.cast::<PyString>() {
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, s.to_string(), fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, s.to_string())
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Fallback: convert to string
    let s = value.str().map_err(|e| e.to_string())?.to_string();
    if let Some(fmt) = column_format {
        worksheet
            .write_string_with_format(row, col, &s, fmt)
            .map_err(|e| e.to_string())?;
    } else {
        worksheet
            .write_string(row, col, &s)
            .map_err(|e| e.to_string())?;
    }

    Ok(())
}

/// Convert a DataFrame (pandas or polars) to XLSX format
#[allow(clippy::too_many_arguments)]
fn convert_dataframe_to_xlsx(
    py: Python<'_>,
    df: &Bound<'_, PyAny>,
    output_path: &str,
    sheet_name: &str,
    include_header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    column_widths: Option<&HashMap<String, f64>>,
    table_name: Option<&str>,
    header_format: Option<&HashMap<String, Py<PyAny>>>,
    row_heights: Option<&HashMap<u32, f64>>,
    constant_memory: bool,
    column_formats: Option<&IndexMap<String, HashMap<String, Py<PyAny>>>>,
    conditional_formats: Option<&IndexMap<String, HashMap<String, Py<PyAny>>>>,
    formula_columns: Option<&IndexMap<String, String>>,
    merged_ranges: Option<&[MergedRange]>,
    hyperlinks: Option<&[Hyperlink]>,
) -> Result<(u32, u16), String> {
    // Create workbook and worksheet
    let mut workbook = Workbook::new();
    let worksheet = if constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Parse header format if provided
    let header_fmt = if let Some(fmt_dict) = header_format {
        Some(parse_header_format(py, fmt_dict)?)
    } else {
        None
    };

    let mut row_idx: u32 = 0;

    // Get column names - check polars first since it also has .columns
    let columns: Vec<String> =
        if df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false) {
            // polars DataFrame (has schema but no iloc)
            let cols = df.getattr("columns").map_err(|e| e.to_string())?;
            cols.extract().map_err(|e: pyo3::PyErr| e.to_string())?
        } else if df.hasattr("columns").unwrap_or(false) {
            // pandas DataFrame
            let cols = df.getattr("columns").map_err(|e| e.to_string())?;
            let col_list = cols.call_method0("tolist").map_err(|e| e.to_string())?;
            col_list.extract().map_err(|e: pyo3::PyErr| e.to_string())?
        } else {
            return Err("Unsupported DataFrame type".to_string());
        };

    let col_count = columns.len() as u16;

    // Build column formats if provided
    let col_formats: Vec<Option<Format>> = if let Some(cf) = column_formats {
        build_column_formats(py, &columns, cf)?
    } else {
        vec![None; columns.len()]
    };

    // Write header if requested (and not using table, since table handles headers)
    if include_header && table_style.is_none() {
        for (col_idx, col_name) in columns.iter().enumerate() {
            if let Some(ref fmt) = header_fmt {
                worksheet
                    .write_string_with_format(row_idx, col_idx as u16, col_name, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(row_idx, col_idx as u16, col_name)
                    .map_err(|e| e.to_string())?;
            }
        }
        row_idx += 1;
    }

    // If using table with header, write header in row 0
    let data_start_row = if table_style.is_some() && include_header {
        for (col_idx, col_name) in columns.iter().enumerate() {
            if let Some(ref fmt) = header_fmt {
                worksheet
                    .write_string_with_format(0, col_idx as u16, col_name, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(0, col_idx as u16, col_name)
                    .map_err(|e| e.to_string())?;
            }
        }
        row_idx = 1;
        0u32
    } else {
        row_idx.saturating_sub(1)
    };

    // Get row count
    let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
        let shape = df
            .getattr("shape")
            .map_err(|e: pyo3::PyErr| e.to_string())?;
        let shape_tuple: (usize, usize) =
            shape.extract().map_err(|e: pyo3::PyErr| e.to_string())?;
        shape_tuple.0
    } else {
        df.call_method0("__len__")
            .map_err(|e: pyo3::PyErr| e.to_string())?
            .extract()
            .map_err(|e: pyo3::PyErr| e.to_string())?
    };

    // Check if it's a polars DataFrame
    let is_polars = df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false);

    if is_polars {
        // Polars: iterate using rows()
        let rows = df.call_method0("iter_rows").map_err(|e| e.to_string())?;
        let iter = rows.try_iter().map_err(|e| e.to_string())?;
        for row_result in iter {
            let row = row_result.map_err(|e| e.to_string())?;
            let row_iter = row.try_iter().map_err(|e| e.to_string())?;
            let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                .collect::<Result<Vec<_>, _>>()
                .map_err(|e: PyErr| e.to_string())?;

            for (col_idx, value) in row_tuple.iter().enumerate() {
                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col_idx as u16,
                    value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx += 1;
        }
    } else {
        // Pandas: use .values for faster access
        let values = df.getattr("values").map_err(|e| e.to_string())?;

        for i in 0..row_count {
            let row = values
                .get_item(i)
                .map_err(|e| format!("Failed to get row {}: {}", i, e))?;

            for col_idx in 0..columns.len() {
                let value = row
                    .get_item(col_idx)
                    .map_err(|e| format!("Failed to get value at ({}, {}): {}", i, col_idx, e))?;

                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col_idx as u16,
                    &value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx += 1;
        }
    }

    // Add Excel Table if requested (not supported in constant_memory mode)
    // Tables require at least one data row, so skip if DataFrame is empty
    if let Some(style_name) = table_style {
        if !constant_memory && row_count > 0 {
            let style = parse_table_style(style_name)?;
            let mut table = Table::new().set_style(style);

            // Apply table name if provided
            if let Some(name) = table_name {
                let sanitized = sanitize_table_name(name);
                table = table.set_name(&sanitized);
            }

            let last_row = row_idx.saturating_sub(1);
            let last_col = col_count.saturating_sub(1);

            if last_row >= data_start_row {
                worksheet
                    .add_table(data_start_row, 0, last_row, last_col, &table)
                    .map_err(|e| format!("Failed to add table: {}", e))?;
            }
        }
    }

    // Apply formula columns (append calculated columns after data)
    // Formula columns are added after the original data columns
    let mut total_col_count = col_count;
    if let Some(formulas) = formula_columns {
        if !formulas.is_empty() && row_count > 0 {
            let data_row_start = if include_header { 1u32 } else { 0u32 };
            let data_row_end = row_idx.saturating_sub(1);
            if data_row_end >= data_row_start {
                let formula_cols_added = apply_formula_columns(
                    worksheet,
                    formulas,
                    col_count, // Start after original data columns
                    data_row_start,
                    data_row_end,
                    header_fmt.as_ref(),
                )?;
                total_col_count += formula_cols_added;
            }
        }
    }

    // Apply conditional formats (not supported in constant_memory mode)
    if let Some(cond_fmts) = conditional_formats {
        if !constant_memory && row_count > 0 {
            let data_row_start = if include_header { 1 } else { 0 };
            let data_row_end = row_idx.saturating_sub(1);
            if data_row_end >= data_row_start {
                apply_conditional_formats(
                    py,
                    worksheet,
                    &columns,
                    data_row_start,
                    data_row_end,
                    cond_fmts,
                )?;
            }
        }
    }

    // Freeze panes (freeze header row) - not supported in constant_memory mode
    if freeze_panes && include_header && !constant_memory {
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|e| format!("Failed to freeze panes: {}", e))?;
    }

    // Apply custom column widths and/or autofit
    if let Some(widths) = column_widths {
        if autofit && widths.contains_key("_all") && !constant_memory {
            // Autofit first, then apply cap from '_all' and specific widths
            apply_column_widths_with_autofit_cap(worksheet, col_count, widths, constant_memory)?;
        } else {
            // Just apply the specified widths
            apply_column_widths(worksheet, col_count, widths)?;
        }
    } else if autofit && !constant_memory {
        // Just autofit, no width constraints
        worksheet.autofit();
    }

    // Apply custom row heights if specified (not supported in constant_memory mode)
    if let Some(heights) = row_heights {
        if !constant_memory {
            for (&row_idx_h, &height) in heights.iter() {
                worksheet
                    .set_row_height(row_idx_h, height)
                    .map_err(|e| format!("Failed to set row height: {}", e))?;
            }
        }
    }

    // Apply merged ranges (not supported in constant_memory mode)
    if let Some(ranges) = merged_ranges {
        if !constant_memory && !ranges.is_empty() {
            apply_merged_ranges(py, worksheet, ranges)?;
        }
    }

    // Apply hyperlinks (not supported in constant_memory mode)
    if let Some(links) = hyperlinks {
        if !constant_memory && !links.is_empty() {
            apply_hyperlinks(worksheet, links)?;
        }
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_idx, total_col_count))
}

// ============================================================================
// Python bindings
// ============================================================================

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// This function reads a CSV file and writes it to an Excel XLSX file,
/// automatically detecting and converting data types:
/// - Numbers (integers and floats) become Excel numbers
/// - "true"/"false" become Excel booleans
/// - Dates (YYYY-MM-DD, etc.) become Excel dates with formatting
/// - Datetimes (ISO 8601) become Excel datetimes
/// - NaN/Inf values become empty cells
/// - Everything else becomes text
///
/// Args:
///     input_path: Path to the input CSV file
///     output_path: Path for the output XLSX file
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     parallel: Use multi-core parallel processing (default: False).
///               Faster for large files (100K+ rows) but uses more memory.
///     date_order: Date parsing order for ambiguous dates like "01-02-2024" (default: "auto").
///                 "auto" - ISO first, then European (DMY), then US (MDY)
///                 "mdy" or "us" - US format: 01-02-2024 = January 2nd
///                 "dmy" or "eu" - European format: 01-02-2024 = February 1st
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "output.xlsx")
///     >>> # For US date format (MM-DD-YYYY):
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "out.xlsx", date_order="us")
///     >>> # For large files, use parallel processing:
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("big.csv", "out.xlsx", parallel=True)
#[pyfunction]
#[pyo3(signature = (input_path, output_path, sheet_name = "Sheet1", parallel = false, date_order = "auto"))]
fn csv_to_xlsx(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    parallel: bool,
    date_order: &str,
) -> PyResult<(u32, u16)> {
    let order = DateOrder::parse(date_order).ok_or_else(|| {
        pyo3::exceptions::PyValueError::new_err(format!(
            "Invalid date_order '{}'. Valid values: auto, mdy, us, dmy, eu",
            date_order
        ))
    })?;

    let result = if parallel {
        convert_csv_to_xlsx_parallel(input_path, output_path, sheet_name, order)
    } else {
        convert_csv_to_xlsx(input_path, output_path, sheet_name, order)
    };
    result.map_err(pyo3::exceptions::PyValueError::new_err)
}

/// Convert a pandas or polars DataFrame to XLSX format.
///
/// This function writes a DataFrame directly to an Excel XLSX file,
/// preserving data types without intermediate CSV conversion.
///
/// Args:
///     df: pandas DataFrame or polars DataFrame to export
///     output_path: Path for the output XLSX file
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index (0-based) to width in characters (default: None)
///                    Example: {0: 20, 1: 15, 3: 30} sets widths for columns A, B, and D
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///                  Example: {0: 20, 5: 30} sets heights for specific rows
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Reduces memory usage but disables table_style, freeze_panes,
///                      row_heights, and autofit features.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [30, 25]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
///     >>> # With table formatting and auto-width columns:
///     >>> xlsxturbo.df_to_xlsx(df, "styled.xlsx", table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With custom column widths and row heights:
///     >>> xlsxturbo.df_to_xlsx(df, "custom.xlsx", column_widths={0: 25, 1: 10}, row_heights={0: 20})
///     >>> # For very large files, use constant_memory mode:
///     >>> xlsxturbo.df_to_xlsx(large_df, "big.xlsx", constant_memory=True)
///     >>> # With conditional formatting (color scales, data bars, icons):
///     >>> xlsxturbo.df_to_xlsx(df, "heatmap.xlsx", conditional_formats={
///     ...     'score': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'},
///     ...     'progress': {'type': 'data_bar', 'bar_color': '#638EC6'},
///     ...     'status': {'type': 'icon_set', 'icon_type': '3_traffic_lights'}
///     ... })
#[pyfunction]
#[pyo3(signature = (df, output_path, sheet_name = "Sheet1", header = true, autofit = false, table_style = None, freeze_panes = false, column_widths = None, table_name = None, header_format = None, row_heights = None, constant_memory = false, column_formats = None, conditional_formats = None, formula_columns = None, merged_ranges = None, hyperlinks = None))]
#[allow(clippy::too_many_arguments)]
fn df_to_xlsx<'py>(
    py: Python<'py>,
    df: &Bound<'py, PyAny>,
    output_path: &str,
    sheet_name: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<HashMap<u32, f64>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
) -> PyResult<(u32, u16)> {
    // Extract column_widths if provided
    let extracted_column_widths = if let Some(cw) = column_widths {
        if let Ok(dict) = cw.cast::<pyo3::types::PyDict>() {
            Some(extract_column_widths(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract header_format if provided
    let extracted_header_format = if let Some(hf) = header_format {
        if let Ok(dict) = hf.cast::<pyo3::types::PyDict>() {
            Some(extract_header_format(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract column_formats if provided (uses IndexMap to preserve order)
    let extracted_column_formats = if let Some(cf) = column_formats {
        if let Ok(dict) = cf.cast::<pyo3::types::PyDict>() {
            Some(extract_column_formats(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract conditional_formats if provided
    let extracted_conditional_formats = if let Some(cf) = conditional_formats {
        if let Ok(dict) = cf.cast::<pyo3::types::PyDict>() {
            Some(extract_conditional_formats(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract formula_columns if provided (column name -> formula template)
    let extracted_formula_columns = if let Some(fc) = formula_columns {
        if let Ok(dict) = fc.cast::<pyo3::types::PyDict>() {
            Some(extract_formula_columns(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract merged_ranges if provided (list of tuples)
    let extracted_merged_ranges = if let Some(mr) = merged_ranges {
        if let Ok(list) = mr.cast::<pyo3::types::PyList>() {
            Some(extract_merged_ranges(list)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract hyperlinks if provided (list of tuples)
    let extracted_hyperlinks = if let Some(hl) = hyperlinks {
        if let Ok(list) = hl.cast::<pyo3::types::PyList>() {
            Some(extract_hyperlinks(list)?)
        } else {
            None
        }
    } else {
        None
    };

    convert_dataframe_to_xlsx(
        py,
        df,
        output_path,
        sheet_name,
        header,
        autofit,
        table_style,
        freeze_panes,
        extracted_column_widths.as_ref(),
        table_name.as_deref(),
        extracted_header_format.as_ref(),
        row_heights.as_ref(),
        constant_memory,
        extracted_column_formats.as_ref(),
        extracted_conditional_formats.as_ref(),
        extracted_formula_columns.as_ref(),
        extracted_merged_ranges.as_deref(),
        extracted_hyperlinks.as_deref(),
    )
    .map_err(pyo3::exceptions::PyValueError::new_err)
}

/// Get the version of the xlsxturbo library
#[pyfunction]
fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

/// Write multiple DataFrames to separate sheets in a single workbook.
///
/// This is a convenience function that writes multiple DataFrames to
/// separate sheets in one workbook, which is more efficient than
/// calling df_to_xlsx multiple times.
///
/// Args:
///     sheets: List of tuples. Each tuple can be:
///             - (DataFrame, sheet_name) - uses global defaults
///             - (DataFrame, sheet_name, options_dict) - per-sheet overrides
///             Options dict keys: header, autofit, table_style, freeze_panes,
///             column_widths, row_heights, table_name, header_format, column_formats,
///             conditional_formats
///     output_path: Path for the output XLSX file
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index or "_all" to width in characters (default: None)
///                    Example: {0: 20, "_all": 50} sets col A to 20, caps others at 50
///     table_name: Name for Excel table (default: auto-generated)
///     header_format: Dict with header cell formatting options (default: None)
///                    Example: {"bold": True, "bg_color": "#4F81BD", "font_color": "white"}
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///     constant_memory: Use constant memory mode for large files (default: False).
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///
/// Returns:
///     List of (rows, columns) tuples for each sheet
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df1 = pd.DataFrame({'a': [1, 2]})
///     >>> df2 = pd.DataFrame({'b': [3, 4]})
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sheet1"), (df2, "Sheet2")], "out.xlsx")
///     >>> # With styling applied to all sheets:
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sales"), (df2, "Regions")], "report.xlsx",
///     ...                       table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With per-sheet options (header=False for one sheet):
///     >>> xlsxturbo.dfs_to_xlsx([
///     ...     (df1, "Data", {"header": True, "table_style": "Medium2"}),
///     ...     (df2, "Instructions", {"header": False})
///     ... ], "report.xlsx", autofit=True)
#[pyfunction]
#[pyo3(signature = (sheets, output_path, header = true, autofit = false, table_style = None, freeze_panes = false, column_widths = None, table_name = None, header_format = None, row_heights = None, constant_memory = false, column_formats = None, conditional_formats = None, formula_columns = None, merged_ranges = None, hyperlinks = None))]
#[allow(clippy::too_many_arguments)]
fn dfs_to_xlsx<'py>(
    py: Python<'py>,
    sheets: Vec<Bound<'py, PyAny>>,
    output_path: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<HashMap<u32, f64>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
) -> PyResult<Vec<(u32, u16)>> {
    let mut workbook = Workbook::new();
    let mut stats = Vec::new();

    // Extract global column_widths if provided
    let extracted_column_widths = if let Some(cw) = column_widths {
        if let Ok(dict) = cw.cast::<pyo3::types::PyDict>() {
            Some(extract_column_widths(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global header_format if provided
    let extracted_header_format = if let Some(hf) = header_format {
        if let Ok(dict) = hf.cast::<pyo3::types::PyDict>() {
            Some(extract_header_format(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global column_formats if provided (uses IndexMap to preserve order)
    let extracted_column_formats = if let Some(cf) = column_formats {
        if let Ok(dict) = cf.cast::<pyo3::types::PyDict>() {
            Some(extract_column_formats(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global conditional_formats if provided
    let extracted_conditional_formats = if let Some(cf) = conditional_formats {
        if let Ok(dict) = cf.cast::<pyo3::types::PyDict>() {
            Some(extract_conditional_formats(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global formula_columns if provided
    let extracted_formula_columns = if let Some(fc) = formula_columns {
        if let Ok(dict) = fc.cast::<pyo3::types::PyDict>() {
            Some(extract_formula_columns(dict)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global merged_ranges if provided
    let extracted_merged_ranges = if let Some(mr) = merged_ranges {
        if let Ok(list) = mr.cast::<pyo3::types::PyList>() {
            Some(extract_merged_ranges(list)?)
        } else {
            None
        }
    } else {
        None
    };

    // Extract global hyperlinks if provided
    let extracted_hyperlinks = if let Some(hl) = hyperlinks {
        if let Ok(list) = hl.cast::<pyo3::types::PyList>() {
            Some(extract_hyperlinks(list)?)
        } else {
            None
        }
    } else {
        None
    };

    // Create formats
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Parse global header format if provided
    let global_header_fmt = if let Some(ref fmt_dict) = extracted_header_format {
        Some(parse_header_format(py, fmt_dict).map_err(pyo3::exceptions::PyValueError::new_err)?)
    } else {
        None
    };

    for sheet_tuple in sheets {
        // Extract sheet info (supports both 2-tuple and 3-tuple formats)
        let (df, sheet_name, sheet_config) = extract_sheet_info(&sheet_tuple)?;

        // Merge per-sheet options with global defaults
        let effective_header = sheet_config.header.unwrap_or(header);
        let effective_autofit = sheet_config.autofit.unwrap_or(autofit);
        let effective_table_style: Option<String> = match sheet_config.table_style {
            Some(style_opt) => style_opt,
            None => table_style.map(|s| s.to_string()),
        };
        let effective_freeze_panes = sheet_config.freeze_panes.unwrap_or(freeze_panes);
        let effective_column_widths = sheet_config
            .column_widths
            .as_ref()
            .or(extracted_column_widths.as_ref());
        let effective_row_heights = sheet_config.row_heights.as_ref().or(row_heights.as_ref());
        let effective_table_name = sheet_config.table_name.as_ref().or(table_name.as_ref());

        // Parse per-sheet header format or use global
        let effective_header_fmt = if let Some(ref fmt_dict) = sheet_config.header_format {
            Some(
                parse_header_format(py, fmt_dict)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?,
            )
        } else {
            global_header_fmt.clone()
        };

        // Get effective column formats (per-sheet or global)
        let effective_column_formats = sheet_config
            .column_formats
            .as_ref()
            .or(extracted_column_formats.as_ref());

        let worksheet = if constant_memory {
            workbook.add_worksheet_with_constant_memory()
        } else {
            workbook.add_worksheet()
        };
        worksheet.set_name(&sheet_name).map_err(|e| {
            pyo3::exceptions::PyValueError::new_err(format!(
                "Failed to set sheet name '{}': {}",
                sheet_name, e
            ))
        })?;

        let mut row_idx: u32 = 0;

        // Get column names - check polars first
        let columns: Vec<String> = if df.hasattr("schema").unwrap_or(false)
            && !df.hasattr("iloc").unwrap_or(false)
        {
            let cols = df
                .getattr("columns")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            cols.extract()
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
        } else if df.hasattr("columns").unwrap_or(false) {
            let cols = df
                .getattr("columns")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            let col_list = cols
                .call_method0("tolist")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            col_list
                .extract()
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
        } else {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "Unsupported DataFrame type",
            ));
        };

        let col_count = columns.len() as u16;

        // Build column formats if provided
        let col_formats: Vec<Option<Format>> = if let Some(cf) = effective_column_formats {
            build_column_formats(py, &columns, cf)
                .map_err(pyo3::exceptions::PyValueError::new_err)?
        } else {
            vec![None; columns.len()]
        };

        // Write header if requested
        if effective_header {
            for (col_idx, col_name) in columns.iter().enumerate() {
                if let Some(ref fmt) = effective_header_fmt {
                    worksheet
                        .write_string_with_format(row_idx, col_idx as u16, col_name, fmt)
                        .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                } else {
                    worksheet
                        .write_string(row_idx, col_idx as u16, col_name)
                        .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                }
            }
            row_idx += 1;
        }

        // Get row count and check if polars
        let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
            let shape = df
                .getattr("shape")
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            let shape_tuple: (usize, usize) = shape
                .extract()
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            shape_tuple.0
        } else {
            df.call_method0("__len__")
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
                .extract()
                .map_err(|e: pyo3::PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
        };

        let is_polars =
            df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false);

        // Write data rows
        if is_polars {
            let rows = df
                .call_method0("iter_rows")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            let iter = rows
                .try_iter()
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            for row_result in iter {
                let row = row_result
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                let row_iter = row
                    .try_iter()
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                    .collect::<Result<Vec<_>, _>>()
                    .map_err(|e: PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;

                for (col_idx, value) in row_tuple.iter().enumerate() {
                    write_py_value_with_format(
                        worksheet,
                        row_idx,
                        col_idx as u16,
                        value,
                        &date_format,
                        &datetime_format,
                        col_formats.get(col_idx).and_then(|f| f.as_ref()),
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                }
                row_idx += 1;
            }
        } else {
            let values = df
                .getattr("values")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            for i in 0..row_count {
                let row = values.get_item(i).map_err(|e| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Failed to get row {}: {}",
                        i, e
                    ))
                })?;

                for col_idx in 0..columns.len() {
                    let value = row.get_item(col_idx).map_err(|e| {
                        pyo3::exceptions::PyValueError::new_err(format!(
                            "Failed to get value at ({}, {}): {}",
                            i, col_idx, e
                        ))
                    })?;

                    write_py_value_with_format(
                        worksheet,
                        row_idx,
                        col_idx as u16,
                        &value,
                        &date_format,
                        &datetime_format,
                        col_formats.get(col_idx).and_then(|f| f.as_ref()),
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                }
                row_idx += 1;
            }
        }

        // Add Excel Table if requested (not supported in constant_memory mode)
        // Tables require at least one data row, so skip if DataFrame is empty
        if let Some(ref style_name) = effective_table_style {
            if !constant_memory && row_count > 0 {
                let style = parse_table_style(style_name)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                let mut table = Table::new().set_style(style);

                // Apply table name if provided
                if let Some(name) = effective_table_name {
                    let sanitized = sanitize_table_name(name);
                    table = table.set_name(&sanitized);
                }

                let data_start_row = 0u32;
                let last_row = row_idx.saturating_sub(1);
                let last_col = col_count.saturating_sub(1);

                if last_row >= data_start_row {
                    worksheet
                        .add_table(data_start_row, 0, last_row, last_col, &table)
                        .map_err(|e| {
                            pyo3::exceptions::PyValueError::new_err(format!(
                                "Failed to add table: {}",
                                e
                            ))
                        })?;
                }
            }
        }

        // Apply formula columns (append calculated columns after data)
        // Use per-sheet formula_columns or fall back to global
        let effective_formula_columns = sheet_config
            .formula_columns
            .as_ref()
            .or(extracted_formula_columns.as_ref());
        let mut total_col_count = col_count;
        if let Some(formulas) = effective_formula_columns {
            if !formulas.is_empty() && row_count > 0 {
                let data_row_start = if effective_header { 1u32 } else { 0u32 };
                let data_row_end = row_idx.saturating_sub(1);
                if data_row_end >= data_row_start {
                    let formula_cols_added = apply_formula_columns(
                        worksheet,
                        formulas,
                        col_count, // Start after original data columns
                        data_row_start,
                        data_row_end,
                        effective_header_fmt.as_ref(),
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                    total_col_count += formula_cols_added;
                }
            }
        }

        // Apply conditional formats (not supported in constant_memory mode)
        // Use per-sheet conditional_formats or fall back to global
        let effective_conditional_formats = sheet_config
            .conditional_formats
            .as_ref()
            .or(extracted_conditional_formats.as_ref());
        if let Some(cond_fmts) = effective_conditional_formats {
            if !constant_memory && row_count > 0 {
                let data_row_start = if effective_header { 1 } else { 0 };
                let data_row_end = row_idx.saturating_sub(1);
                if data_row_end >= data_row_start {
                    apply_conditional_formats(
                        py,
                        worksheet,
                        &columns,
                        data_row_start,
                        data_row_end,
                        cond_fmts,
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                }
            }
        }

        // Freeze panes (freeze header row) - not supported in constant_memory mode
        if effective_freeze_panes && effective_header && !constant_memory {
            worksheet.set_freeze_panes(1, 0).map_err(|e| {
                pyo3::exceptions::PyValueError::new_err(format!("Failed to freeze panes: {}", e))
            })?;
        }

        // Apply custom column widths and/or autofit
        if let Some(widths) = effective_column_widths {
            if effective_autofit && widths.contains_key("_all") && !constant_memory {
                // Autofit first, then apply cap from '_all' and specific widths
                apply_column_widths_with_autofit_cap(worksheet, col_count, widths, constant_memory)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
            } else {
                // Just apply the specified widths
                apply_column_widths(worksheet, col_count, widths)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
            }
        } else if effective_autofit && !constant_memory {
            // Just autofit, no width constraints
            worksheet.autofit();
        }

        // Apply custom row heights if specified (not supported in constant_memory mode)
        if let Some(heights) = effective_row_heights {
            if !constant_memory {
                for (&row_idx_h, &height) in heights.iter() {
                    worksheet.set_row_height(row_idx_h, height).map_err(|e| {
                        pyo3::exceptions::PyValueError::new_err(format!(
                            "Failed to set row height: {}",
                            e
                        ))
                    })?;
                }
            }
        }

        // Apply merged ranges (not supported in constant_memory mode)
        // Use per-sheet merged_ranges or fall back to global
        let effective_merged_ranges = sheet_config
            .merged_ranges
            .as_ref()
            .or(extracted_merged_ranges.as_ref());
        if let Some(ranges) = effective_merged_ranges {
            if !constant_memory && !ranges.is_empty() {
                apply_merged_ranges(py, worksheet, ranges)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
            }
        }

        // Apply hyperlinks (not supported in constant_memory mode)
        // Use per-sheet hyperlinks or fall back to global
        let effective_hyperlinks = sheet_config
            .hyperlinks
            .as_ref()
            .or(extracted_hyperlinks.as_ref());
        if let Some(links) = effective_hyperlinks {
            if !constant_memory && !links.is_empty() {
                apply_hyperlinks(worksheet, links)
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
            }
        }

        stats.push((row_idx, total_col_count));
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| pyo3::exceptions::PyValueError::new_err(format!("Failed to save: {}", e)))?;

    Ok(stats)
}

/// xlsxturbo - High-performance Excel writer
///
/// A Rust-powered library for converting DataFrames and CSV files to Excel XLSX format.
/// Up to 25x faster than pure Python solutions.
///
/// Features:
/// - Direct DataFrame support (pandas and polars)
/// - Automatic type detection (numbers, booleans, dates, datetimes)
/// - Proper Excel formatting for dates and times
/// - Handles NaN/Inf/None gracefully
/// - Memory-efficient for large files
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'a': [1, 2], 'b': [3.14, 2.71]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
#[pymodule]
fn xlsxturbo(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(csv_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(df_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(dfs_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_integer() {
        assert!(matches!(
            parse_value("123", DateOrder::Auto),
            CellValue::Integer(123)
        ));
        assert!(matches!(
            parse_value("-456", DateOrder::Auto),
            CellValue::Integer(-456)
        ));
    }

    #[test]
    fn test_parse_float() {
        if let CellValue::Float(v) = parse_value("3.14", DateOrder::Auto) {
            assert!((v - 3.14).abs() < 0.001);
        } else {
            panic!("Expected float");
        }
    }

    #[test]
    fn test_parse_boolean() {
        assert!(matches!(
            parse_value("true", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("TRUE", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("false", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
        assert!(matches!(
            parse_value("False", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_parse_empty() {
        assert!(matches!(parse_value("", DateOrder::Auto), CellValue::Empty));
        assert!(matches!(
            parse_value("   ", DateOrder::Auto),
            CellValue::Empty
        ));
        assert!(matches!(
            parse_value("NaN", DateOrder::Auto),
            CellValue::Empty
        ));
    }

    #[test]
    fn test_parse_date() {
        assert!(matches!(
            parse_value("2024-01-15", DateOrder::Auto),
            CellValue::Date(_)
        ));
        assert!(matches!(
            parse_value("2024/01/15", DateOrder::Auto),
            CellValue::Date(_)
        ));
    }

    #[test]
    fn test_parse_datetime() {
        assert!(matches!(
            parse_value("2024-01-15T10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
        assert!(matches!(
            parse_value("2024-01-15 10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
    }

    #[test]
    fn test_parse_string() {
        assert!(matches!(
            parse_value("hello", DateOrder::Auto),
            CellValue::String(_)
        ));
    }

    #[test]
    fn test_matches_pattern_exact() {
        assert!(matches_pattern("column_name", "column_name"));
        assert!(!matches_pattern("column_name", "other"));
    }

    #[test]
    fn test_matches_pattern_prefix() {
        assert!(matches_pattern("price_usd", "price_*"));
        assert!(matches_pattern("price_", "price_*"));
        assert!(!matches_pattern("cost_usd", "price_*"));
    }

    #[test]
    fn test_matches_pattern_suffix() {
        assert!(matches_pattern("col_weight", "*_weight"));
        assert!(matches_pattern("_weight", "*_weight"));
        assert!(!matches_pattern("col_height", "*_weight"));
    }

    #[test]
    fn test_matches_pattern_contains() {
        assert!(matches_pattern("leadframe_difference", "*difference*"));
        assert!(matches_pattern("difference", "*difference*"));
        assert!(matches_pattern("my_difference_col", "*difference*"));
        assert!(!matches_pattern("other_column", "*difference*"));
    }
}
