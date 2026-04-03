//! Python extraction functions for converting Python objects to Rust types

use crate::parse::{parse_cell_ref, parse_horizontal_alignment, parse_vertical_alignment};
use crate::types::{
    CellWrite, Comment, ConditionalFormatConfigs, Hyperlink, ImageConfig, MergedRange,
    RichTextSegment, SheetConfig, ValidationConfig,
};
use indexmap::IndexMap;
use pyo3::prelude::*;
use std::collections::HashMap;

/// Convert a Python dict to a Rust HashMap<String, Py<PyAny>>
pub(crate) fn pydict_to_hashmap(
    dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    let mut map = HashMap::new();
    for (k, v) in dict.iter() {
        map.insert(k.extract()?, v.unbind());
    }
    Ok(map)
}

/// Helper: extract an optional scalar field from a Python dict into a SheetConfig field
macro_rules! extract_scalar {
    ($opts:expr, $config:expr, $key:literal, $field:ident) => {
        if let Ok(val) = $opts.get_item($key) {
            if !val.is_none() {
                $config.$field = Some(val.extract()?);
            }
        }
    };
}

/// Helper: extract an optional dict field, run an extractor, and set if non-empty
macro_rules! extract_dict_field {
    ($opts:expr, $config:expr, $key:literal, $field:ident, $extractor:expr) => {
        if let Ok(val) = $opts.get_item($key) {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = $extractor(dict)?;
                    if !extracted.is_empty() {
                        $config.$field = Some(extracted);
                    }
                }
            }
        }
    };
}

/// Helper: extract an optional list field, run an extractor, and set if non-empty
macro_rules! extract_list_field {
    ($opts:expr, $config:expr, $key:literal, $field:ident, $extractor:expr) => {
        if let Ok(val) = $opts.get_item($key) {
            if !val.is_none() {
                if let Ok(list) = val.cast::<pyo3::types::PyList>() {
                    let extracted = $extractor(list)?;
                    if !extracted.is_empty() {
                        $config.$field = Some(extracted);
                    }
                }
            }
        }
    };
}

/// Extract sheet info from a Python tuple (supports both 2-tuple and 3-tuple formats)
/// 2-tuple: (df, sheet_name)
/// 3-tuple: (df, sheet_name, options_dict)
pub(crate) fn extract_sheet_info<'py>(
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

        // Extract scalar fields
        extract_scalar!(opts, config, "header", header);
        extract_scalar!(opts, config, "autofit", autofit);
        extract_scalar!(opts, config, "freeze_panes", freeze_panes);
        extract_scalar!(opts, config, "row_heights", row_heights);
        extract_scalar!(opts, config, "table_name", table_name);

        // table_style needs special handling: None means "explicitly no style"
        if let Ok(val) = opts.get_item("table_style") {
            if val.is_none() {
                config.table_style = Some(None);
            } else {
                config.table_style = Some(Some(val.extract()?));
            }
        }

        // Extract complex dict fields
        extract_dict_field!(
            opts,
            config,
            "column_widths",
            column_widths,
            extract_column_widths
        );
        extract_dict_field!(
            opts,
            config,
            "header_format",
            header_format,
            extract_header_format
        );
        extract_dict_field!(
            opts,
            config,
            "column_formats",
            column_formats,
            extract_column_formats
        );
        extract_dict_field!(
            opts,
            config,
            "conditional_formats",
            conditional_formats,
            extract_conditional_formats
        );
        extract_dict_field!(
            opts,
            config,
            "formula_columns",
            formula_columns,
            extract_formula_columns
        );
        extract_dict_field!(opts, config, "comments", comments, extract_comments);
        extract_dict_field!(
            opts,
            config,
            "validations",
            validations,
            extract_validations
        );
        extract_dict_field!(opts, config, "rich_text", rich_text, extract_rich_text);
        extract_dict_field!(opts, config, "images", images, extract_images);

        // Extract cells
        if let Ok(val) = opts.get_item("cells") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = extract_cells(dict)?;
                    if !extracted.is_empty() {
                        config.cells = Some(extracted);
                    }
                }
            }
        }

        // Extract complex list fields
        extract_list_field!(
            opts,
            config,
            "merged_ranges",
            merged_ranges,
            extract_merged_ranges
        );
        extract_list_field!(opts, config, "hyperlinks", hyperlinks, extract_hyperlinks);

        config
    } else {
        SheetConfig::default()
    };

    Ok((df, sheet_name, config))
}

/// Extract column_widths from Python dict, supporting both integer and string keys
pub(crate) fn extract_column_widths(
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
pub(crate) fn extract_header_format(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    pydict_to_hashmap(py_dict)
}

/// Extract column_formats from Python dict (pattern -> format dict)
/// Uses IndexMap to preserve insertion order from Python dict
pub(crate) fn extract_column_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (pattern, fmt_dict) in py_dict.iter() {
        let pattern_str: String = pattern.extract()?;
        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
            col_fmts.insert(pattern_str, pydict_to_hashmap(inner_dict)?);
        }
    }
    Ok(col_fmts)
}

/// Extract conditional_formats from Python dict (column/pattern -> config dict or list of dicts)
/// Uses IndexMap to preserve insertion order for pattern matching (first match wins)
pub(crate) fn extract_conditional_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<ConditionalFormatConfigs> {
    let mut cond_fmts: ConditionalFormatConfigs = IndexMap::new();
    for (col_name, fmt_value) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        // Accept either a single dict or a list of dicts
        if let Ok(list) = fmt_value.cast::<pyo3::types::PyList>() {
            let mut configs = Vec::new();
            for (i, item) in list.iter().enumerate() {
                let d = item.cast::<pyo3::types::PyDict>().map_err(|_| {
                    pyo3::exceptions::PyTypeError::new_err(format!(
                        "conditional_formats['{}']: list item {} must be a dict",
                        col_str, i
                    ))
                })?;
                configs.push(pydict_to_hashmap(d)?);
            }
            cond_fmts.insert(col_str, configs);
        } else if let Ok(inner_dict) = fmt_value.cast::<pyo3::types::PyDict>() {
            cond_fmts.insert(col_str, vec![pydict_to_hashmap(inner_dict)?]);
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                "conditional_formats['{}']: value must be a dict or list of dicts",
                col_str
            )));
        }
    }
    Ok(cond_fmts)
}

/// Extract formula_columns from Python dict (column name -> formula template)
/// Uses IndexMap to preserve column order
pub(crate) fn extract_formula_columns(
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

/// Extract merged_ranges from Python list of tuples
/// Each tuple: (range_str, text) or (range_str, text, format_dict)
pub(crate) fn extract_merged_ranges(
    py_list: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<MergedRange>> {
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
                    Some(pydict_to_hashmap(dict)?)
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

/// Extract hyperlinks from Python list of tuples
/// Each tuple: (cell_ref, url) or (cell_ref, url, display_text)
pub(crate) fn extract_hyperlinks(
    py_list: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<Hyperlink>> {
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

/// Extract comments from Python dict
/// Supports: {'A1': 'text'} or {'A1': {'text': 'note', 'author': 'John'}}
pub(crate) fn extract_comments(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Comment>> {
    let mut comments: HashMap<String, Comment> = HashMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;

        // Check if value is a dict or simple string
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            // Dict format: {'text': '...', 'author': '...'}
            let text: String = inner_dict
                .get_item("text")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Comment at '{}' missing 'text' key",
                        cell_str
                    ))
                })?
                .extract()?;
            let author: Option<String> = if let Ok(Some(a)) = inner_dict.get_item("author") {
                if !a.is_none() {
                    Some(a.extract()?)
                } else {
                    None
                }
            } else {
                None
            };
            comments.insert(cell_str, (text, author));
        } else {
            // Simple string format
            let text: String = value.extract()?;
            comments.insert(cell_str, (text, None));
        }
    }

    Ok(comments)
}

/// Extract validations from Python dict (column name/pattern -> validation config)
pub(crate) fn extract_validations(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, ValidationConfig>> {
    let mut validations: IndexMap<String, ValidationConfig> = IndexMap::new();
    for (col_name, config) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        if let Ok(inner_dict) = config.cast::<pyo3::types::PyDict>() {
            validations.insert(col_str, pydict_to_hashmap(inner_dict)?);
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                "validations['{}']: expected dict, got {}",
                col_str,
                config.get_type().name()?
            )));
        }
    }
    Ok(validations)
}

/// Extract rich_text from Python dict (cell_ref -> list of segments)
pub(crate) fn extract_rich_text(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Vec<RichTextSegment>>> {
    let mut rich_text: HashMap<String, Vec<RichTextSegment>> = HashMap::new();

    for (cell_ref, segments_list) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;
        let mut segments: Vec<RichTextSegment> = Vec::new();

        if let Ok(list) = segments_list.cast::<pyo3::types::PyList>() {
            for (idx, item) in list.iter().enumerate() {
                // Check if item is a tuple (text, format_dict) or just a string
                if let Ok(tuple) = item.cast::<pyo3::types::PyTuple>() {
                    let text: String = tuple.get_item(0)?.extract()?;
                    let format_dict = if tuple.len() >= 2 {
                        let fmt_item = tuple.get_item(1)?;
                        if let Ok(dict) = fmt_item.cast::<pyo3::types::PyDict>() {
                            Some(pydict_to_hashmap(dict)?)
                        } else {
                            None
                        }
                    } else {
                        None
                    };
                    segments.push((text, format_dict));
                } else if let Ok(text) = item.extract::<String>() {
                    // Plain string segment
                    segments.push((text, None));
                } else {
                    return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                        "rich_text['{}']: segment {} must be a string or tuple (text, format_dict), got {}",
                        cell_str,
                        idx,
                        item.get_type().name()?
                    )));
                }
            }
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                "rich_text['{}']: expected list of segments, got {}",
                cell_str,
                segments_list.get_type().name()?
            )));
        }

        if !segments.is_empty() {
            rich_text.insert(cell_str, segments);
        }
    }

    Ok(rich_text)
}

/// Extract images from Python dict (cell_ref -> path or config dict)
pub(crate) fn extract_images(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, ImageConfig>> {
    let mut images: HashMap<String, ImageConfig> = HashMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;

        // Check if value is a dict or simple string (path)
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            // Dict format: {'path': '...', 'scale_width': 0.5, ...}
            let path: String = inner_dict
                .get_item("path")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Image at '{}' missing 'path' key",
                        cell_str
                    ))
                })?
                .extract()?;
            let mut options = pydict_to_hashmap(inner_dict)?;
            options.remove("path");
            images.insert(cell_str, (path, Some(options)));
        } else {
            // Simple string format (just path)
            let path: String = value.extract()?;
            images.insert(cell_str, (path, None));
        }
    }

    Ok(images)
}

/// Extract cells from Python dict (cell_ref -> value or {value, num_format, align_horizontal, ...})
pub(crate) fn extract_cells(py_dict: &Bound<'_, pyo3::types::PyDict>) -> PyResult<Vec<CellWrite>> {
    let mut cells = Vec::new();
    for (key, value) in py_dict.iter() {
        let cell_ref: String = key.extract()?;
        let (row, col) =
            parse_cell_ref(&cell_ref).map_err(pyo3::exceptions::PyValueError::new_err)?;

        // Check if value is a dict with "value" and optional formatting keys
        if let Ok(d) = value.cast::<pyo3::types::PyDict>() {
            let val = d.get_item("value")?.ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "cells['{}'] dict missing 'value' key",
                    cell_ref
                ))
            })?;
            let num_fmt: Option<String> = d
                .get_item("num_format")?
                .map(|v| v.extract::<String>())
                .transpose()?;
            let align_h: Option<String> = d
                .get_item("align_horizontal")?
                .map(|v| v.extract::<String>())
                .transpose()?;
            if let Some(ref ah) = align_h {
                parse_horizontal_alignment(ah).map_err(pyo3::exceptions::PyValueError::new_err)?;
            }
            let align_v: Option<String> = d
                .get_item("align_vertical")?
                .map(|v| v.extract::<String>())
                .transpose()?;
            if let Some(ref av) = align_v {
                parse_vertical_alignment(av).map_err(pyo3::exceptions::PyValueError::new_err)?;
            }
            let wrap: bool = d
                .get_item("wrap_text")?
                .map(|v| v.extract::<bool>().unwrap_or(false))
                .unwrap_or(false);
            cells.push(CellWrite {
                row,
                col,
                value: val.unbind(),
                num_format: num_fmt,
                align_horizontal: align_h,
                align_vertical: align_v,
                wrap_text: wrap,
            });
        } else {
            cells.push(CellWrite {
                row,
                col,
                value: value.unbind(),
                num_format: None,
                align_horizontal: None,
                align_vertical: None,
                wrap_text: false,
            });
        }
    }
    Ok(cells)
}
