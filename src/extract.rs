//! Python extraction functions for converting Python objects to Rust types

use crate::parse::{parse_cell_ref, parse_horizontal_alignment, parse_vertical_alignment};
use crate::types::{
    pydict_to_hashmap, pytype_name, reject_unknown_keys as types_reject_unknown_keys, CellWrite,
    ChartConfig, CheckboxConfig, Comment, ConditionalFormatConfigs, Hyperlink, ImageConfig,
    MergedRange, RichTextSegment, SheetConfig, SparklineConfig, TextboxConfig, ValidationConfig,
};
use indexmap::IndexMap;
use pyo3::prelude::*;
use std::collections::HashMap;

const SHEET_OPTION_NAMES: &[&str] = &[
    "header",
    "autofit",
    "table_style",
    "freeze_panes",
    "column_widths",
    "row_heights",
    "table_name",
    "header_format",
    "column_formats",
    "conditional_formats",
    "formula_columns",
    "merged_ranges",
    "hyperlinks",
    "comments",
    "validations",
    "rich_text",
    "images",
    "checkboxes",
    "textboxes",
    "charts",
    "sparklines",
    "cells",
];

/// Helper: extract a nested option value, classifying a conversion failure as a
/// `ConfigurationTypeError` that names the option and what was expected.
///
/// Exists because a bare `extract()?` propagates PyO3's own `TypeError`, which is
/// *outside* the public hierarchy. Until 0.19.1 that made `docs/errors.md`'s
/// promise — "every failure xlsxturbo itself raises is an `XlsxTurboError`" —
/// false for every nested key and value: a non-string `column_formats` key, a
/// non-string `formula_columns` value, a bad `merged_ranges` tuple element and
/// several more all escaped unclassified.
///
/// The reachability tests in `tests/test_errors.py` could not catch that. They
/// exercise one trigger per exported class, which proves those five paths reach
/// the hierarchy and says nothing about the population of extraction paths — a
/// consistency check proves agreement, never completeness. The guard that does
/// catch it is `TestNestedExtractionStaysInTheHierarchy` in that same file,
/// which drives a negative-input matrix across every extractor family.
///
/// Takes `$value` by reference so a temporary (`item.get_item(0)?`) lives to the
/// end of the macro's block and the caller's binding is not moved.
macro_rules! extract_typed {
    ($value:expr, $type_desc:literal, $($context:tt)+) => {{
        let value = &$value;
        value.extract().map_err(|_| {
            crate::errors::configuration_type(format!(
                "{}: expected {}, got {}",
                format!($($context)+),
                $type_desc,
                pytype_name(value)
            ))
        })?
    }};
}

/// Helper: extract an optional scalar field from a Python dict into a SheetConfig field.
///
/// `$opts` is a `Bound<PyAny>`, so `.get_item($key)` goes through the mapping
/// protocol and raises a Python `KeyError` for a missing key rather than
/// returning `Ok(None)` — that specific error is the expected "option not
/// given" case and is swallowed here, same as before. Any *other* error
/// (e.g. a non-string-keyed options mapping) now propagates instead of being
/// silently discarded, and a wrong-typed value produces a context-rich
/// `TypeError` naming the option and the received type (via `pytype_name`)
/// instead of falling through to pyo3's generic conversion error.
macro_rules! extract_scalar {
    ($opts:expr, $config:expr, $key:literal, $field:ident, $type_desc:literal) => {
        match $opts.get_item($key) {
            Ok(val) => {
                if !val.is_none() {
                    $config.$field = Some(val.extract().map_err(|_| {
                        crate::errors::configuration_type(format!(
                            "sheet option '{}' must be {}, got {}",
                            $key,
                            $type_desc,
                            pytype_name(&val)
                        ))
                    })?);
                }
            }
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyKeyError>($opts.py()) => {}
            Err(e) => return Err(e),
        }
    };
}

/// Helper: extract an optional dict field, run an extractor, and set it.
/// An explicitly-passed empty dict is still recorded as `Some(empty)` (not
/// dropped) — on the multi-sheet path this is a deliberate per-sheet "off
/// switch" that shadows a non-empty global default via `SheetConfig::merge_with`,
/// matching the existing `table_style` explicit-`None`-means-off precedent.
///
/// Missing-key handling follows `extract_scalar!`: only the `KeyError` that
/// means "option not given" is swallowed, and every other lookup failure
/// propagates.
macro_rules! extract_dict_field {
    ($opts:expr, $config:expr, $key:literal, $field:ident, $extractor:expr) => {
        match $opts.get_item($key) {
            Ok(val) => {
                if !val.is_none() {
                    let dict = val.cast::<pyo3::types::PyDict>().map_err(|_| {
                        crate::errors::configuration_type(format!(
                            "sheet option '{}' must be a dict, got {}",
                            $key,
                            pytype_name(&val)
                        ))
                    })?;
                    let extracted = $extractor(dict)?;
                    $config.$field = Some(extracted);
                }
            }
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyKeyError>($opts.py()) => {}
            Err(e) => return Err(e),
        }
    };
}

/// Helper: extract an optional list field, run an extractor, and set it.
/// See `extract_dict_field!` — an explicitly-passed empty list is likewise
/// kept as `Some(empty)` rather than dropped, and only a missing-key
/// `KeyError` is swallowed.
macro_rules! extract_list_field {
    ($opts:expr, $config:expr, $key:literal, $field:ident, $extractor:expr) => {
        match $opts.get_item($key) {
            Ok(val) => {
                if !val.is_none() {
                    let list = val.cast::<pyo3::types::PyList>().map_err(|_| {
                        crate::errors::configuration_type(format!(
                            "sheet option '{}' must be a list, got {}",
                            $key,
                            pytype_name(&val)
                        ))
                    })?;
                    let extracted = $extractor(list)?;
                    $config.$field = Some(extracted);
                }
            }
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyKeyError>($opts.py()) => {}
            Err(e) => return Err(e),
        }
    };
}

fn validate_sheet_option_keys(opts: &Bound<'_, pyo3::types::PyDict>) -> PyResult<()> {
    for key in opts.keys().iter() {
        let key_str: String = key.extract().map_err(|_| {
            crate::errors::configuration_type(format!(
                "Sheet option keys must be strings, got {}",
                pytype_name(&key)
            ))
        })?;
        if !SHEET_OPTION_NAMES.contains(&key_str.as_str()) {
            return Err(crate::errors::configuration(format!(
                "Unknown sheet option '{}'. Valid keys: {}",
                key_str,
                SHEET_OPTION_NAMES.join(", ")
            )));
        }
    }
    Ok(())
}

/// Reject any dict key not in `allowed`. Shared by the dict-form extractors
/// (comments, checkboxes, cells) that accept a handful of recognized keys and
/// previously dropped anything else silently. Delegates the actual
/// unknown-key policy to `types::reject_unknown_keys` (this crate's single
/// source of truth for that phrasing); this wrapper only does the PyO3 key
/// extraction.
fn reject_unknown_dict_keys(
    dict: &Bound<'_, pyo3::types::PyDict>,
    context: &str,
    allowed: &[&str],
) -> PyResult<()> {
    let mut key_strs: Vec<String> = Vec::with_capacity(dict.len());
    for key in dict.keys().iter() {
        let key_str: String = key.extract().map_err(|_| {
            crate::errors::configuration_type(format!(
                "{}: keys must be strings, got {}",
                context,
                pytype_name(&key)
            ))
        })?;
        key_strs.push(key_str);
    }
    types_reject_unknown_keys(key_strs.iter().map(String::as_str), context, None, allowed)
        .map_err(crate::errors::configuration)
}

/// Extract sheet info from a Python tuple (supports both 2-tuple and 3-tuple formats)
/// 2-tuple: (df, sheet_name)
/// 3-tuple: (df, sheet_name, options_dict)
pub(crate) fn extract_sheet_info<'py>(
    sheet_tuple: &Bound<'py, PyAny>,
) -> PyResult<(Bound<'py, PyAny>, String, SheetConfig)> {
    let len: usize = sheet_tuple.len()?;

    if !(2..=3).contains(&len) {
        return Err(crate::errors::configuration(
            format!(
                "Sheet tuple must have exactly 2 or 3 elements, got {}: (df, sheet_name[, options_dict])",
                len
            ),
        ));
    }

    let df = sheet_tuple.get_item(0)?;
    let sheet_name: String = extract_typed!(
        sheet_tuple.get_item(1)?,
        "a string",
        "sheet tuple element 1 (sheet name)"
    );

    let config = if len >= 3 {
        let opts = sheet_tuple.get_item(2)?;
        if opts.is_none() {
            return Ok((df, sheet_name, SheetConfig::default()));
        }
        let opts_dict = opts.cast::<pyo3::types::PyDict>().map_err(|_| {
            crate::errors::configuration_type(format!(
                "Sheet options must be a dict, got {}",
                pytype_name(&opts)
            ))
        })?;
        validate_sheet_option_keys(opts_dict)?;
        let mut config = SheetConfig::default();

        // Extract scalar fields
        extract_scalar!(opts, config, "header", header, "a bool");
        extract_scalar!(opts, config, "autofit", autofit, "a bool");
        extract_scalar!(opts, config, "freeze_panes", freeze_panes, "a bool");
        extract_scalar!(
            opts,
            config,
            "row_heights",
            row_heights,
            "a dict mapping row index (int) to height (number)"
        );
        extract_scalar!(opts, config, "table_name", table_name, "a string");

        // table_style needs special handling: None means "explicitly no style".
        // Lookup failures are split the same way `extract_scalar!` splits them.
        match opts.get_item("table_style") {
            Ok(val) => {
                if val.is_none() {
                    config.table_style = Some(None);
                } else {
                    config.table_style = Some(Some(extract_typed!(
                        val,
                        "a string",
                        "sheet option 'table_style'"
                    )));
                }
            }
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyKeyError>(opts.py()) => {}
            Err(e) => return Err(e),
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
        extract_dict_field!(opts, config, "checkboxes", checkboxes, extract_checkboxes);
        extract_dict_field!(opts, config, "textboxes", textboxes, extract_textboxes);
        extract_dict_field!(opts, config, "charts", charts, extract_charts);
        extract_dict_field!(opts, config, "sparklines", sparklines, extract_sparklines);

        extract_dict_field!(opts, config, "cells", cells, extract_cells);

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

/// Excel's maximum column index (zero-based; column XFD is the 16384th column).
const MAX_COLUMN_INDEX: i64 = 16_383;

/// Validate a resolved column_widths integer key against Excel's column range
/// (0..=16383). `label` is the key's original representation — the int
/// restringified, or the source string key — used to build the
/// `column_widths['<label>']: ...` context-rich error so int and string keys
/// share identical messages modulo that label.
fn validate_column_widths_index(i: i64, label: &str) -> PyResult<()> {
    if i < 0 {
        return Err(crate::errors::configuration(format!(
            "column_widths['{}']: must be a non-negative column index",
            label
        )));
    }
    if i > MAX_COLUMN_INDEX {
        return Err(crate::errors::configuration(format!(
            "column_widths['{}']: exceeds Excel's maximum column index ({}, i.e. column XFD)",
            label, MAX_COLUMN_INDEX
        )));
    }
    Ok(())
}

/// Extract column_widths from Python dict, supporting both integer and string keys.
/// Integer keys are column indices and are validated against Excel's column range
/// (0..=16383, i.e. up to XFD); the literal string key `"_all"` is a special
/// global-width cap applied to every data column (see `apply_column_widths`).
/// A string key other than `"_all"` must itself parse as an integer column
/// index and passes through the same range validation — it is not merely
/// forwarded verbatim, since the backend only recognizes numeric-string keys
/// and `"_all"`.
pub(crate) fn extract_column_widths(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, f64>> {
    let mut widths: HashMap<String, f64> = HashMap::new();
    for (k, v) in py_dict.iter() {
        let key_str = if let Ok(i) = k.extract::<i64>() {
            validate_column_widths_index(i, &i.to_string())?;
            i.to_string()
        } else if let Ok(s) = k.extract::<String>() {
            if s == "_all" {
                s
            } else {
                let i: i64 = s.parse().map_err(|_| {
                    crate::errors::configuration_type(format!(
                        "column_widths['{}']: must be an integer column index or the string \
                         '_all', got a non-numeric string",
                        s
                    ))
                })?;
                validate_column_widths_index(i, &s)?;
                i.to_string()
            }
        } else {
            let key_repr = k
                .str()
                .map(|s| s.to_string())
                .unwrap_or_else(|_| "?".to_string());
            return Err(crate::errors::configuration_type(format!(
                "column_widths['{}']: must be an integer column index or the string '_all', got {}",
                key_repr,
                pytype_name(&k)
            )));
        };
        let width = extract_typed!(v, "a number", "column_widths['{}']", key_str);
        widths.insert(key_str, width);
    }
    Ok(widths)
}

/// Extract header_format from Python dict
pub(crate) fn extract_header_format(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    pydict_to_hashmap(py_dict, "header_format")
}

/// Extract column_formats from Python dict (pattern -> format dict)
/// Uses IndexMap to preserve insertion order from Python dict
pub(crate) fn extract_column_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (pattern, fmt_dict) in py_dict.iter() {
        let pattern_str: String = extract_typed!(
            pattern,
            "a string column name or pattern",
            "column_formats key"
        );
        let inner_dict = fmt_dict.cast::<pyo3::types::PyDict>().map_err(|_| {
            crate::errors::configuration_type(format!(
                "column_formats['{}']: expected dict, got {}",
                pattern_str,
                pytype_name(&fmt_dict)
            ))
        })?;
        let inner = pydict_to_hashmap(inner_dict, &format!("column_formats['{}']", pattern_str))?;
        col_fmts.insert(pattern_str, inner);
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
        let col_str: String = extract_typed!(
            col_name,
            "a string column name or pattern",
            "conditional_formats key"
        );
        // Accept either a single dict or a list of dicts
        if let Ok(list) = fmt_value.cast::<pyo3::types::PyList>() {
            let mut configs = Vec::new();
            for (i, item) in list.iter().enumerate() {
                let d = item.cast::<pyo3::types::PyDict>().map_err(|_| {
                    crate::errors::configuration_type(format!(
                        "conditional_formats['{}']: list item {} must be a dict",
                        col_str, i
                    ))
                })?;
                configs.push(pydict_to_hashmap(
                    d,
                    &format!("conditional_formats['{}'] list item {}", col_str, i),
                )?);
            }
            cond_fmts.insert(col_str, configs);
        } else if let Ok(inner_dict) = fmt_value.cast::<pyo3::types::PyDict>() {
            let inner =
                pydict_to_hashmap(inner_dict, &format!("conditional_formats['{}']", col_str))?;
            cond_fmts.insert(col_str, vec![inner]);
        } else {
            return Err(crate::errors::configuration_type(format!(
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
        let col_str: String =
            extract_typed!(col_name, "a string column name", "formula_columns key");
        let formula_str: String = extract_typed!(
            formula,
            "a string formula template",
            "formula_columns['{}']",
            col_str
        );
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
        if !(2..=3).contains(&tuple_len) {
            return Err(crate::errors::configuration(format!(
                "merged_ranges tuple must have exactly 2 or 3 elements, got {}",
                tuple_len
            )));
        }

        let range_str: String = extract_typed!(
            item.get_item(0)?,
            "a string range",
            "merged_ranges tuple element 0"
        );
        let text: String = extract_typed!(
            item.get_item(1)?,
            "a string",
            "merged_ranges['{}'] element 1 (text)",
            range_str
        );

        let format_dict = if tuple_len >= 3 {
            let fmt_item = item.get_item(2)?;
            if !fmt_item.is_none() {
                let dict = fmt_item.cast::<pyo3::types::PyDict>().map_err(|_| {
                    crate::errors::configuration_type(format!(
                        "merged_ranges['{}']: format must be a dict, got {}",
                        range_str,
                        pytype_name(&fmt_item)
                    ))
                })?;
                Some(pydict_to_hashmap(
                    dict,
                    &format!("merged_ranges['{}'] format", range_str),
                )?)
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
        if !(2..=3).contains(&tuple_len) {
            return Err(crate::errors::configuration(format!(
                "hyperlinks tuple must have exactly 2 or 3 elements, got {}",
                tuple_len
            )));
        }

        let cell_ref: String = extract_typed!(
            item.get_item(0)?,
            "a string cell reference",
            "hyperlinks tuple element 0"
        );
        let url: String = extract_typed!(
            item.get_item(1)?,
            "a string URL",
            "hyperlinks['{}'] element 1 (url)",
            cell_ref
        );

        let display_text = if tuple_len >= 3 {
            let text_item = item.get_item(2)?;
            if !text_item.is_none() {
                Some(extract_typed!(
                    text_item,
                    "a string",
                    "hyperlinks['{}'] element 2 (display text)",
                    cell_ref
                ))
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
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_comments(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, Comment>> {
    let mut comments: IndexMap<String, Comment> = IndexMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = extract_typed!(cell_ref, "a string cell reference", "comments key");

        // Check if value is a dict or simple string
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            reject_unknown_dict_keys(
                inner_dict,
                &format!("comments['{}']", cell_str),
                &["text", "author"],
            )?;
            // Dict format: {'text': '...', 'author': '...'}
            let text_item = inner_dict.get_item("text")?.ok_or_else(|| {
                crate::errors::configuration(format!(
                    "Comment at '{}' missing 'text' key",
                    cell_str
                ))
            })?;
            let text: String =
                extract_typed!(text_item, "a string", "comments['{}']['text']", cell_str);
            let author: Option<String> = if let Ok(Some(a)) = inner_dict.get_item("author") {
                if !a.is_none() {
                    Some(extract_typed!(
                        a,
                        "a string",
                        "comments['{}']['author']",
                        cell_str
                    ))
                } else {
                    None
                }
            } else {
                None
            };
            comments.insert(cell_str, (text, author));
        } else {
            // Simple string format
            let text: String = extract_typed!(
                value,
                "a string or a dict with a 'text' key",
                "comments['{}']",
                cell_str
            );
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
        let col_str: String = extract_typed!(
            col_name,
            "a string column name or pattern",
            "validations key"
        );
        if let Ok(inner_dict) = config.cast::<pyo3::types::PyDict>() {
            let inner = pydict_to_hashmap(inner_dict, &format!("validations['{}']", col_str))?;
            validations.insert(col_str, inner);
        } else {
            return Err(crate::errors::configuration_type(format!(
                "validations['{}']: expected dict, got {}",
                col_str,
                pytype_name(&config)
            )));
        }
    }
    Ok(validations)
}

/// Extract rich_text from Python dict (cell_ref -> list of segments)
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_rich_text(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, Vec<RichTextSegment>>> {
    let mut rich_text: IndexMap<String, Vec<RichTextSegment>> = IndexMap::new();

    for (cell_ref, segments_list) in py_dict.iter() {
        let cell_str: String = extract_typed!(cell_ref, "a string cell reference", "rich_text key");
        let mut segments: Vec<RichTextSegment> = Vec::new();

        if let Ok(list) = segments_list.cast::<pyo3::types::PyList>() {
            for (idx, item) in list.iter().enumerate() {
                // Check if item is a tuple (text, format_dict) or just a string
                if let Ok(tuple) = item.cast::<pyo3::types::PyTuple>() {
                    if tuple.len() != 2 {
                        return Err(crate::errors::configuration(format!(
                            "rich_text['{}']: segment {} tuple must have exactly 2 elements, got {}",
                            cell_str,
                            idx,
                            tuple.len()
                        )));
                    }
                    let text: String = extract_typed!(
                        tuple.get_item(0)?,
                        "a string",
                        "rich_text['{}']: segment {} text",
                        cell_str,
                        idx
                    );
                    let fmt_item = tuple.get_item(1)?;
                    let format_dict = if fmt_item.is_none() {
                        None
                    } else {
                        let dict = fmt_item.cast::<pyo3::types::PyDict>().map_err(|_| {
                            crate::errors::configuration_type(format!(
                                "rich_text['{}']: segment {} format must be a dict, got {}",
                                cell_str,
                                idx,
                                pytype_name(&fmt_item)
                            ))
                        })?;
                        Some(pydict_to_hashmap(
                            dict,
                            &format!("rich_text['{}'] segment {} format", cell_str, idx),
                        )?)
                    };
                    segments.push((text, format_dict));
                } else if let Ok(text) = item.extract::<String>() {
                    // Plain string segment
                    segments.push((text, None));
                } else {
                    return Err(crate::errors::configuration_type(format!(
                        "rich_text['{}']: segment {} must be a string or tuple (text, format_dict), got {}",
                        cell_str,
                        idx,
                        pytype_name(&item)
                    )));
                }
            }
        } else {
            return Err(crate::errors::configuration_type(format!(
                "rich_text['{}']: expected list of segments, got {}",
                cell_str,
                pytype_name(&segments_list)
            )));
        }

        if !segments.is_empty() {
            rich_text.insert(cell_str, segments);
        }
    }

    Ok(rich_text)
}

/// Extract images from Python dict (cell_ref -> path or config dict)
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_images(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, ImageConfig>> {
    let mut images: IndexMap<String, ImageConfig> = IndexMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = extract_typed!(cell_ref, "a string cell reference", "images key");

        // Check if value is a dict or simple string (path)
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            // Dict format: {'path': '...', 'scale_width': 0.5, ...}
            let path_item = inner_dict.get_item("path")?.ok_or_else(|| {
                crate::errors::configuration(format!("Image at '{}' missing 'path' key", cell_str))
            })?;
            let path: String =
                extract_typed!(path_item, "a string path", "images['{}']['path']", cell_str);
            let mut options = pydict_to_hashmap(inner_dict, &format!("images['{}']", cell_str))?;
            options.remove("path");
            images.insert(
                cell_str,
                ImageConfig {
                    path,
                    options: Some(options),
                },
            );
        } else {
            // Simple string format (just path)
            let path: String = extract_typed!(
                value,
                "a string path or a dict with a 'path' key",
                "images['{}']",
                cell_str
            );
            images.insert(
                cell_str,
                ImageConfig {
                    path,
                    options: None,
                },
            );
        }
    }

    Ok(images)
}

/// Extract checkboxes from Python dict (cell_ref -> bool or config dict)
/// Simple form: {'A1': True}
/// Dict form: {'A1': {'checked': True, 'format': {...}}}
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_checkboxes(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, CheckboxConfig>> {
    let mut checkboxes: IndexMap<String, CheckboxConfig> = IndexMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String =
            extract_typed!(cell_ref, "a string cell reference", "checkboxes key");

        // Dict form must be tried before bool, since a dict would extract as False for bool otherwise.
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            reject_unknown_dict_keys(
                inner_dict,
                &format!("checkboxes['{}']", cell_str),
                &["checked", "format"],
            )?;
            let checked_item = inner_dict.get_item("checked")?.ok_or_else(|| {
                crate::errors::configuration(format!(
                    "checkboxes['{}'] dict missing 'checked' key",
                    cell_str
                ))
            })?;
            let checked: bool = extract_typed!(
                checked_item,
                "a bool",
                "checkboxes['{}']['checked']",
                cell_str
            );
            let format_dict = if let Ok(Some(fmt)) = inner_dict.get_item("format") {
                if fmt.is_none() {
                    None
                } else if let Ok(d) = fmt.cast::<pyo3::types::PyDict>() {
                    Some(pydict_to_hashmap(
                        d,
                        &format!("checkboxes['{}']['format']", cell_str),
                    )?)
                } else {
                    return Err(crate::errors::configuration_type(format!(
                        "checkboxes['{}']: 'format' must be a dict",
                        cell_str
                    )));
                }
            } else {
                None
            };
            checkboxes.insert(
                cell_str,
                CheckboxConfig {
                    checked,
                    format: format_dict,
                },
            );
        } else {
            let checked: bool = value.extract().map_err(|_| {
                crate::errors::configuration_type(format!(
                    "checkboxes['{}']: expected bool or dict, got {}",
                    cell_str,
                    pytype_name(&value)
                ))
            })?;
            checkboxes.insert(
                cell_str,
                CheckboxConfig {
                    checked,
                    format: None,
                },
            );
        }
    }

    Ok(checkboxes)
}

/// Extract textboxes from Python dict (cell_ref -> text or config dict)
/// Simple form: {'A1': 'Some text'}
/// Dict form: {'A1': {'text': 'Some text', 'width': 200, 'height': 100, 'font': {...}, ...}}
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_textboxes(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, TextboxConfig>> {
    let mut textboxes: IndexMap<String, TextboxConfig> = IndexMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = extract_typed!(cell_ref, "a string cell reference", "textboxes key");

        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            let text_item = inner_dict.get_item("text")?.ok_or_else(|| {
                crate::errors::configuration(format!(
                    "textboxes['{}'] dict missing 'text' key",
                    cell_str
                ))
            })?;
            let text: String =
                extract_typed!(text_item, "a string", "textboxes['{}']['text']", cell_str);
            let mut options = pydict_to_hashmap(inner_dict, &format!("textboxes['{}']", cell_str))?;
            options.remove("text");
            textboxes.insert(
                cell_str,
                TextboxConfig {
                    text,
                    options: Some(options),
                },
            );
        } else if let Ok(text) = value.extract::<String>() {
            textboxes.insert(
                cell_str,
                TextboxConfig {
                    text,
                    options: None,
                },
            );
        } else {
            return Err(crate::errors::configuration_type(format!(
                "textboxes['{}']: expected str or dict, got {}",
                cell_str,
                pytype_name(&value)
            )));
        }
    }

    Ok(textboxes)
}

/// Extract charts from Python dict (cell_ref -> chart options dict)
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_charts(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, ChartConfig>> {
    let mut charts: IndexMap<String, ChartConfig> = IndexMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = extract_typed!(cell_ref, "a string cell reference", "charts key");
        let inner_dict = value.cast::<pyo3::types::PyDict>().map_err(|_| {
            crate::errors::configuration_type(format!(
                "charts['{}']: expected dict, got {}",
                cell_str,
                pytype_name(&value)
            ))
        })?;
        let inner = pydict_to_hashmap(inner_dict, &format!("charts['{}']", cell_str))?;
        charts.insert(cell_str, inner);
    }

    Ok(charts)
}

/// Extract sparklines from Python dict (location ref -> sparkline options dict)
/// Uses IndexMap to preserve insertion order so output is reproducible.
pub(crate) fn extract_sparklines(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, SparklineConfig>> {
    let mut sparklines: IndexMap<String, SparklineConfig> = IndexMap::new();

    for (loc_ref, value) in py_dict.iter() {
        let loc_str: String =
            extract_typed!(loc_ref, "a string location reference", "sparklines key");
        let inner_dict = value.cast::<pyo3::types::PyDict>().map_err(|_| {
            crate::errors::configuration_type(format!(
                "sparklines['{}']: expected dict, got {}",
                loc_str,
                pytype_name(&value)
            ))
        })?;
        let inner = pydict_to_hashmap(inner_dict, &format!("sparklines['{}']", loc_str))?;
        sparklines.insert(loc_str, inner);
    }

    Ok(sparklines)
}

/// Look up an optional sub-value of a `cells` entry dict.
///
/// An explicitly-passed `None` reads as "not given", the same as an absent key.
/// The `cells` dict form is the only place a caller writes these keys out by
/// name, so a `None` produced by their own conditional expression must not be
/// harder to pass than omitting the key.
fn present_cell_field<'py>(
    d: &Bound<'py, pyo3::types::PyDict>,
    key: &str,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    Ok(d.get_item(key)?.filter(|v| !v.is_none()))
}

/// Extract an optional string sub-value of a `cells` entry dict.
///
/// Classified rather than a bare `extract()?`: that propagates PyO3's own
/// `TypeError`, which is outside the public hierarchy — the same escape
/// `extract_typed!` closes for the top-level extractors.
fn cell_string_field(
    d: &Bound<'_, pyo3::types::PyDict>,
    cell_ref: &str,
    key: &str,
) -> PyResult<Option<String>> {
    match present_cell_field(d, key)? {
        Some(v) => v.extract::<String>().map(Some).map_err(|_| {
            crate::errors::configuration_type(format!(
                "cells['{}']: '{}' must be a string, got {}",
                cell_ref,
                key,
                pytype_name(&v)
            ))
        }),
        None => Ok(None),
    }
}

/// Extract cells from Python dict (cell_ref -> value or {value, num_format, align_horizontal, ...})
pub(crate) fn extract_cells(py_dict: &Bound<'_, pyo3::types::PyDict>) -> PyResult<Vec<CellWrite>> {
    let mut cells = Vec::new();
    for (key, value) in py_dict.iter() {
        let cell_ref: String = extract_typed!(key, "a string cell reference", "cells key");
        let (row, col) = parse_cell_ref(&cell_ref).map_err(crate::errors::configuration)?;

        // Check if value is a dict with "value" and optional formatting keys
        if let Ok(d) = value.cast::<pyo3::types::PyDict>() {
            reject_unknown_dict_keys(
                d,
                &format!("cells['{}']", cell_ref),
                &[
                    "value",
                    "num_format",
                    "align_horizontal",
                    "align_vertical",
                    "wrap_text",
                ],
            )?;
            let val = d.get_item("value")?.ok_or_else(|| {
                crate::errors::configuration(format!(
                    "cells['{}'] dict missing 'value' key",
                    cell_ref
                ))
            })?;
            let num_fmt = cell_string_field(d, &cell_ref, "num_format")?;
            let align_h = cell_string_field(d, &cell_ref, "align_horizontal")?;
            if let Some(ref ah) = align_h {
                parse_horizontal_alignment(ah).map_err(crate::errors::configuration)?;
            }
            let align_v = cell_string_field(d, &cell_ref, "align_vertical")?;
            if let Some(ref av) = align_v {
                parse_vertical_alignment(av).map_err(crate::errors::configuration)?;
            }
            let wrap: bool = match present_cell_field(d, "wrap_text")? {
                Some(v) => v.extract::<bool>().map_err(|_| {
                    crate::errors::configuration_type(format!(
                        "cells['{}']: 'wrap_text' must be a bool, got {}",
                        cell_ref,
                        pytype_name(&v)
                    ))
                })?,
                None => false,
            };
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

#[cfg(test)]
mod sheet_option_name_tests {
    use super::SHEET_OPTION_NAMES;
    use crate::types::EffectiveOpts;

    /// Every complex feature option declared via `define_options!` must also be a
    /// recognized per-sheet option key. Otherwise `dfs_to_xlsx`'s
    /// `validate_sheet_option_keys` would reject an option that `df_to_xlsx`
    /// accepts — a silent feature gap on the multi-sheet path with no compile
    /// error. Adding a `define_options!` field auto-grows `COMPLEX_OPTION_NAMES`,
    /// failing this test until the name is added to `SHEET_OPTION_NAMES` too.
    #[test]
    fn every_complex_option_is_a_valid_sheet_option() {
        for &name in EffectiveOpts::COMPLEX_OPTION_NAMES {
            assert!(
                SHEET_OPTION_NAMES.contains(&name),
                "complex option '{}' is declared in define_options! but missing from \
                 SHEET_OPTION_NAMES — dfs_to_xlsx would reject it as an unknown \
                 per-sheet option",
                name
            );
        }
    }
}

// `reject_unknown_keys` unit tests and the `SheetConfig::merge_with`
// empty-dict/absent-fallback tests live in `types.rs`, the natural home of
// the shared `reject_unknown_keys` helper and `merge_with` respectively —
// see `types::reject_unknown_keys_tests` and
// `types::complex_option_presence_tests`.
