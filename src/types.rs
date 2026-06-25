//! Type definitions for xlsxturbo

use indexmap::IndexMap;
use pyo3::prelude::*;
use std::collections::HashMap;

/// Date formats by locale/order preference
/// ISO formats (YYYY-MM-DD) are always tried first as they're unambiguous
pub(crate) const DATE_PATTERNS_ISO: &[&str] = &[
    "%Y-%m-%d", // 2024-01-15
    "%Y/%m/%d", // 2024/01/15
];

/// European date formats (day first): DD-MM-YYYY
pub(crate) const DATE_PATTERNS_DMY: &[&str] = &[
    "%d-%m-%Y", // 15-01-2024
    "%d/%m/%Y", // 15/01/2024
];

/// US date formats (month first): MM-DD-YYYY
pub(crate) const DATE_PATTERNS_MDY: &[&str] = &[
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
    pub(crate) fn patterns(&self) -> Vec<&'static str> {
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
pub(crate) const DATETIME_PATTERNS: &[&str] = &[
    "%Y-%m-%dT%H:%M:%S",    // ISO 8601
    "%Y-%m-%d %H:%M:%S",    // Common format
    "%Y-%m-%dT%H:%M:%S%.f", // ISO 8601 with fractional seconds
    "%Y-%m-%d %H:%M:%S%.f", // With fractional seconds
];

/// Represents the detected type of a cell value
#[derive(Debug, Clone)]
pub(crate) enum CellValue {
    Empty,
    Integer(i64),
    Float(f64),
    Boolean(bool),
    Date(f64),     // Excel serial date
    DateTime(f64), // Excel serial datetime
    String(String),
}

/// Type alias for merged range tuple: (range_str, text, optional format_dict)
pub(crate) type MergedRange = (String, String, Option<HashMap<String, Py<PyAny>>>);

/// Type alias for hyperlink tuple: (cell_ref, url, optional display_text)
pub(crate) type Hyperlink = (String, String, Option<String>);

/// Type alias for comment: either simple text or dict with 'text' and optionally 'author'
pub(crate) type Comment = (String, Option<String>); // (text, author)

/// Type alias for validation: column name/pattern -> validation config
pub(crate) type ValidationConfig = HashMap<String, Py<PyAny>>;

/// Type alias for rich text segment: (text, optional format_dict) or just text
pub(crate) type RichTextSegment = (String, Option<HashMap<String, Py<PyAny>>>);

/// Image insertion config extracted from the Python API.
#[derive(Debug)]
pub(crate) struct ImageConfig {
    pub(crate) path: String,
    pub(crate) options: Option<HashMap<String, Py<PyAny>>>,
}

/// Checkbox insertion config extracted from the Python API.
#[derive(Debug)]
pub(crate) struct CheckboxConfig {
    pub(crate) checked: bool,
    pub(crate) format: Option<HashMap<String, Py<PyAny>>>,
}

/// Textbox insertion config extracted from the Python API.
#[derive(Debug)]
pub(crate) struct TextboxConfig {
    pub(crate) text: String,
    pub(crate) options: Option<HashMap<String, Py<PyAny>>>,
}

/// Type alias for native Excel chart config: cell_ref -> chart options dict
pub(crate) type ChartConfig = HashMap<String, Py<PyAny>>;

/// Type alias for sparkline config: location ref -> sparkline options dict
pub(crate) type SparklineConfig = HashMap<String, Py<PyAny>>;

/// Type alias for conditional format configs: column/pattern -> list of format config dicts
pub(crate) type ConditionalFormatConfigs = IndexMap<String, Vec<HashMap<String, Py<PyAny>>>>;

/// Represents a single cell write operation with optional formatting
#[derive(Debug)]
pub(crate) struct CellWrite {
    pub(crate) row: u32,
    pub(crate) col: u16,
    pub(crate) value: Py<PyAny>,
    pub(crate) num_format: Option<String>,
    pub(crate) align_horizontal: Option<String>,
    pub(crate) align_vertical: Option<String>,
    pub(crate) wrap_text: bool,
}

/// Infallible variant of `PyAny::get_type().name()` returning "unknown" on failure.
/// Used for error-message construction where we must produce a String even if the
/// name lookup itself errors (e.g., during another exception's formatting).
pub(crate) fn pytype_name(v: &Bound<'_, PyAny>) -> String {
    v.get_type()
        .name()
        .map_or_else(|_| "unknown".to_string(), |n| n.to_string())
}

/// Extract an optional typed value from an already-extracted option entry.
///
/// `entry` is the result of `map.get(key)` (works for any `HashMap`/`IndexMap`
/// of `Py<PyAny>`). A missing key or an explicit Python `None` yields `Ok(None)`.
/// A present value of the wrong type yields `Err(on_err(bound))`, letting the
/// caller build a context-rich message (e.g. embedding `pytype_name(bound)`).
///
/// This is the single shared implementation behind the per-feature
/// `*_field` extractor helpers in `apply/*` and `parse/formats.rs`.
pub(crate) fn extract_opt<'py, T, F>(
    py: Python<'py>,
    entry: Option<&Py<PyAny>>,
    on_err: F,
) -> Result<Option<T>, String>
where
    T: for<'a> FromPyObject<'a, 'py>,
    F: FnOnce(&Bound<'py, PyAny>) -> String,
{
    let Some(obj) = entry else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<T>().map(Some).map_err(|_| on_err(bound))
}

/// Convert a Python dict to a Rust `HashMap<String, Py<PyAny>>`.
///
/// Lives in the lowest layer so both `extract` (Python→Rust) and `apply`
/// (which re-reads nested option dicts at write time) can use it without
/// `apply` depending back up on `extract`.
pub(crate) fn pydict_to_hashmap(
    dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    let mut map = HashMap::new();
    for (k, v) in dict.iter() {
        map.insert(k.extract()?, v.unbind());
    }
    Ok(map)
}

/// Detect whether a Python object is a Polars or Pandas DataFrame.
/// Returns true for Polars, false for Pandas.
/// Errors if the object is neither.
pub(crate) fn is_polars_dataframe(df: &Bound<'_, PyAny>) -> Result<bool, String> {
    // Check the actual module to avoid misidentifying objects that happen to
    // have similar attributes (e.g., Pydantic models with .schema)
    let module = df
        .get_type()
        .getattr("__module__")
        .and_then(|m| m.extract::<String>())
        .unwrap_or_else(|_| String::new());

    if module.starts_with("polars") {
        Ok(true)
    } else if module.starts_with("pandas") {
        Ok(false)
    } else {
        Err(format!(
            "Unsupported DataFrame type: {}.{}. Expected pandas or polars DataFrame.",
            module,
            pytype_name(df)
        ))
    }
}

/// Extract column names from a DataFrame (Polars or Pandas).
pub(crate) fn extract_columns(
    df: &Bound<'_, PyAny>,
    is_polars: bool,
) -> Result<Vec<String>, String> {
    if is_polars {
        let cols = df
            .getattr("columns")
            .map_err(|e| format!("Failed to access DataFrame columns: {}", e))?;
        cols.extract().map_err(|e: pyo3::PyErr| e.to_string())
    } else {
        let cols = df
            .getattr("columns")
            .map_err(|e| format!("Failed to access DataFrame columns: {}", e))?;
        let col_list = cols
            .call_method0("tolist")
            .map_err(|e| format!("Failed to list DataFrame columns: {}", e))?;
        let py_list = col_list
            .cast::<pyo3::types::PyList>()
            .map_err(|e| e.to_string())?;
        py_list
            .iter()
            .map(|col| col.str().map(|s| s.to_string()).map_err(|e| e.to_string()))
            .collect()
    }
}

/// Per-sheet configuration options (all optional, defaults to global settings)
#[derive(Debug, Default)]
pub(crate) struct SheetConfig {
    pub(crate) header: Option<bool>,
    pub(crate) autofit: Option<bool>,
    pub(crate) table_style: Option<Option<String>>, // None = use default, Some(None) = explicitly no style
    pub(crate) freeze_panes: Option<bool>,
    pub(crate) column_widths: Option<HashMap<String, f64>>, // Keys: "0", "1", "_all" for global cap
    pub(crate) table_name: Option<String>,
    pub(crate) header_format: Option<HashMap<String, Py<PyAny>>>,
    pub(crate) row_heights: Option<HashMap<u32, f64>>,
    pub(crate) column_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>, // Pattern -> format dict (ordered)
    pub(crate) conditional_formats: Option<ConditionalFormatConfigs>, // Column/pattern -> list of conditional format configs
    pub(crate) formula_columns: Option<IndexMap<String, String>>, // Column name -> formula template (ordered)
    pub(crate) merged_ranges: Option<Vec<MergedRange>>,           // (range, text, format)
    pub(crate) hyperlinks: Option<Vec<Hyperlink>>, // (cell, url, optional display_text)
    pub(crate) comments: Option<HashMap<String, Comment>>, // cell_ref -> (text, author)
    pub(crate) validations: Option<IndexMap<String, ValidationConfig>>, // column name/pattern -> validation config
    pub(crate) rich_text: Option<HashMap<String, Vec<RichTextSegment>>>, // cell_ref -> segments
    pub(crate) images: Option<HashMap<String, ImageConfig>>,
    pub(crate) checkboxes: Option<HashMap<String, CheckboxConfig>>,
    pub(crate) textboxes: Option<HashMap<String, TextboxConfig>>,
    pub(crate) charts: Option<HashMap<String, ChartConfig>>, // cell_ref -> chart options
    pub(crate) sparklines: Option<HashMap<String, SparklineConfig>>, // location ref -> sparkline options
    pub(crate) cells: Option<Vec<CellWrite>>,
}

/// Scalar configuration for writing a single sheet.
/// Groups the simple config fields to reduce parameter counts in write_sheet_data
/// and apply_worksheet_features.
pub(crate) struct WriteConfig<'a> {
    pub(crate) include_header: bool,
    pub(crate) autofit: bool,
    pub(crate) table_style: Option<&'a str>,
    pub(crate) freeze_panes: bool,
    pub(crate) table_name: Option<&'a str>,
    pub(crate) row_heights: Option<&'a HashMap<u32, f64>>,
    pub(crate) constant_memory: bool,
}

/// Define the complex (feature) write-option fields in one place and generate
/// every structure that has to stay in lockstep with them:
///
/// - `ExtractedOptions` — owned `Option<T>` per field (global defaults).
/// - `EffectiveOpts<'a>` — borrowed `Option<&'a T>` per field (resolved, no clone of `Py<PyAny>`).
/// - `ExtractedOptions::as_effective` — single-sheet: borrow every field.
/// - `SheetConfig::merge_with` — multi-sheet: per-sheet value, else global fallback.
/// - `EffectiveOpts::present_complex_options` — names of the present fields,
///   so the `constant_memory` skip warning is derived, not hand-maintained.
///
/// Adding a feature option is now one line here (plus the matching field on the
/// hand-written `SheetConfig`, which `merge_with` references — a missing field
/// is a compile error, never a silent drift).
macro_rules! define_options {
    ($( $field:ident : $ty:ty ),+ $(,)?) => {
        /// Extracted and validated write options from Python parameters.
        /// Generated by `define_options!`.
        #[derive(Debug, Default)]
        pub(crate) struct ExtractedOptions {
            $( pub(crate) $field: Option<$ty>, )+
        }

        /// Resolved effective options for writing a single sheet (references
        /// only, avoids cloning `Py<PyAny>`). Generated by `define_options!`.
        pub(crate) struct EffectiveOpts<'a> {
            $( pub(crate) $field: Option<&'a $ty>, )+
        }

        impl ExtractedOptions {
            /// Create `EffectiveOpts` from this `ExtractedOptions` (all fields as references).
            pub(crate) fn as_effective(&self) -> EffectiveOpts<'_> {
                EffectiveOpts {
                    $( $field: self.$field.as_ref(), )+
                }
            }
        }

        impl SheetConfig {
            /// Merge per-sheet options with global defaults, returning `EffectiveOpts`.
            /// Per-sheet values take priority; global defaults are the fallback.
            pub(crate) fn merge_with<'a>(
                &'a self,
                global: &'a ExtractedOptions,
            ) -> EffectiveOpts<'a> {
                EffectiveOpts {
                    $( $field: self.$field.as_ref().or(global.$field.as_ref()), )+
                }
            }
        }

        impl EffectiveOpts<'_> {
            /// Every complex feature-option field name, in declaration order.
            /// Generated with the fields so it can never drift from the set;
            /// used by the `constant_memory` classification guard test to force
            /// a deliberate safe-vs-skipped decision when a field is added.
            /// Test-only: the runtime path uses `present_complex_options`.
            #[cfg(test)]
            pub(crate) const COMPLEX_OPTION_NAMES: &'static [&'static str] =
                &[$( stringify!($field) ),+];

            /// Names of the complex feature options currently present (`Some`),
            /// in field-declaration order. Generated alongside the fields so it
            /// can never drift from the set — used to derive the
            /// `constant_memory` skip warning.
            pub(crate) fn present_complex_options(&self) -> Vec<&'static str> {
                let mut present = Vec::new();
                $( if self.$field.is_some() { present.push(stringify!($field)); } )+
                present
            }
        }
    };
}

define_options! {
    column_widths: HashMap<String, f64>,
    header_format: HashMap<String, Py<PyAny>>,
    column_formats: IndexMap<String, HashMap<String, Py<PyAny>>>,
    conditional_formats: ConditionalFormatConfigs,
    formula_columns: IndexMap<String, String>,
    merged_ranges: Vec<MergedRange>,
    hyperlinks: Vec<Hyperlink>,
    comments: HashMap<String, Comment>,
    validations: IndexMap<String, ValidationConfig>,
    rich_text: HashMap<String, Vec<RichTextSegment>>,
    images: HashMap<String, ImageConfig>,
    checkboxes: HashMap<String, CheckboxConfig>,
    textboxes: HashMap<String, TextboxConfig>,
    charts: HashMap<String, ChartConfig>,
    sparklines: HashMap<String, SparklineConfig>,
    cells: Vec<CellWrite>,
}
