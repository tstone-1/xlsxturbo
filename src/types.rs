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
/// This is the single shared implementation behind [`extract_field`] and,
/// through it, every [`OptionMap`] typed accessor used by `apply/*` and
/// `parse/formats.rs`.
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

/// Extract an optional typed field, producing a uniform error message of the
/// form `"<context>: '<key>' must be <type_desc>, got <actual>"`.
///
/// Thin wrapper over [`extract_opt`] that centralizes the per-feature error
/// phrasing so call sites can't drift — some used to include the offending
/// type and some didn't. `context` is the already-formatted option locator
/// (e.g. `"charts['D2']"`). The sole caller is [`OptionMap::field`]; other
/// code should extract through an `OptionMap` instead of calling this
/// directly.
pub(crate) fn extract_field<'py, T>(
    py: Python<'py>,
    entry: Option<&Py<PyAny>>,
    context: &str,
    key: &str,
    type_desc: &str,
) -> Result<Option<T>, String>
where
    T: for<'a> FromPyObject<'a, 'py>,
{
    extract_opt(py, entry, |bound| {
        format!(
            "{}: '{}' must be {}, got {}",
            context,
            key,
            type_desc,
            pytype_name(bound)
        )
    })
}

/// A context-bound view over a raw `HashMap<String, Py<PyAny>>` option dict,
/// centralizing the "extract typed field, on wrong type produce a context-rich
/// error" pattern.
///
/// Before this existed, every blob-extracted feature (charts, sparklines,
/// validations, images/checkboxes/textboxes, conditional formats, format
/// dicts) hand-rolled its own near-identical family of `<feature>_string_field`
/// / `<feature>_bool_field` / ... free functions on top of [`extract_opt`],
/// each repeating the same "build the context string, call extract_opt, map
/// the error" boilerplate (~400 lines total across `apply/*` and
/// `parse/formats.rs`). `OptionMap` holds the context once and exposes typed
/// accessors, so a new call site is one line instead of a new wrapper
/// function. See `AGENTS.md` (7-touchpoint checklist, touchpoint 4) for when
/// to reach for this vs. eager typed extraction at extract time.
pub(crate) struct OptionMap<'py, 'm> {
    py: Python<'py>,
    map: &'m HashMap<String, Py<PyAny>>,
    context: String,
}

impl<'py, 'm> OptionMap<'py, 'm> {
    /// Build a view over `map` whose errors are prefixed with `context`
    /// (e.g. `"charts['D2']"` or `"textboxes['B2']: font"`).
    pub(crate) fn new(
        py: Python<'py>,
        map: &'m HashMap<String, Py<PyAny>>,
        context: String,
    ) -> Self {
        Self { py, map, context }
    }

    /// The context string every error from this view is prefixed with.
    pub(crate) fn context(&self) -> &str {
        &self.context
    }

    /// The GIL token this view was built with, for call sites that need to
    /// hand off to a helper taking `Python<'py>` directly (e.g. a nested
    /// format-dict parser).
    pub(crate) fn py(&self) -> Python<'py> {
        self.py
    }

    /// Raw entry lookup, for call sites that need to hand-extract (e.g. a
    /// required field with bespoke wording, or a value that may be either a
    /// string or a number).
    pub(crate) fn get(&self, key: &str) -> Option<&'m Py<PyAny>> {
        self.map.get(key)
    }

    fn field<T>(&self, key: &str, type_desc: &str) -> Result<Option<T>, String>
    where
        T: for<'a> FromPyObject<'a, 'py>,
    {
        extract_field(self.py, self.map.get(key), &self.context, key, type_desc)
    }

    /// Extract an optional string field. Missing or `None` yields `Ok(None)`.
    pub(crate) fn string(&self, key: &str) -> Result<Option<String>, String> {
        self.field(key, "a string")
    }

    /// Extract an optional bool field. Missing or `None` yields `Ok(None)`.
    pub(crate) fn bool(&self, key: &str) -> Result<Option<bool>, String> {
        self.field(key, "a bool")
    }

    /// Extract an optional f64 (number) field. Missing or `None` yields `Ok(None)`.
    pub(crate) fn f64(&self, key: &str) -> Result<Option<f64>, String> {
        self.field(key, "a number")
    }

    /// Extract an optional i64 (integer) field. Missing or `None` yields `Ok(None)`.
    pub(crate) fn i64(&self, key: &str) -> Result<Option<i64>, String> {
        self.field(key, "an integer")
    }

    /// Extract an optional u32 (non-negative integer) field.
    pub(crate) fn u32(&self, key: &str) -> Result<Option<u32>, String> {
        self.field(key, "a non-negative integer")
    }

    /// Extract an optional u8 (0-255 integer) field.
    pub(crate) fn u8(&self, key: &str) -> Result<Option<u8>, String> {
        self.field(key, "an integer in the range 0-255")
    }

    /// Extract a required string field: missing/`None` is an error naming the key.
    pub(crate) fn required_string(&self, key: &str) -> Result<String, String> {
        self.string(key)?
            .ok_or_else(|| format!("{}: missing '{}' key", self.context, key))
    }

    /// Extract a required f64 field: missing/`None` is an error naming the key.
    pub(crate) fn required_f64(&self, key: &str) -> Result<f64, String> {
        self.f64(key)?
            .ok_or_else(|| format!("{}: missing '{}' key", self.context, key))
    }

    /// Extract an optional nested-dict field into a plain `HashMap`. Missing
    /// or `None` yields `Ok(None)`; a present non-dict value is a context-rich
    /// error naming the key and the received type.
    pub(crate) fn dict(&self, key: &str) -> Result<Option<HashMap<String, Py<PyAny>>>, String> {
        let Some(obj) = self.map.get(key) else {
            return Ok(None);
        };
        let bound = obj.bind(self.py);
        if bound.is_none() {
            return Ok(None);
        }
        let inner = bound.cast::<pyo3::types::PyDict>().map_err(|_| {
            format!(
                "{}: '{}' must be a dict, got {}",
                self.context,
                key,
                pytype_name(bound)
            )
        })?;
        pydict_to_hashmap(inner)
            .map(Some)
            .map_err(|e| format!("{}: {}", self.context, e))
    }

    /// Reject any key not in `allowed`, using this view's context and no
    /// qualifier (see [`reject_unknown_keys`]).
    pub(crate) fn reject_unknown(&self, allowed: &[&str]) -> Result<(), String> {
        reject_unknown_keys(
            self.map.keys().map(String::as_str),
            &self.context,
            None,
            allowed,
        )
    }

    /// Reject any key not in `allowed`, using this view's context and
    /// `qualifier` (e.g. a resolved conditional-format type) to narrow the
    /// "Valid for `<qualifier>`" listing (see [`reject_unknown_keys`]).
    pub(crate) fn reject_unknown_for(
        &self,
        qualifier: &str,
        allowed: &[&str],
    ) -> Result<(), String> {
        reject_unknown_keys(
            self.map.keys().map(String::as_str),
            &self.context,
            Some(qualifier),
            allowed,
        )
    }
}

/// Reject the first key not in `allowed`, producing a context-rich error that
/// names the feature/ref and lists valid keys.
///
/// Single source of truth for the "unknown option" phrasing used across
/// `extract.rs` (comments/checkboxes/cells dict forms) and, via
/// [`OptionMap::reject_unknown`]/[`OptionMap::reject_unknown_for`], every
/// blob-extracted `apply/*` feature (charts, sparklines, validations,
/// images/checkboxes/textboxes, conditional formats) and `parse/formats.rs`.
/// `qualifier`, when present, names the more specific thing `allowed` applies
/// to (e.g. a conditional format type) and is rendered as "Valid for
/// `<qualifier>`"; when `None` it renders as plain "Valid".
///
/// Pure string-key logic (no PyO3 types) so it is unit-testable without a
/// Python interpreter (this crate's `cargo test` does not embed one); callers
/// that start from a `PyDict` extract the key strings first and delegate the
/// actual policy here.
pub(crate) fn reject_unknown_keys<'a, I: IntoIterator<Item = &'a str>>(
    keys: I,
    context: &str,
    qualifier: Option<&str>,
    allowed: &[&str],
) -> Result<(), String> {
    for key in keys {
        if !allowed.contains(&key) {
            let valid_label = match qualifier {
                Some(q) => format!("Valid for {}", q),
                None => "Valid".to_string(),
            };
            return Err(format!(
                "{}: unknown option '{}'. {}: {}",
                context,
                key,
                valid_label,
                allowed.join(", ")
            ));
        }
    }
    Ok(())
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
    // The following feature maps use `IndexMap` (not `HashMap`) so their
    // iteration order follows Python dict insertion order — a `HashMap`'s
    // random iteration order would make generated workbooks non-reproducible
    // byte-for-byte across runs (the XML parts list objects in insertion
    // order).
    pub(crate) comments: Option<IndexMap<String, Comment>>, // cell_ref -> (text, author)
    pub(crate) validations: Option<IndexMap<String, ValidationConfig>>, // column name/pattern -> validation config
    pub(crate) rich_text: Option<IndexMap<String, Vec<RichTextSegment>>>, // cell_ref -> segments
    pub(crate) images: Option<IndexMap<String, ImageConfig>>,
    pub(crate) checkboxes: Option<IndexMap<String, CheckboxConfig>>,
    pub(crate) textboxes: Option<IndexMap<String, TextboxConfig>>,
    pub(crate) charts: Option<IndexMap<String, ChartConfig>>, // cell_ref -> chart options
    pub(crate) sparklines: Option<IndexMap<String, SparklineConfig>>, // location ref -> sparkline options
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

/// Minimal "is this collection empty" trait so `present_complex_options` can
/// distinguish "present but empty" from "present with content". Per the
/// per-sheet empty-dict/list override (see `extract_dict_field!` /
/// `extract_list_field!` in `extract.rs`), an explicitly-passed empty dict or
/// list is a deliberate "off switch" that shadows a global default with
/// `Some(empty)` rather than `None`. Left unfiltered, that would make the
/// `constant_memory` skip warning fire for an option that does nothing
/// anyway. Implemented for every collection type used by a complex field
/// below (`HashMap`, `IndexMap`, `Vec`).
pub(crate) trait ComplexOptionValue {
    /// True if this option's value has no entries and thus has nothing to apply.
    fn is_empty_value(&self) -> bool;
}

impl<K, V> ComplexOptionValue for HashMap<K, V> {
    fn is_empty_value(&self) -> bool {
        self.is_empty()
    }
}

impl<K, V> ComplexOptionValue for IndexMap<K, V> {
    fn is_empty_value(&self) -> bool {
        self.is_empty()
    }
}

impl<T> ComplexOptionValue for Vec<T> {
    fn is_empty_value(&self) -> bool {
        self.is_empty()
    }
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

            /// Names of the complex feature options currently present (`Some`)
            /// AND non-empty, in field-declaration order. An explicitly-passed
            /// empty dict/list is a deliberate no-op "off switch" (see
            /// `ComplexOptionValue`), so it is excluded here — it has nothing
            /// for `constant_memory` to skip. Generated alongside the fields so
            /// it can never drift from the set — used to derive the
            /// `constant_memory` skip warning.
            pub(crate) fn present_complex_options(&self) -> Vec<&'static str> {
                let mut present = Vec::new();
                $( if self.$field.is_some_and(|v| !v.is_empty_value()) { present.push(stringify!($field)); } )+
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
    comments: IndexMap<String, Comment>,
    validations: IndexMap<String, ValidationConfig>,
    rich_text: IndexMap<String, Vec<RichTextSegment>>,
    images: IndexMap<String, ImageConfig>,
    checkboxes: IndexMap<String, CheckboxConfig>,
    textboxes: IndexMap<String, TextboxConfig>,
    charts: IndexMap<String, ChartConfig>,
    sparklines: IndexMap<String, SparklineConfig>,
    cells: Vec<CellWrite>,
}

#[cfg(test)]
mod complex_option_presence_tests {
    use super::{ExtractedOptions, SheetConfig};
    use std::collections::HashMap;

    /// An explicitly-passed empty per-sheet dict/list must shadow ("turn off")
    /// a non-empty global default rather than being dropped and falling back
    /// to it. `column_widths` (`HashMap<String, f64>`) needs no PyO3 types, so
    /// it stands in for every complex option's `merge_with` behavior.
    #[test]
    fn empty_per_sheet_value_overrides_non_empty_global_after_merge() {
        let mut global = ExtractedOptions::default();
        let mut global_widths = HashMap::new();
        global_widths.insert("0".to_string(), 20.0);
        global.column_widths = Some(global_widths);

        let sheet = SheetConfig {
            column_widths: Some(HashMap::new()),
            ..Default::default()
        };

        let effective = sheet.merge_with(&global);
        let widths = effective.column_widths.expect(
            "explicit empty per-sheet dict must merge to Some(empty), not fall back to global",
        );
        assert!(
            widths.is_empty(),
            "per-sheet empty dict must win over the non-empty global default"
        );
    }

    /// A present-but-empty complex option is a no-op "off switch" and must not
    /// trigger the `constant_memory` skip warning for a feature that does
    /// nothing anyway.
    #[test]
    fn present_complex_options_excludes_empty_collections() {
        let opts = ExtractedOptions {
            column_widths: Some(HashMap::new()),
            ..Default::default()
        };
        let effective = opts.as_effective();
        assert!(
            !effective
                .present_complex_options()
                .contains(&"column_widths"),
            "an explicitly-empty column_widths dict has nothing to apply and \
             should not appear as 'present' for the constant_memory warning"
        );

        let mut opts_non_empty = ExtractedOptions::default();
        let mut widths = HashMap::new();
        widths.insert("0".to_string(), 10.0);
        opts_non_empty.column_widths = Some(widths);
        let effective_non_empty = opts_non_empty.as_effective();
        assert!(
            effective_non_empty
                .present_complex_options()
                .contains(&"column_widths"),
            "a non-empty column_widths dict must still be reported as present"
        );
    }

    /// When a sheet does not mention the option at all, the global default
    /// still applies (merge_with's ordinary fallback path is unaffected by
    /// the empty-dict "off switch" behavior above, which only changes
    /// behavior for an explicitly-present value). Moved here from
    /// `extract.rs`'s former `sheet_config_merge_tests` module, which
    /// duplicated `merge_with` coverage that belongs with its owner.
    #[test]
    fn absent_per_sheet_column_widths_falls_back_to_global() {
        let mut global = ExtractedOptions::default();
        let mut global_widths = HashMap::new();
        global_widths.insert("0".to_string(), 20.0);
        global.column_widths = Some(global_widths);

        let sheet = SheetConfig::default();
        let effective = sheet.merge_with(&global);
        let widths = effective
            .column_widths
            .expect("no per-sheet value present, global default should apply");
        assert_eq!(widths.get("0"), Some(&20.0));
    }
}

#[cfg(test)]
mod reject_unknown_keys_tests {
    use super::reject_unknown_keys;

    /// comments['<ref>'] dict form only accepts 'text' and 'author'; a typo'd
    /// key must be named in the error alongside the valid keys.
    #[test]
    fn comments_dict_rejects_typo_key() {
        let keys = ["text", "auhtor"];
        let err =
            reject_unknown_keys(keys, "comments['A1']", None, &["text", "author"]).unwrap_err();
        assert!(err.contains("unknown option 'auhtor'"), "{}", err);
        assert!(err.contains("Valid: text, author"), "{}", err);
    }

    /// checkboxes['<ref>'] dict form only accepts 'checked' and 'format'.
    #[test]
    fn checkboxes_dict_rejects_typo_key() {
        let keys = ["checked", "fromat"];
        let err = reject_unknown_keys(keys, "checkboxes['A1']", None, &["checked", "format"])
            .unwrap_err();
        assert!(err.contains("unknown option 'fromat'"), "{}", err);
        assert!(err.contains("Valid: checked, format"), "{}", err);
    }

    /// cells['<ref>'] dict form: a stray key like 'bold' (which belongs in a
    /// format dict elsewhere, not here) must be rejected, not silently dropped.
    #[test]
    fn cells_dict_rejects_stray_bold_key() {
        let keys = ["value", "bold"];
        let allowed = &[
            "value",
            "num_format",
            "align_horizontal",
            "align_vertical",
            "wrap_text",
        ];
        let err = reject_unknown_keys(keys, "cells['A1']", None, allowed).unwrap_err();
        assert!(err.contains("unknown option 'bold'"), "{}", err);
        assert!(
            err.contains("Valid: value, num_format, align_horizontal, align_vertical, wrap_text"),
            "{}",
            err
        );
    }

    #[test]
    fn accepts_all_valid_keys() {
        let keys = ["value", "num_format"];
        let allowed = &[
            "value",
            "num_format",
            "align_horizontal",
            "align_vertical",
            "wrap_text",
        ];
        assert!(reject_unknown_keys(keys, "cells['A1']", None, allowed).is_ok());
    }

    /// A typo'd key (e.g. "min_colour" instead of "min_color") must be
    /// rejected by name, with the valid keys listed for the resolved
    /// conditional-format type via `qualifier`.
    #[test]
    fn qualifier_is_included_when_present() {
        let keys = ["type", "min_colour"];
        let err = reject_unknown_keys(
            keys,
            "conditional_formats['A:A']",
            Some("2_color_scale"),
            &["type", "min_color", "max_color"],
        )
        .unwrap_err();
        assert!(
            err.contains("unknown option 'min_colour'"),
            "error should name the bad key: {}",
            err
        );
        assert!(
            err.contains("Valid for 2_color_scale: type, min_color, max_color"),
            "error should list the valid keys under the qualifier: {}",
            err
        );
        assert!(
            err.contains("conditional_formats['A:A']"),
            "error should include the range: {}",
            err
        );
    }
}
