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

/// Type alias for image config: cell_ref -> image path or config dict
pub(crate) type ImageConfig = (String, Option<HashMap<String, Py<PyAny>>>); // (path, options)

/// Extracted and validated write options from Python parameters.
/// Used to eliminate duplication between df_to_xlsx and dfs_to_xlsx.
#[derive(Debug, Default)]
pub(crate) struct ExtractedOptions {
    pub(crate) column_widths: Option<HashMap<String, f64>>,
    pub(crate) header_format: Option<HashMap<String, Py<PyAny>>>,
    pub(crate) column_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>,
    pub(crate) conditional_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>,
    pub(crate) formula_columns: Option<IndexMap<String, String>>,
    pub(crate) merged_ranges: Option<Vec<MergedRange>>,
    pub(crate) hyperlinks: Option<Vec<Hyperlink>>,
    pub(crate) comments: Option<HashMap<String, Comment>>,
    pub(crate) validations: Option<IndexMap<String, ValidationConfig>>,
    pub(crate) rich_text: Option<HashMap<String, Vec<RichTextSegment>>>,
    pub(crate) images: Option<HashMap<String, ImageConfig>>,
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
    pub(crate) conditional_formats: Option<IndexMap<String, HashMap<String, Py<PyAny>>>>, // Column/pattern -> conditional format config (ordered)
    pub(crate) formula_columns: Option<IndexMap<String, String>>, // Column name -> formula template (ordered)
    pub(crate) merged_ranges: Option<Vec<MergedRange>>,           // (range, text, format)
    pub(crate) hyperlinks: Option<Vec<Hyperlink>>, // (cell, url, optional display_text)
    pub(crate) comments: Option<HashMap<String, Comment>>, // cell_ref -> (text, author)
    pub(crate) validations: Option<IndexMap<String, ValidationConfig>>, // column name/pattern -> validation config
    pub(crate) rich_text: Option<HashMap<String, Vec<RichTextSegment>>>, // cell_ref -> segments
    pub(crate) images: Option<HashMap<String, ImageConfig>>, // cell_ref -> (path, options)
}
