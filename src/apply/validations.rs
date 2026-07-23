//! Data validation application helpers.

use crate::parse::matches_pattern;
use crate::types::{pytype_name, OptionMap, ValidationConfig};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{DataValidation, DataValidationErrorStyle, Worksheet};

/// Every validation type accepts `type` plus the shared input/error
/// title/message keys; only the type-specific keys (e.g. `values`, or
/// `min`/`max`) differ. Combine them here instead of indexing into a shared
/// keys array positionally at each call site.
fn keys_with(extra: &[&'static str]) -> Vec<&'static str> {
    let mut keys = vec!["type"];
    keys.extend_from_slice(extra);
    keys.extend_from_slice(&[
        "input_title",
        "input_message",
        "error_title",
        "error_message",
    ]);
    keys
}

/// Extract an `i32` field for `whole_number` validation's `min`/`max`.
/// rust_xlsxwriter's whole-number validation is bounded to the `i32` range;
/// a Python int outside it (e.g. `3_000_000_000`) previously surfaced as the
/// generic "'max' must be an integer, got int" — misleading, since it *is*
/// an integer, just out of range for this validation type. Distinguish that
/// case from an actual type mismatch. Kept bespoke (not an `OptionMap`
/// method) because of that out-of-range wording, which no other feature needs.
fn validation_i32_field(view: &OptionMap<'_, '_>, key: &str, default: i32) -> Result<i32, String> {
    let Some(entry) = view.get(key) else {
        return Ok(default);
    };
    let bound = entry.bind(view.py());
    if bound.is_none() {
        return Ok(default);
    }
    if let Ok(v) = bound.extract::<i32>() {
        return Ok(v);
    }
    if let Ok(v) = bound.extract::<i64>() {
        return Err(format!(
            "{}: '{}' must be within the i32 range ({}..={}) for \
             whole_number validation, got {}",
            view.context(),
            key,
            i32::MIN,
            i32::MAX,
            v
        ));
    }
    // A Python int too large even for i64 (e.g. 2**70) still needs the
    // i32-range message, not the misleading generic "must be an integer"
    // fallback below (it *is* an integer, just wildly out of range). The
    // value itself may be too large to embed cleanly, so the message omits
    // it rather than reaching for `str()`.
    if bound.is_instance_of::<pyo3::types::PyInt>() {
        return Err(format!(
            "{}: '{}' must be within the i32 range ({}..={}) for \
             whole_number validation, got an out-of-range integer",
            view.context(),
            key,
            i32::MIN,
            i32::MAX
        ));
    }
    Err(format!(
        "{}: '{}' must be an integer, got {}",
        view.context(),
        key,
        pytype_name(bound)
    ))
}

/// Build the `DataValidation` for a single validation config: rejects unknown
/// keys, dispatches by `type` to the type-specific rule, and layers on the
/// optional input/error message. Called once per `(col_pattern, config)`
/// pair — not once per matched column — since the built value is identical
/// for every column the pattern matches.
fn build_validation(view: &OptionMap<'_, '_>, col_pattern: &str) -> Result<DataValidation, String> {
    let val_type: String = view
        .get("type")
        .ok_or_else(|| format!("validations['{}']: missing 'type' key", col_pattern))?
        .bind(view.py())
        .extract()
        .map_err(|e| format!("validations['{}']: invalid 'type': {}", col_pattern, e))?;

    let validation = match val_type.to_lowercase().as_str() {
        "list" => {
            view.reject_unknown(&keys_with(&["values"]))?;
            // List validation: dropdown with values
            let values: Vec<String> = view
                .get("values")
                .ok_or_else(|| {
                    format!(
                        "validations['{}']: list type requires 'values'",
                        col_pattern
                    )
                })?
                .bind(view.py())
                .extract()
                .map_err(|e| format!("validations['{}']: invalid 'values': {}", col_pattern, e))?;

            // Check Excel's 255 character limit for list validation.
            // Count characters, not bytes — Excel's limit is on characters.
            let total_chars: usize = values.iter().map(|s| s.chars().count()).sum::<usize>()
                + values.len().saturating_sub(1); // commas between items
            if total_chars > 255 {
                return Err(format!(
                    "validations['{}']: list values exceed Excel's 255 character limit ({} chars). \
                     Use fewer or shorter values.",
                    col_pattern, total_chars
                ));
            }

            let values_refs: Vec<&str> = values.iter().map(|s| s.as_str()).collect();
            DataValidation::new()
                .allow_list_strings(&values_refs)
                .map_err(|e| format!("Failed to create list validation: {}", e))?
        }
        "whole_number" | "whole" | "integer" => {
            view.reject_unknown(&keys_with(&["min", "max"]))?;
            // Whole number validation with between rule
            let min = validation_i32_field(view, "min", i32::MIN)?;
            let max = validation_i32_field(view, "max", i32::MAX)?;
            DataValidation::new()
                .allow_whole_number(rust_xlsxwriter::DataValidationRule::Between(min, max))
        }
        "decimal" | "number" => {
            view.reject_unknown(&keys_with(&["min", "max"]))?;
            // Decimal validation with between rule
            let min = view.f64("min")?.unwrap_or(f64::MIN);
            let max = view.f64("max")?.unwrap_or(f64::MAX);
            DataValidation::new()
                .allow_decimal_number(rust_xlsxwriter::DataValidationRule::Between(min, max))
        }
        "text_length" | "textlength" | "length" => {
            view.reject_unknown(&keys_with(&["min", "max"]))?;
            // Text length validation with between rule
            let min = view.u32("min")?.unwrap_or(0);
            let max = view.u32("max")?.unwrap_or(u32::MAX);
            DataValidation::new()
                .allow_text_length(rust_xlsxwriter::DataValidationRule::Between(min, max))
        }
        _ => {
            return Err(format!(
                "Unknown validation type '{}'. Valid types: list, whole_number, decimal, text_length",
                val_type
            ));
        }
    };

    // Add optional input message
    let validation = if let Some(msg) = view.string("input_message")? {
        let title = view.string("input_title")?.unwrap_or_default();
        validation
            .set_input_title(&title)
            .map_err(|e| format!("Failed to set input title: {}", e))?
            .set_input_message(&msg)
            .map_err(|e| format!("Failed to set input message: {}", e))?
    } else {
        validation
    };

    // Add optional error message
    let validation = if let Some(msg) = view.string("error_message")? {
        let title = view.string("error_title")?.unwrap_or_default();
        validation
            .set_error_title(&title)
            .map_err(|e| format!("Failed to set error title: {}", e))?
            .set_error_message(&msg)
            .map_err(|e| format!("Failed to set error message: {}", e))?
            .set_error_style(DataValidationErrorStyle::Stop)
    } else {
        validation
    };

    Ok(validation)
}

/// Apply data validations to worksheet
pub(crate) fn apply_validations(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_start_row: u32,
    data_end_row: u32,
    validations: &IndexMap<String, ValidationConfig>,
) -> Result<(), String> {
    for (col_pattern, config) in validations {
        // Find matching columns
        let col_indices: Vec<u16> = columns
            .iter()
            .enumerate()
            .filter(|(_, name)| matches_pattern(name, col_pattern))
            .map(|(idx, _)| idx as u16) // safe: col_count already validated via u16::try_from
            .collect();

        if col_indices.is_empty() {
            return Err(format!(
                "validations['{}']: pattern matched no columns",
                col_pattern
            ));
        }

        // Unknown-key rejection, type dispatch, and message construction all
        // run exactly once per (col_pattern, config) pair here — not once per
        // matched column — since the resulting DataValidation is identical
        // for every column the pattern matches; only the worksheet write
        // below repeats per column.
        let view = OptionMap::new(py, config, format!("validations['{}']", col_pattern));
        let validation = build_validation(&view, col_pattern)?;

        for col_idx in col_indices {
            worksheet
                .add_data_validation(data_start_row, col_idx, data_end_row, col_idx, &validation)
                .map_err(|e| format!("Failed to add validation: {}", e))?;
        }
    }

    Ok(())
}
