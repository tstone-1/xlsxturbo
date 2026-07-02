//! Data validation application helpers.

use crate::parse::matches_pattern;
use crate::types::{
    extract_field, pytype_name, reject_unknown_keys as types_reject_unknown_keys, ValidationConfig,
};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{DataValidation, DataValidationErrorStyle, Worksheet};

/// Extract an optional string field from a validation config dict.
/// Returns Ok(None) for missing or Python-None values. Wrong types produce an error.
fn validation_string_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
) -> Result<Option<String>, String> {
    extract_field(
        py,
        config.get(key),
        &format!("validations['{}']", col_pattern),
        key,
        "a string",
    )
}

/// Delegates to the shared `types::reject_unknown_keys` (no behavior change:
/// no qualifier, same "unknown option ... Valid: ..." phrasing this already
/// used).
fn reject_unknown_keys(
    config: &ValidationConfig,
    col_pattern: &str,
    allowed: &[&str],
) -> Result<(), String> {
    types_reject_unknown_keys(
        config.keys().map(String::as_str),
        &format!("validations['{}']", col_pattern),
        None,
        allowed,
    )
}

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
/// case from an actual type mismatch.
fn validation_i32_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
    default: i32,
) -> Result<i32, String> {
    let Some(entry) = config.get(key) else {
        return Ok(default);
    };
    let bound = entry.bind(py);
    if bound.is_none() {
        return Ok(default);
    }
    if let Ok(v) = bound.extract::<i32>() {
        return Ok(v);
    }
    if let Ok(v) = bound.extract::<i64>() {
        return Err(format!(
            "validations['{}']: '{}' must be within the i32 range ({}..={}) for \
             whole_number validation, got {}",
            col_pattern,
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
            "validations['{}']: '{}' must be within the i32 range ({}..={}) for \
             whole_number validation, got an out-of-range integer",
            col_pattern,
            key,
            i32::MIN,
            i32::MAX
        ));
    }
    Err(format!(
        "validations['{}']: '{}' must be an integer, got {}",
        col_pattern,
        key,
        pytype_name(bound)
    ))
}

fn validation_u32_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
    default: u32,
) -> Result<u32, String> {
    Ok(extract_field(
        py,
        config.get(key),
        &format!("validations['{}']", col_pattern),
        key,
        "a non-negative integer",
    )?
    .unwrap_or(default))
}

fn validation_f64_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
    default: f64,
) -> Result<f64, String> {
    Ok(extract_field(
        py,
        config.get(key),
        &format!("validations['{}']", col_pattern),
        key,
        "a number",
    )?
    .unwrap_or(default))
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
            continue;
        }

        // Get validation type
        let val_type: String = config
            .get("type")
            .ok_or_else(|| format!("validations['{}']: missing 'type' key", col_pattern))?
            .bind(py)
            .extract()
            .map_err(|e| format!("validations['{}']: invalid 'type': {}", col_pattern, e))?;

        for col_idx in col_indices {
            let validation = match val_type.to_lowercase().as_str() {
                "list" => {
                    reject_unknown_keys(config, col_pattern, &keys_with(&["values"]))?;
                    // List validation: dropdown with values
                    let values: Vec<String> = config
                        .get("values")
                        .ok_or_else(|| {
                            format!(
                                "validations['{}']: list type requires 'values'",
                                col_pattern
                            )
                        })?
                        .bind(py)
                        .extract()
                        .map_err(|e| {
                            format!("validations['{}']: invalid 'values': {}", col_pattern, e)
                        })?;

                    // Check Excel's 255 character limit for list validation.
                    // Count characters, not bytes — Excel's limit is on characters.
                    let total_chars: usize =
                        values.iter().map(|s| s.chars().count()).sum::<usize>()
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
                    reject_unknown_keys(config, col_pattern, &keys_with(&["min", "max"]))?;
                    // Whole number validation with between rule
                    let min = validation_i32_field(py, config, col_pattern, "min", i32::MIN)?;
                    let max = validation_i32_field(py, config, col_pattern, "max", i32::MAX)?;
                    DataValidation::new()
                        .allow_whole_number(rust_xlsxwriter::DataValidationRule::Between(min, max))
                }
                "decimal" | "number" => {
                    reject_unknown_keys(config, col_pattern, &keys_with(&["min", "max"]))?;
                    // Decimal validation with between rule
                    let min = validation_f64_field(py, config, col_pattern, "min", f64::MIN)?;
                    let max = validation_f64_field(py, config, col_pattern, "max", f64::MAX)?;
                    DataValidation::new().allow_decimal_number(
                        rust_xlsxwriter::DataValidationRule::Between(min, max),
                    )
                }
                "text_length" | "textlength" | "length" => {
                    reject_unknown_keys(config, col_pattern, &keys_with(&["min", "max"]))?;
                    // Text length validation with between rule
                    let min = validation_u32_field(py, config, col_pattern, "min", 0)?;
                    let max = validation_u32_field(py, config, col_pattern, "max", u32::MAX)?;
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
            let validation = if let Some(msg) =
                validation_string_field(py, config, col_pattern, "input_message")?
            {
                let title = validation_string_field(py, config, col_pattern, "input_title")?
                    .unwrap_or_default();
                validation
                    .set_input_title(&title)
                    .map_err(|e| format!("Failed to set input title: {}", e))?
                    .set_input_message(&msg)
                    .map_err(|e| format!("Failed to set input message: {}", e))?
            } else {
                validation
            };

            // Add optional error message
            let validation = if let Some(msg) =
                validation_string_field(py, config, col_pattern, "error_message")?
            {
                let title = validation_string_field(py, config, col_pattern, "error_title")?
                    .unwrap_or_default();
                validation
                    .set_error_title(&title)
                    .map_err(|e| format!("Failed to set error title: {}", e))?
                    .set_error_message(&msg)
                    .map_err(|e| format!("Failed to set error message: {}", e))?
                    .set_error_style(DataValidationErrorStyle::Stop)
            } else {
                validation
            };

            worksheet
                .add_data_validation(data_start_row, col_idx, data_end_row, col_idx, &validation)
                .map_err(|e| format!("Failed to add validation: {}", e))?;
        }
    }

    Ok(())
}
