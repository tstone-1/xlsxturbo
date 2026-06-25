//! Data validation application helpers.

use crate::parse::matches_pattern;
use crate::types::{extract_field, ValidationConfig};
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

fn reject_unknown_keys(
    config: &ValidationConfig,
    col_pattern: &str,
    allowed: &[&str],
) -> Result<(), String> {
    for key in config.keys() {
        if !allowed.contains(&key.as_str()) {
            return Err(format!(
                "validations['{}']: unknown option '{}'. Valid: {}",
                col_pattern,
                key,
                allowed.join(", ")
            ));
        }
    }
    Ok(())
}

fn validation_i32_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
    default: i32,
) -> Result<i32, String> {
    Ok(extract_field(
        py,
        config.get(key),
        &format!("validations['{}']", col_pattern),
        key,
        "an integer",
    )?
    .unwrap_or(default))
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
            const MESSAGE_KEYS: &[&str] = &[
                "type",
                "input_title",
                "input_message",
                "error_title",
                "error_message",
            ];

            let validation = match val_type.to_lowercase().as_str() {
                "list" => {
                    reject_unknown_keys(
                        config,
                        col_pattern,
                        &[
                            MESSAGE_KEYS[0],
                            "values",
                            MESSAGE_KEYS[1],
                            MESSAGE_KEYS[2],
                            MESSAGE_KEYS[3],
                            MESSAGE_KEYS[4],
                        ],
                    )?;
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
                    reject_unknown_keys(
                        config,
                        col_pattern,
                        &[
                            MESSAGE_KEYS[0],
                            "min",
                            "max",
                            MESSAGE_KEYS[1],
                            MESSAGE_KEYS[2],
                            MESSAGE_KEYS[3],
                            MESSAGE_KEYS[4],
                        ],
                    )?;
                    // Whole number validation with between rule
                    let min = validation_i32_field(py, config, col_pattern, "min", i32::MIN)?;
                    let max = validation_i32_field(py, config, col_pattern, "max", i32::MAX)?;
                    DataValidation::new()
                        .allow_whole_number(rust_xlsxwriter::DataValidationRule::Between(min, max))
                }
                "decimal" | "number" => {
                    reject_unknown_keys(
                        config,
                        col_pattern,
                        &[
                            MESSAGE_KEYS[0],
                            "min",
                            "max",
                            MESSAGE_KEYS[1],
                            MESSAGE_KEYS[2],
                            MESSAGE_KEYS[3],
                            MESSAGE_KEYS[4],
                        ],
                    )?;
                    // Decimal validation with between rule
                    let min = validation_f64_field(py, config, col_pattern, "min", f64::MIN)?;
                    let max = validation_f64_field(py, config, col_pattern, "max", f64::MAX)?;
                    DataValidation::new().allow_decimal_number(
                        rust_xlsxwriter::DataValidationRule::Between(min, max),
                    )
                }
                "text_length" | "textlength" | "length" => {
                    reject_unknown_keys(
                        config,
                        col_pattern,
                        &[
                            MESSAGE_KEYS[0],
                            "min",
                            "max",
                            MESSAGE_KEYS[1],
                            MESSAGE_KEYS[2],
                            MESSAGE_KEYS[3],
                            MESSAGE_KEYS[4],
                        ],
                    )?;
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
