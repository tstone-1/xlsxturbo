//! Shared workbook-level helpers.

use rust_xlsxwriter::Workbook;
use std::collections::HashMap;

pub(crate) fn apply_defined_names(
    workbook: &mut Workbook,
    defined_names: Option<&HashMap<String, String>>,
) -> Result<(), String> {
    if let Some(names) = defined_names {
        for (name, reference) in names {
            // The local part (after a sheet-qualifying '!') must be non-empty:
            // rust_xlsxwriter's define_name calls `chars().next().unwrap()` and
            // would panic on an empty name (e.g. "" or "Sheet1!").
            let local = name.rsplit('!').next().unwrap_or("");
            if local.is_empty() {
                return Err(format!(
                    "Invalid defined name '{}': name must not be empty",
                    name
                ));
            }
            workbook
                .define_name(name, reference)
                .map_err(|e| format!("Failed to define name '{}': {}", name, e))?;
        }
    }
    Ok(())
}
