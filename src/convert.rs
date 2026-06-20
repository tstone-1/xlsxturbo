//! Core conversion functions for CSV and DataFrame to XLSX

use crate::apply::{
    apply_cells, apply_charts, apply_checkboxes, apply_column_widths,
    apply_column_widths_with_autofit_cap, apply_comments, apply_conditional_formats,
    apply_formula_columns, apply_hyperlinks, apply_images, apply_merged_ranges, apply_rich_text,
    apply_textboxes, apply_validations,
};
use crate::parse::{
    build_column_formats, parse_header_format, parse_table_style, parse_value, sanitize_table_name,
};
use crate::types::{
    extract_columns, is_polars_dataframe, CellValue, DateOrder, EffectiveOpts, ExtractedOptions,
    WriteConfig,
};
use crate::workbook::apply_defined_names;
use crate::write::{write_cell, write_py_value_with_format, DATETIME_NUM_FORMAT, DATE_NUM_FORMAT};
use csv::ReaderBuilder;
use pyo3::prelude::*;
use rayon::prelude::*;
use rust_xlsxwriter::{Format, Table, Workbook, Worksheet};
use std::collections::HashMap;
use std::fs::File;

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
    // Open CSV file (csv::ReaderBuilder handles buffering internally)
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .buffer_capacity(1024 * 1024)
        .from_reader(file);

    // Create workbook and worksheet
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    let mut row_count: u32 = 0;
    let mut col_count: u16 = 0;

    // Process records
    for result in csv_reader.records() {
        let record = result.map_err(|e| format!("CSV parse error at row {}: {}", row_count, e))?;
        let num_cols = u16::try_from(record.len())
            .map_err(|_| format!("Column count {} exceeds u16 limit", record.len()))?;
        if num_cols > col_count {
            col_count = num_cols;
        }

        for (col_idx, value) in record.iter().enumerate() {
            let cell_value = parse_value(value, date_order);
            let col = u16::try_from(col_idx)
                .map_err(|_| format!("Column index {} exceeds u16 limit", col_idx))?;
            write_cell(
                worksheet,
                row_count,
                col,
                cell_value,
                &date_format,
                &datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_count, col_idx, e))?;
        }

        row_count = row_count
            .checked_add(1)
            .ok_or("Row count exceeds u32 limit")?;
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

/// Rows per chunk for parallel CSV processing. Picked so a parsed chunk's
/// peak memory stays bounded regardless of total file size.
const PARALLEL_CHUNK_ROWS: usize = 10_000;

/// Convert a CSV file to XLSX format using parallel processing.
///
/// Reads the CSV in chunks of `PARALLEL_CHUNK_ROWS` rows, parses each chunk in
/// parallel across rayon's thread pool, then writes the parsed chunk before
/// reading the next one. Peak memory is O(chunk) rather than O(file), so this
/// scales to CSVs larger than available RAM.
pub fn convert_csv_to_xlsx_parallel(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    date_order: DateOrder,
) -> Result<(u32, u16), String> {
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .buffer_capacity(1024 * 1024)
        .from_reader(file);

    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    let mut row_count: u32 = 0;
    let mut col_count: u16 = 0;
    let mut chunk: Vec<Vec<String>> = Vec::with_capacity(PARALLEL_CHUNK_ROWS);

    for result in csv_reader.records() {
        let absolute_row = row_count as usize + chunk.len();
        let record =
            result.map_err(|e| format!("CSV parse error at row {}: {}", absolute_row, e))?;
        let num_cols = u16::try_from(record.len())
            .map_err(|_| format!("Column count {} exceeds u16 limit", record.len()))?;
        if num_cols > col_count {
            col_count = num_cols;
        }
        chunk.push(record.iter().map(|s| s.to_string()).collect());

        if chunk.len() >= PARALLEL_CHUNK_ROWS {
            flush_parallel_chunk(
                worksheet,
                &mut chunk,
                &mut row_count,
                date_order,
                &date_format,
                &datetime_format,
            )?;
        }
    }

    if !chunk.is_empty() {
        flush_parallel_chunk(
            worksheet,
            &mut chunk,
            &mut row_count,
            date_order,
            &date_format,
            &datetime_format,
        )?;
    }

    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

/// Parse the chunk in parallel, write it sequentially, clear it, and advance `row_count`.
fn flush_parallel_chunk(
    worksheet: &mut Worksheet,
    chunk: &mut Vec<Vec<String>>,
    row_count: &mut u32,
    date_order: DateOrder,
    date_format: &Format,
    datetime_format: &Format,
) -> Result<(), String> {
    let parsed_rows: Vec<Vec<CellValue>> = chunk
        .par_iter()
        .map(|row| {
            row.iter()
                .map(|value| parse_value(value, date_order))
                .collect()
        })
        .collect();

    for (offset, parsed_row) in parsed_rows.into_iter().enumerate() {
        let row_u32 = row_count
            .checked_add(offset as u32)
            .ok_or("Row count exceeds u32 limit")?;
        for (col_idx, cell_value) in parsed_row.into_iter().enumerate() {
            let col_u16 = col_idx as u16; // safe: column count already validated via u16::try_from
            write_cell(
                worksheet,
                row_u32,
                col_u16,
                cell_value,
                date_format,
                datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_u32, col_idx, e))?;
        }
    }

    *row_count = row_count
        .checked_add(chunk.len() as u32)
        .ok_or("Row count exceeds u32 limit")?;
    chunk.clear();
    Ok(())
}

// ============================================================================
// DataFrame support
// ============================================================================

/// Write DataFrame data and apply all features to a worksheet.
///
/// This is the shared per-sheet write logic used by both `convert_dataframe_to_xlsx`
/// (single-sheet) and `dfs_to_xlsx` (multi-sheet). The caller is responsible for
/// creating the workbook/worksheet and saving.
pub(crate) fn write_sheet_data(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    df: &Bound<'_, PyAny>,
    config: &WriteConfig<'_>,
    opts: EffectiveOpts<'_>,
) -> Result<(u32, u16), String> {
    // Create formats
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    // Parse header format if provided
    let header_fmt = if let Some(fmt_dict) = opts.header_format {
        Some(parse_header_format(py, fmt_dict)?)
    } else {
        None
    };

    let mut row_idx: u32 = 0;

    // Get column names
    let is_polars = is_polars_dataframe(df)?;
    let columns: Vec<String> = extract_columns(df, is_polars)?;

    let col_count = u16::try_from(columns.len())
        .map_err(|_| format!("Column count {} exceeds u16 limit", columns.len()))?;

    // Build column formats if provided
    let col_formats: Vec<Option<Format>> = if let Some(cf) = opts.column_formats {
        build_column_formats(py, &columns, cf)?
    } else {
        vec![None; columns.len()]
    };

    // Track max content lengths for autofit+cap (only when both are active)
    let track_widths = config.autofit && opts.column_widths.is_some_and(|w| w.contains_key("_all"));
    let mut max_lens = vec![0usize; columns.len()];

    // Write header if requested
    if config.include_header {
        for (col_idx, col_name) in columns.iter().enumerate() {
            let col = col_idx as u16; // safe: col_count already validated via u16::try_from
            if track_widths {
                // Char count, not byte count: width is a visual estimate.
                max_lens[col_idx] = col_name.chars().count();
            }
            if let Some(ref fmt) = header_fmt {
                worksheet
                    .write_string_with_format(row_idx, col, col_name, fmt)
                    .map_err(|e| format!("Failed to write header '{}': {}", col_name, e))?;
            } else {
                worksheet
                    .write_string(row_idx, col, col_name)
                    .map_err(|e| format!("Failed to write header '{}': {}", col_name, e))?;
            }
        }
        row_idx = 1;
    }

    // Get row count
    let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
        let shape = df
            .getattr("shape")
            .map_err(|e| format!("Failed to get DataFrame shape: {}", e))?;
        let shape_tuple: (usize, usize) = shape
            .extract()
            .map_err(|e| format!("Failed to extract DataFrame shape: {}", e))?;
        shape_tuple.0
    } else {
        df.call_method0("__len__")
            .map_err(|e| format!("Failed to get DataFrame length: {}", e))?
            .extract()
            .map_err(|e| format!("Failed to extract DataFrame length: {}", e))?
    };

    if is_polars {
        // Polars: iterate using rows()
        let rows = df
            .call_method0("iter_rows")
            .map_err(|e| format!("Failed to iterate polars rows: {}", e))?;
        let iter = rows
            .try_iter()
            .map_err(|e| format!("Failed to create polars row iterator: {}", e))?;
        for row_result in iter {
            let row = row_result.map_err(|e| format!("Failed to read polars row: {}", e))?;
            let row_iter = row
                .try_iter()
                .map_err(|e| format!("Failed to iterate polars row values: {}", e))?;
            let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                .collect::<Result<Vec<_>, _>>()
                .map_err(|e| format!("Failed to collect polars row values: {}", e))?;

            for (col_idx, value) in row_tuple.iter().enumerate() {
                let col = col_idx as u16; // safe: col_count already validated via u16::try_from
                if track_widths {
                    // Char count, not byte count: width is a visual estimate.
                    let len = value
                        .str()
                        .map(|s| s.to_string_lossy().chars().count())
                        .unwrap_or(0);
                    if len > max_lens[col_idx] {
                        max_lens[col_idx] = len;
                    }
                }
                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col,
                    value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx = row_idx
                .checked_add(1)
                .ok_or("Row count exceeds u32 limit")?;
        }
    } else {
        // Pandas: use .values for faster access
        let values = df
            .getattr("values")
            .map_err(|e| format!("Failed to access DataFrame.values: {}", e))?;

        for i in 0..row_count {
            let row = values
                .get_item(i)
                .map_err(|e| format!("Failed to get row {}: {}", i, e))?;

            #[allow(clippy::needless_range_loop)]
            for col_idx in 0..columns.len() {
                let value = row
                    .get_item(col_idx)
                    .map_err(|e| format!("Failed to get value at ({}, {}): {}", i, col_idx, e))?;

                if track_widths {
                    // Char count, not byte count: width is a visual estimate.
                    let len = value
                        .str()
                        .map(|s| s.to_string_lossy().chars().count())
                        .unwrap_or(0);
                    if len > max_lens[col_idx] {
                        max_lens[col_idx] = len;
                    }
                }

                let col = col_idx as u16; // safe: col_count already validated via u16::try_from
                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col,
                    &value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx = row_idx
                .checked_add(1)
                .ok_or("Row count exceeds u32 limit")?;
        }
    }

    // Convert tracked content lengths to approximate Excel column widths
    let content_widths: Vec<f64> = if track_widths {
        max_lens
            .iter()
            .map(|&len| (len as f64 + 1.0).max(8.43))
            .collect()
    } else {
        Vec::new()
    };

    // Apply all worksheet features (table, formulas, formatting, etc.)
    let total_col_count = apply_worksheet_features(
        py,
        worksheet,
        &columns,
        col_count,
        row_idx,
        row_count,
        config,
        header_fmt.as_ref(),
        &opts,
        &content_widths,
    )?;

    Ok((row_idx, total_col_count))
}

pub(crate) fn write_configured_sheet(
    py: Python<'_>,
    workbook: &mut Workbook,
    df: &Bound<'_, PyAny>,
    sheet_name: &str,
    config: &WriteConfig<'_>,
    opts: EffectiveOpts<'_>,
) -> Result<(u32, u16), String> {
    let worksheet = if config.constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name '{}': {}", sheet_name, e))?;
    write_sheet_data(py, worksheet, df, config, opts)
}

/// Complex feature options that still work under `constant_memory` because they
/// are applied during the data-write phase (in `write_sheet_data`), not in
/// `apply_worksheet_features`. Every other present complex option is skipped.
const CONSTANT_MEMORY_SAFE_OPTIONS: &[&str] = &["column_widths", "header_format", "column_formats"];

/// Emit a `RuntimeWarning` listing the features that `constant_memory` mode
/// skips. The complex-feature list is derived from
/// `EffectiveOpts::present_complex_options` (single source of truth, generated
/// with the option fields), so adding a feature can never silently drop it from
/// this warning — only the handful of scalar flags are listed by hand.
fn warn_constant_memory_skips(
    py: Python<'_>,
    config: &WriteConfig<'_>,
    opts: &EffectiveOpts<'_>,
) -> Result<(), String> {
    let mut disabled: Vec<&str> = Vec::new();
    // Scalar config flags disabled by constant_memory.
    if config.table_style.is_some() {
        disabled.push("table_style");
    }
    if config.freeze_panes {
        disabled.push("freeze_panes");
    }
    if config.autofit {
        disabled.push("autofit");
    }
    if config.row_heights.is_some() {
        disabled.push("row_heights");
    }
    // Complex feature options: every present one is skipped except those applied
    // during the write phase. New features default to "skipped + warned", the
    // safe direction.
    for name in opts.present_complex_options() {
        if !CONSTANT_MEMORY_SAFE_OPTIONS.contains(&name) {
            disabled.push(name);
        }
    }
    if disabled.is_empty() {
        return Ok(());
    }
    let warnings = py
        .import("warnings")
        .map_err(|e| format!("Failed to import warnings: {}", e))?;
    let msg = format!(
        "constant_memory=True disables these features: {}",
        disabled.join(", ")
    );
    let runtime_warning = py
        .import("builtins")
        .and_then(|b| b.getattr("RuntimeWarning"))
        .map_err(|e| format!("Failed to get RuntimeWarning: {}", e))?;
    warnings
        .call_method1("warn", (msg, runtime_warning))
        .map_err(|e| format!("Failed to emit warning: {}", e))?;
    Ok(())
}

/// Apply all worksheet features after data has been written.
///
/// Handles: table formatting, formula columns, conditional formats, freeze panes,
/// column widths/autofit, row heights, merged ranges, hyperlinks, comments,
/// validations, rich text, and images. All features except column widths are
/// skipped in constant_memory mode.
#[allow(clippy::too_many_arguments)]
fn apply_worksheet_features(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    col_count: u16,
    last_row_idx: u32,
    row_count: usize,
    config: &WriteConfig<'_>,
    header_fmt: Option<&Format>,
    opts: &EffectiveOpts<'_>,
    content_widths: &[f64],
) -> Result<u16, String> {
    // In constant_memory mode, only column widths (without autofit) are supported.
    // Warn about every other requested feature right here, next to the skip.
    if config.constant_memory {
        warn_constant_memory_skips(py, config, opts)?;
        if let Some(widths) = opts.column_widths {
            apply_column_widths(worksheet, col_count, widths)?;
        }
        return Ok(col_count);
    }

    // Add Excel Table if requested (requires header + at least one data row)
    if let Some(style_name) = config.table_style {
        if row_count > 0 && config.include_header {
            let style = parse_table_style(style_name)?;
            let mut table = Table::new().set_style(style);

            if let Some(name) = config.table_name {
                let sanitized = sanitize_table_name(name);
                table = table.set_name(&sanitized);
            }

            let last_row = last_row_idx.saturating_sub(1);
            let last_col = col_count.saturating_sub(1);

            worksheet
                .add_table(0, 0, last_row, last_col, &table)
                .map_err(|e| format!("Failed to add table: {}", e))?;
        }
    }

    let data_row_start = if config.include_header { 1u32 } else { 0u32 };
    let data_row_end = last_row_idx.saturating_sub(1);
    let has_data_rows = row_count > 0 && data_row_end >= data_row_start;

    // Apply formula columns
    let mut total_col_count = col_count;
    if let Some(formulas) = opts.formula_columns {
        if !formulas.is_empty() && has_data_rows {
            let formula_cols_added = apply_formula_columns(
                worksheet,
                formulas,
                col_count,
                data_row_start,
                data_row_end,
                config.include_header,
                header_fmt,
            )?;
            total_col_count = col_count
                .checked_add(formula_cols_added)
                .ok_or("Total column count exceeds u16 limit")?;
        }
    }

    // Apply conditional formats
    if let Some(cond_fmts) = opts.conditional_formats {
        if has_data_rows {
            apply_conditional_formats(
                py,
                worksheet,
                columns,
                data_row_start,
                data_row_end,
                cond_fmts,
            )?;
        }
    }

    // Freeze panes (freeze header row)
    if config.freeze_panes && config.include_header {
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|e| format!("Failed to freeze panes: {}", e))?;
    }

    // Apply custom column widths and/or autofit
    if let Some(widths) = opts.column_widths {
        if config.autofit && widths.contains_key("_all") {
            apply_column_widths_with_autofit_cap(worksheet, col_count, widths, content_widths)?;
        } else {
            apply_column_widths(worksheet, col_count, widths)?;
        }
    } else if config.autofit {
        worksheet.autofit();
    }

    // Apply custom row heights
    if let Some(heights) = config.row_heights {
        for (&row_idx_h, &height) in heights.iter() {
            worksheet
                .set_row_height(row_idx_h, height)
                .map_err(|e| format!("Failed to set row height: {}", e))?;
        }
    }

    // Apply merged ranges
    if let Some(ranges) = opts.merged_ranges {
        if !ranges.is_empty() {
            apply_merged_ranges(py, worksheet, ranges)?;
        }
    }

    // Apply hyperlinks
    if let Some(links) = opts.hyperlinks {
        if !links.is_empty() {
            apply_hyperlinks(worksheet, links)?;
        }
    }

    // Apply comments/notes
    if let Some(cmts) = opts.comments {
        if !cmts.is_empty() {
            apply_comments(worksheet, cmts)?;
        }
    }

    // Apply data validations
    if let Some(vals) = opts.validations {
        if has_data_rows {
            apply_validations(py, worksheet, columns, data_row_start, data_row_end, vals)?;
        }
    }

    // Apply rich text
    if let Some(rt) = opts.rich_text {
        if !rt.is_empty() {
            apply_rich_text(py, worksheet, rt)?;
        }
    }

    // Apply images
    if let Some(imgs) = opts.images {
        if !imgs.is_empty() {
            apply_images(py, worksheet, imgs)?;
        }
    }

    // Apply checkboxes
    if let Some(cbxs) = opts.checkboxes {
        if !cbxs.is_empty() {
            apply_checkboxes(py, worksheet, cbxs)?;
        }
    }

    // Apply textboxes
    if let Some(tbxs) = opts.textboxes {
        if !tbxs.is_empty() {
            apply_textboxes(py, worksheet, tbxs)?;
        }
    }

    // Apply native Excel charts
    if let Some(charts) = opts.charts {
        if !charts.is_empty() {
            apply_charts(py, worksheet, charts)?;
        }
    }

    // Apply cells (arbitrary cell writes, after all DataFrame data)
    if let Some(cells) = opts.cells {
        if !cells.is_empty() {
            apply_cells(py, worksheet, cells)?;
        }
    }

    Ok(total_col_count)
}

/// Convert a DataFrame (pandas or polars) to XLSX format
#[allow(clippy::too_many_arguments)]
pub(crate) fn convert_dataframe_to_xlsx(
    py: Python<'_>,
    df: &Bound<'_, PyAny>,
    output_path: &str,
    sheet_name: &str,
    include_header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    table_name: Option<&str>,
    row_heights: Option<&HashMap<u32, f64>>,
    constant_memory: bool,
    opts: &ExtractedOptions,
    defined_names: Option<&HashMap<String, String>>,
) -> Result<(u32, u16), String> {
    let mut workbook = rust_xlsxwriter::Workbook::new();

    let config = WriteConfig {
        include_header,
        autofit,
        table_style,
        freeze_panes,
        table_name,
        row_heights,
        constant_memory,
    };

    let result = write_configured_sheet(
        py,
        &mut workbook,
        df,
        sheet_name,
        &config,
        opts.as_effective(),
    )?;

    apply_defined_names(&mut workbook, defined_names)?;

    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok(result)
}

#[cfg(test)]
mod constant_memory_tests {
    use super::CONSTANT_MEMORY_SAFE_OPTIONS;
    use crate::types::EffectiveOpts;

    /// Guard the one remaining hand-maintained part of the `constant_memory`
    /// skip warning: the safe/skipped classification of each complex option.
    ///
    /// `warn_constant_memory_skips` warns every present complex option except
    /// those in `CONSTANT_MEMORY_SAFE_OPTIONS`. Adding a field to the
    /// `define_options!` list auto-grows `EffectiveOpts::COMPLEX_OPTION_NAMES`,
    /// which makes this test fail until the new option is placed in exactly one
    /// of the two sets — forcing a deliberate "does this work under
    /// constant_memory?" decision instead of defaulting silently.
    #[test]
    fn every_complex_option_is_classified_for_constant_memory() {
        // The complex options that `apply_worksheet_features` skips under
        // constant_memory (everything applied only after the data write).
        const EXPECTED_SKIPPED: &[&str] = &[
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
            "cells",
        ];

        for &name in EffectiveOpts::COMPLEX_OPTION_NAMES {
            let safe = CONSTANT_MEMORY_SAFE_OPTIONS.contains(&name);
            let skipped = EXPECTED_SKIPPED.contains(&name);
            assert!(
                safe ^ skipped,
                "complex option '{}' must be classified exactly once as \
                 constant_memory-safe OR skipped — a newly added option needs a \
                 deliberate decision in CONSTANT_MEMORY_SAFE_OPTIONS (applied during \
                 the data write) or EXPECTED_SKIPPED (skipped + warned)",
                name
            );
        }

        // No stray/duplicate names in either set, and full coverage.
        assert_eq!(
            EffectiveOpts::COMPLEX_OPTION_NAMES.len(),
            CONSTANT_MEMORY_SAFE_OPTIONS.len() + EXPECTED_SKIPPED.len(),
            "every complex option must be accounted for exactly once across the \
             safe and skipped sets"
        );
    }
}
