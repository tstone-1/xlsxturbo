//! xlsxturbo - High-performance Excel writer with automatic type detection
//!
//! This library provides fast DataFrame and CSV to Excel conversion:
//! - Integers and floats → Excel numbers
//! - Booleans (true/false) → Excel booleans
//! - Dates → Excel dates
//! - Datetimes → Excel datetimes
//! - NaN/Inf/None → Empty cells
//! - Everything else → Strings
//!
//! Supports pandas DataFrames, polars DataFrames, and CSV files.

mod apply;
mod convert;
mod errors;
mod extract;
mod parse;
mod types;
mod workbook;
mod write;

// Re-export public API for the CLI binary (main.rs)
pub use convert::{convert_csv_to_xlsx, convert_csv_to_xlsx_parallel, ConvertError};
pub use types::DateOrder;
pub use workbook::Destination;

use convert::{
    add_named_worksheet, apply_cell_options, claimed_table_name, convert_dataframe_to_xlsx,
    dataframe_row_count, finish_sheet, finish_workbook, write_frame, FrameOutcome, WorkbookOptions,
};
use extract::{
    extract_cells, extract_charts, extract_checkboxes, extract_column_formats,
    extract_column_widths, extract_comments, extract_conditional_formats, extract_formula_columns,
    extract_header_format, extract_hyperlinks, extract_images, extract_merged_ranges,
    extract_rich_text, extract_row_heights, extract_sheet_info, extract_sparklines,
    extract_textboxes, extract_validations,
};
use types::pytype_name;
use types::ExtractedOptions;
use types::WriteConfig;
use types::{EffectiveOpts, FreezePanes, SheetConfig};
use workbook::{
    reject_frames_out_of_order, reject_overlapping_frames, reject_table_name_collisions,
    ClaimedTableName, FrameSpan,
};

use pyo3::prelude::*;
use pyo3::types::PyBytes;
use rust_xlsxwriter::Workbook;
use std::collections::HashMap;

fn path_arg_to_string(value: &Bound<'_, PyAny>, param_name: &str) -> PyResult<String> {
    if let Ok(path) = value.extract::<String>() {
        return Ok(path);
    }
    if let Ok(pathlike) = value.call_method0("__fspath__") {
        if let Ok(path) = pathlike.extract::<String>() {
            return Ok(path);
        }
    }
    Err(errors::configuration_type(format!(
        "'{}' must be str or a path-like object returning str (bytes paths are not supported), got {}",
        param_name,
        pytype_name(value)
    )))
}

/// What `output_path` named: a file to write, or an object to write into.
enum OutputArg<'py> {
    /// A `str` or `os.PathLike` path.
    Path(String),
    /// A binary file-like object -- `io.BytesIO`, a file opened `"wb"`, a web
    /// framework's response body. Anything with a `write(bytes)` method.
    Writer(Bound<'py, PyAny>),
}

impl<'py> OutputArg<'py> {
    /// Classify `output_path`.
    ///
    /// A path is tried first, so a `str` subclass or a path-like object that
    /// also happens to have a `write` method is still a path, as it always was.
    /// A text-mode stream is refused here rather than at the end of the export,
    /// where its `write()` would reject the bytes with a message naming neither
    /// this library nor the argument.
    fn from_py(value: &Bound<'py, PyAny>) -> PyResult<Self> {
        if let Ok(path) = value.extract::<String>() {
            return Ok(Self::Path(path));
        }
        if let Ok(pathlike) = value.call_method0("__fspath__") {
            if let Ok(path) = pathlike.extract::<String>() {
                return Ok(Self::Path(path));
            }
        }
        if value.hasattr("write")? {
            let text_io = value.py().import("io")?.getattr("TextIOBase")?;
            if value.is_instance(&text_io)? {
                return Err(errors::configuration_type(format!(
                    "'output_path' is a text stream ({}); an .xlsx file is binary, so pass a \
                     binary stream such as io.BytesIO() or a file opened with 'wb'",
                    pytype_name(value)
                )));
            }
            return Ok(Self::Writer(value.clone()));
        }
        Err(errors::configuration_type(format!(
            "'output_path' must be str, a path-like object returning str (bytes paths are not \
             supported), or a binary file-like object with a write() method, got {}",
            pytype_name(value)
        )))
    }

    /// The destination the save writes to. For a writer, the archive lands in
    /// `buffer` and [`OutputArg::deliver`] passes it on.
    fn destination<'a>(&'a self, buffer: &'a mut Vec<u8>) -> Destination<'a> {
        match self {
            Self::Path(path) => Destination::Path(path),
            Self::Writer(_) => Destination::Memory(buffer),
        }
    }

    /// Hand a finished in-memory archive to the writer. Nothing to do for a path.
    ///
    /// The object is written to and left open at its new position: it belongs to
    /// the caller, who may be about to `seek(0)` or `getvalue()` it. A `write()`
    /// that returns a count shorter than it was given -- allowed for raw,
    /// unbuffered streams -- is called again with the rest. One that returns
    /// `None` means would-block for RawIOBase, but acceptance for Django's
    /// `HttpResponse.write` and similar response bodies. An exception raised
    /// by `write()` propagates unchanged.
    fn deliver(&self, buffer: &[u8]) -> PyResult<()> {
        let Self::Writer(writer) = self else {
            return Ok(());
        };
        let mut offset = 0;
        while offset < buffer.len() {
            let chunk = PyBytes::new(writer.py(), &buffer[offset..]);
            let returned = writer.call_method1("write", (chunk,))?;
            let written = if returned.is_none() {
                let raw_io = writer.py().import("io")?.getattr("RawIOBase")?;
                if writer.is_instance(&raw_io)? {
                    return Err(errors::file(errors::FileFailure {
                        message: format!(
                            "Failed to write workbook to {}: write() would block with {} bytes remaining",
                            pytype_name(writer),
                            buffer.len() - offset
                        ),
                        errno: Some(writer.py().import("errno")?.getattr("EAGAIN")?.extract()?),
                    }));
                }
                buffer.len() - offset
            } else {
                returned.extract::<usize>()?
            };
            if written == 0 {
                return Err(errors::file(errors::FileFailure {
                    message: format!(
                        "Failed to write workbook to {}: write() accepted 0 of the remaining {} bytes",
                        pytype_name(writer),
                        buffer.len() - offset
                    ),
                    errno: None,
                }));
            }
            offset += written.min(buffer.len() - offset);
        }
        Ok(())
    }
}

/// Helper: take a dict, or any other mapping as a dict (see
/// `types::mapping_as_dict`), or raise TypeError with a clear message.
fn require_dict<'py>(
    value: &Bound<'py, PyAny>,
    param_name: &str,
) -> PyResult<Bound<'py, pyo3::types::PyDict>> {
    types::mapping_as_dict(value)?.ok_or_else(|| {
        errors::configuration_type(format!(
            "expected dict for '{}', got {}",
            param_name,
            pytype_name(value)
        ))
    })
}

/// Extract the `defined_names` kwarg from any mapping.
fn extract_defined_names_kwarg(
    value: Option<&Bound<'_, PyAny>>,
) -> PyResult<Option<HashMap<String, String>>> {
    match value {
        Some(v) if !v.is_none() => require_dict(v, "defined_names")
            .and_then(|d| extract::extract_defined_names(&d))
            .map(Some),
        _ => Ok(None),
    }
}

/// Extract the `freeze_panes` keyword argument; `None` means off.
fn extract_freeze_panes_kwarg(value: Option<&Bound<'_, PyAny>>) -> PyResult<types::FreezePanes> {
    match value {
        Some(v) if !v.is_none() => extract::extract_freeze_panes(v, "freeze_panes"),
        _ => Ok(types::FreezePanes::Off),
    }
}

/// Extract the `compression_level` keyword argument; `None` keeps the
/// writer's default, level 6.
fn extract_compression_level_kwarg(value: Option<&Bound<'_, PyAny>>) -> PyResult<Option<u8>> {
    match value {
        Some(v) if !v.is_none() => extract::extract_compression_level(v).map(Some),
        _ => Ok(None),
    }
}

/// Extract the `tab_color` keyword argument; `None` means no colour.
fn extract_tab_color_kwarg(value: Option<&Bound<'_, PyAny>>) -> PyResult<Option<u32>> {
    match value {
        Some(v) if !v.is_none() => extract::extract_tab_color(v, "tab_color").map(Some),
        _ => Ok(None),
    }
}

/// Extract the `row_heights` kwarg, which both entry points take as a raw
/// object rather than through `ExtractedOptions`.
///
/// It is not part of `RawOptions` because it is not a per-sheet *complex*
/// option in the `define_options!` sense -- it travels in `WriteConfig`
/// alongside the scalars. What it does share with the complex options, since
/// this function exists, is that its keys and values are validated by
/// xlsxturbo rather than by PyO3's `HashMap<u32, f64>` conversion; see
/// `extract_row_heights` for what that conversion used to let through.
fn extract_row_heights_kwarg(
    value: Option<&Bound<'_, PyAny>>,
) -> PyResult<Option<HashMap<u32, f64>>> {
    value
        .map(|v| require_dict(v, "row_heights").and_then(|d| extract_row_heights(&d)))
        .transpose()
}

/// Extract the `startrow` keyword argument; `None` means row 0.
fn extract_startrow_kwarg(value: Option<&Bound<'_, PyAny>>) -> PyResult<u32> {
    match value {
        Some(v) if !v.is_none() => extract::extract_startrow(v, "startrow"),
        _ => Ok(0),
    }
}

/// Helper: take a list, or any other sequence as a list (see
/// `types::sequence_as_list`), or raise TypeError with a clear message.
fn require_list<'py>(
    value: &Bound<'py, PyAny>,
    param_name: &str,
) -> PyResult<Bound<'py, pyo3::types::PyList>> {
    types::sequence_as_list(value)?.ok_or_else(|| {
        errors::configuration_type(format!(
            "expected list for '{}', got {}",
            param_name,
            pytype_name(value)
        ))
    })
}

/// Reject anything that is not a supported DataFrame, at the boundary.
///
/// This runs the *same* predicate the write pipeline runs, and exists for two reasons
/// that the inner call cannot serve. First, classification: inside the pipeline the
/// failure is a `String` that reaches `lib.rs` indistinguishable from a bad option, so
/// `InputDataError` would be unraisable -- verified, before this existed
/// `df_to_xlsx([1, 2, 3], ...)` raised `ConfigurationError`. Second, ordering: an
/// unsupported input now fails before any file is created.
///
/// Calling `types::is_polars_dataframe` rather than re-implementing the check is what
/// keeps the two sites from drifting apart.
fn require_supported_dataframe(df: &Bound<'_, PyAny>, sheet: Option<&str>) -> PyResult<()> {
    types::is_polars_dataframe(df).map(|_| ()).map_err(|e| {
        errors::input_data(match sheet {
            Some(name) => format!("sheet '{}': {}", name, e),
            None => e,
        })
    })
}

struct RawOptions<'a, 'py> {
    column_widths: Option<&'a Bound<'py, PyAny>>,
    header_format: Option<&'a Bound<'py, PyAny>>,
    column_formats: Option<&'a Bound<'py, PyAny>>,
    conditional_formats: Option<&'a Bound<'py, PyAny>>,
    formula_columns: Option<&'a Bound<'py, PyAny>>,
    merged_ranges: Option<&'a Bound<'py, PyAny>>,
    hyperlinks: Option<&'a Bound<'py, PyAny>>,
    comments: Option<&'a Bound<'py, PyAny>>,
    validations: Option<&'a Bound<'py, PyAny>>,
    rich_text: Option<&'a Bound<'py, PyAny>>,
    images: Option<&'a Bound<'py, PyAny>>,
    checkboxes: Option<&'a Bound<'py, PyAny>>,
    textboxes: Option<&'a Bound<'py, PyAny>>,
    charts: Option<&'a Bound<'py, PyAny>>,
    sparklines: Option<&'a Bound<'py, PyAny>>,
    cells: Option<&'a Bound<'py, PyAny>>,
}

/// Extract and validate all optional write parameters from Python into typed Rust structs.
fn extract_options(raw: &RawOptions<'_, '_>) -> PyResult<ExtractedOptions> {
    Ok(ExtractedOptions {
        column_widths: raw
            .column_widths
            .map(|v| require_dict(v, "column_widths").and_then(|d| extract_column_widths(&d)))
            .transpose()?,
        header_format: raw
            .header_format
            .map(|v| require_dict(v, "header_format").and_then(|d| extract_header_format(&d)))
            .transpose()?,
        column_formats: raw
            .column_formats
            .map(|v| require_dict(v, "column_formats").and_then(|d| extract_column_formats(&d)))
            .transpose()?,
        conditional_formats: raw
            .conditional_formats
            .map(|v| {
                require_dict(v, "conditional_formats").and_then(|d| extract_conditional_formats(&d))
            })
            .transpose()?,
        formula_columns: raw
            .formula_columns
            .map(|v| require_dict(v, "formula_columns").and_then(|d| extract_formula_columns(&d)))
            .transpose()?,
        merged_ranges: raw
            .merged_ranges
            .map(|v| require_list(v, "merged_ranges").and_then(|l| extract_merged_ranges(&l)))
            .transpose()?,
        hyperlinks: raw
            .hyperlinks
            .map(|v| require_list(v, "hyperlinks").and_then(|l| extract_hyperlinks(&l)))
            .transpose()?,
        comments: raw
            .comments
            .map(|v| require_dict(v, "comments").and_then(|d| extract_comments(&d)))
            .transpose()?,
        validations: raw
            .validations
            .map(|v| require_dict(v, "validations").and_then(|d| extract_validations(&d)))
            .transpose()?,
        rich_text: raw
            .rich_text
            .map(|v| require_dict(v, "rich_text").and_then(|d| extract_rich_text(&d)))
            .transpose()?,
        images: raw
            .images
            .map(|v| require_dict(v, "images").and_then(|d| extract_images(&d)))
            .transpose()?,
        checkboxes: raw
            .checkboxes
            .map(|v| require_dict(v, "checkboxes").and_then(|d| extract_checkboxes(&d)))
            .transpose()?,
        textboxes: raw
            .textboxes
            .map(|v| require_dict(v, "textboxes").and_then(|d| extract_textboxes(&d)))
            .transpose()?,
        charts: raw
            .charts
            .map(|v| require_dict(v, "charts").and_then(|d| extract_charts(&d)))
            .transpose()?,
        sparklines: raw
            .sparklines
            .map(|v| require_dict(v, "sparklines").and_then(|d| extract_sparklines(&d)))
            .transpose()?,
        cells: raw
            .cells
            .map(|v| require_dict(v, "cells").and_then(|d| extract_cells(&d)))
            .transpose()?,
    })
}

/// One `dfs_to_xlsx` entry with its options resolved against the workbook-wide
/// keywords, held until every entry has been checked.
///
/// Several entries may name one sheet. `first` is the entry that created the
/// sheet, which is the one that decides the options describing the whole sheet.
struct SheetEntry<'py> {
    df: Bound<'py, PyAny>,
    /// The name as this entry spells it. The sheet carries the first entry's
    /// spelling; Excel compares sheet names without regard to case.
    name: String,
    config: SheetConfig,
    header: bool,
    autofit: bool,
    table_style: Option<String>,
    freeze_panes: FreezePanes,
    autofilter: bool,
    tab_color: Option<u32>,
    startrow: u32,
    table_name: Option<String>,
    /// Index of the first entry for the same sheet; its own index for that entry.
    first: usize,
}

impl SheetEntry<'_> {
    /// The scalar write configuration of this entry.
    ///
    /// The fields that describe the whole sheet (`autofit`, `freeze_panes`,
    /// `tab_color`, `row_heights`) are read by `finish_sheet`, which is called
    /// with the first entry's configuration only. On a later entry they are the
    /// workbook-wide defaults and nothing applies them.
    fn write_config<'a>(
        &'a self,
        row_heights: Option<&'a HashMap<u32, f64>>,
        constant_memory: bool,
    ) -> WriteConfig<'a> {
        WriteConfig {
            include_header: self.header,
            autofit: self.autofit,
            table_style: self.table_style.as_deref(),
            freeze_panes: self.freeze_panes,
            autofilter: self.autofilter,
            tab_color: self.tab_color,
            startrow: self.startrow,
            table_name: self.table_name.as_deref(),
            row_heights,
            constant_memory,
        }
    }
}

/// Refuse an option that describes the whole sheet on a second or later entry
/// for that sheet.
///
/// The first entry decides those. Taking the last one instead would be silent
/// and would depend on the order of the list; refusing by name is neither.
fn reject_whole_sheet_options(
    config: &SheetConfig,
    index: usize,
    sheet_name: &str,
    first: usize,
) -> PyResult<()> {
    let set = [
        ("autofit", config.autofit.is_some()),
        ("freeze_panes", config.freeze_panes.is_some()),
        ("tab_color", config.tab_color.is_some()),
        ("column_widths", config.column_widths.is_some()),
        ("row_heights", config.row_heights.is_some()),
    ];
    match set.iter().find(|(_, present)| *present) {
        Some((option, _)) => Err(errors::configuration(format!(
            "sheets[{}] ('{}'): '{}' describes the whole sheet, and sheets[{}] is the first \
             entry for this sheet. Set it there.",
            index, sheet_name, option, first
        ))),
        None => Ok(()),
    }
}

/// Check the layout of every sheet that holds more than one frame, before a
/// cell is written: the frames may not overlap, at most one may carry the
/// plain-range autofilter, and under `constant_memory` they are listed from top
/// to bottom.
fn check_frame_layout(entries: &[SheetEntry<'_>], constant_memory: bool) -> PyResult<()> {
    for (first, lead) in entries.iter().enumerate() {
        if lead.first != first {
            continue;
        }
        let members: Vec<usize> = (first..entries.len())
            .filter(|&index| entries[index].first == first)
            .collect();
        if members.len() < 2 {
            continue;
        }

        let filtered: Vec<usize> = members
            .iter()
            .copied()
            .filter(|&index| entries[index].autofilter)
            .collect();
        if let [one, two, ..] = filtered[..] {
            return Err(errors::configuration(format!(
                "sheet '{}': sheets[{}] and sheets[{}] both have autofilter=True, and a sheet \
                 holds one filter. Set 'autofilter': False on one of them, or give each frame \
                 a table_style: a table has its own filter.",
                lead.name, one, two
            )));
        }

        let mut spans = Vec::with_capacity(members.len());
        for &index in &members {
            let entry = &entries[index];
            let rows = dataframe_row_count(&entry.df)
                .map_err(|e| errors::input_data(format!("sheet '{}': {}", entry.name, e)))?;
            let start = u64::from(entry.startrow);
            spans.push(FrameSpan {
                entry: index,
                start,
                end: start + rows as u64 + u64::from(entry.header),
            });
        }
        reject_overlapping_frames(&lead.name, &spans).map_err(errors::frame_overlap)?;
        if constant_memory {
            reject_frames_out_of_order(&lead.name, &spans).map_err(errors::configuration)?;
        }
    }
    Ok(())
}

/// The column extent of a sheet over every frame written to it: the widest
/// frame, and per column the widest estimated content.
fn sheet_extent(frames: &[&FrameOutcome]) -> (u16, Vec<f64>) {
    let col_count = frames.iter().map(|f| f.col_count).max().unwrap_or(0);
    let mut widths: Vec<f64> = Vec::new();
    for frame in frames {
        for (column, &width) in frame.content_widths.iter().enumerate() {
            match widths.get_mut(column) {
                Some(widest) => *widest = widest.max(width),
                None => widths.push(width),
            }
        }
    }
    (col_count, widths)
}

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// Reads a CSV file and converts it to an Excel XLSX file, automatically
/// detecting data types:
/// - Integers and floats become Excel numbers
/// - "true"/"false" (case-insensitive) become Excel booleans
/// - Dates (YYYY-MM-DD, DD-MM-YYYY, MM-DD-YYYY) become Excel dates
/// - Datetimes (ISO 8601) become Excel datetimes
/// - NaN/Inf values become empty cells
/// - Everything else becomes text
///
/// Args:
///     input_path: Path to the input CSV file
///     output_path: Path for the output XLSX file, or a binary file-like
///         object (io.BytesIO, a file opened 'wb') to write the workbook into
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     parallel: Use multi-core parallel processing (default: False).
///               Faster for large files (100K+ rows) but uses more memory.
///     date_order: Date parsing order for ambiguous dates like "01-02-2024" (default: "auto").
///                 "auto" - ISO first, then European (DMY), then US (MDY)
///                 "mdy" or "us" - US format: 01-02-2024 = January 2nd
///                 "dmy" or "eu" - European format: 01-02-2024 = February 1st
///     compression_level: Zip compression level of the file, 1 (fastest, largest)
///                        to 9 (default: None, which is level 6).
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     XlsxTurboError: If the conversion fails. The subclass says what kind of
///         failure it was -- ConfigurationError / ConfigurationTypeError for an
///         option, InputDataError for the frame, FileError for the filesystem.
///         Every one of them is also the builtin its failures raised before
///         0.19.0, so `except ValueError` still works; note
///         ConfigurationTypeError is a TypeError rather than a ValueError.
///         Full contract: https://tstone-1.github.io/xlsxturbo/errors/
///
/// Example:
///     >>> import xlsxturbo
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "output.xlsx")
///     >>> # For US date format (MM-DD-YYYY):
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "out.xlsx", date_order="us")
///     >>> # For large files, use parallel processing:
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("big.csv", "out.xlsx", parallel=True)
#[pyfunction]
#[pyo3(signature = (input_path, output_path, sheet_name = "Sheet1", parallel = false, date_order = "auto", compression_level = None))]
fn csv_to_xlsx(
    py: Python<'_>,
    input_path: &Bound<'_, PyAny>,
    output_path: &Bound<'_, PyAny>,
    sheet_name: &str,
    parallel: bool,
    date_order: &str,
    compression_level: Option<&Bound<'_, PyAny>>,
) -> PyResult<(u32, u16)> {
    let input_path = path_arg_to_string(input_path, "input_path")?;
    let output = OutputArg::from_py(output_path)?;
    let sheet_name = sheet_name.to_string();
    let order = DateOrder::parse(date_order).ok_or_else(|| {
        errors::configuration(format!(
            "Invalid date_order '{}'. Valid values: auto, mdy, us, dmy, eu, european",
            date_order
        ))
    })?;
    let compression_level = extract_compression_level_kwarg(compression_level)?;

    // No Python objects are touched below this point, so release the GIL for
    // the (potentially rayon-parallel) pure-Rust conversion work.
    let mut buffer = Vec::new();
    let destination = output.destination(&mut buffer);
    let result = py.detach(|| {
        if parallel {
            convert_csv_to_xlsx_parallel(
                &input_path,
                destination,
                &sheet_name,
                order,
                compression_level,
            )
        } else {
            convert_csv_to_xlsx(
                &input_path,
                destination,
                &sheet_name,
                order,
                compression_level,
            )
        }
    })?;
    output.deliver(&buffer)?;
    Ok(result)
}

/// Convert a pandas or polars DataFrame to XLSX format.
///
/// This function writes a DataFrame directly to an Excel XLSX file,
/// preserving data types without intermediate CSV conversion.
///
/// Args:
///     df: pandas DataFrame or polars DataFrame to export
///     output_path: Path for the output XLSX file, or a binary file-like
///         object (io.BytesIO, a file opened 'wb') to write the workbook into
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///              Combined with column_widths: explicit widths win for the columns
///              they name; every other column is still autofitted (rather than
///              left at Excel's default width). Add an "_all" entry in
///              column_widths to cap the autofit width instead of overriding it.
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: True freezes the header row (default: False). A (row, col) tuple
///                   freezes at that zero-based cell instead, as XlsxWriter's
///                   freeze_panes(row, col) does: it is the top-left cell of the scrolling
///                   area, so (4, 1) keeps four rows and one column in view. True follows
///                   startrow and does nothing with header=False; a tuple is taken as given.
///     column_widths: Dict mapping column index (0-based) or "_all" to width in characters
///                    (default: None). Example: {0: 20, 1: 15, 3: 30} sets widths for columns
///                    A, B, and D. An integer key must be a non-negative index within Excel's
///                    column range (0..=16383); a negative key, a key beyond 16383, or a
///                    non-integer/non-"_all" key raises. With autofit=True and no "_all" key:
///                    listed columns get the explicit width, unlisted columns are autofitted.
///                    With autofit=True and an "_all" key: "_all" caps the autofit width for
///                    unlisted columns instead of overriding it.
///     table_name: Custom name for the Excel table (requires table_style; default: auto-generated,
///                 `Table1`, `Table2`, ... in sheet order).
///                 Must be alphanumeric/underscore, max 255 chars; anything else is
///                 rewritten. Effective names must be unique across the workbook after
///                 that rewrite, and must not collide with a defined_names entry.
///     header_format: Dict with header cell formatting options (default: None)
///                    Example: {"bold": True, "bg_color": "#4F81BD", "font_color": "white"}
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///                  Example: {0: 20, 5: 30} sets heights for specific rows
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Emits a RuntimeWarning and disables: table_style,
///                      row_heights, autofit, column_widths with autofit cap, conditional_formats,
///                      formula_columns, merged_ranges, hyperlinks, comments, validations,
///                      rich_text, images, checkboxes, textboxes, charts, sparklines, and cells.
///                      Plain column_widths, header_format, column_formats, freeze_panes,
///                      autofilter and tab_color remain supported.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline, border.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names/patterns to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     formula_columns: Dict mapping column names to Excel formula templates (default: None).
///                      Use {row} as placeholder for the current row number.
///                      Example: {"Total": "=SUM(A{row}:C{row})"}
///     merged_ranges: List of merge specs: (range, text) or (range, text, format_dict) (default: None).
///                    Example: [("A1:D1", "Title", {"bold": True, "bg_color": "#4F81BD"})]
///     hyperlinks: List of link specs: (cell_ref, url) or (cell_ref, url, display_text) (default: None).
///                 Example: [("A1", "https://example.com", "Click here")]
///     comments: Dict mapping cell refs to note text or config dict (default: None).
///               Example: {"A1": "Note text"} or {"A1": {"text": "Note", "author": "John"}}
///     validations: Dict mapping column names/patterns to validation configs (default: None).
///                  Types: list, whole_number, decimal, text_length
///                  (aliases accepted, e.g. integer/number/length — the full table is at
///                  https://tstone-1.github.io/xlsxturbo/data-validation/).
///                  Example: {"status": {"type": "list", "values": ["Open", "Closed"]}}
///                  For "whole_number", min/max are bounded to the i32 range
///                  (-2147483648..=2147483647); a value outside that range raises
///                  ValueError naming the field and range.
///     rich_text: Dict mapping cell refs to lists of formatted text segments (default: None).
///                Example: {"A1": [("Bold text", {"bold": True}), (" normal text",)]}
///     images: Dict mapping cell refs to image paths or config dicts (default: None).
///             Example: {"A1": "logo.png"} or {"A1": {"path": "logo.png", "scale_width": 0.5}}
///     checkboxes: Dict mapping cell refs to checkbox state (default: None).
///                 Simple form: {"A1": True, "A2": False}
///                 Dict form with optional cell format: {"A3": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}
///     textboxes: Dict mapping cell refs to floating text shapes (default: None).
///                Simple form: {"B2": "Some text"}
///                Dict form: {"B2": {"text": "Note", "width": 200, "height": 100,
///                            "x_offset": 10, "y_offset": 5,
///                            "font": {"name": "Arial", "size": 14, "bold": True, "color": "#FF0000"},
///                            "fill_color": "#F0F0F0", "line_color": "#000000",
///                            "alt_text": "Descriptive alt text"}}
///     charts: Dict mapping cell refs to native Excel chart configs (default: None).
///             "data_range"/"values_range"/"values" and "categories_range"/"categories"
///             (including a chart-level fallback used by series without their own)
///             must include a sheet name (e.g. "Sheet1!$B$2:$B$10"); a bare range
///             raises ValueError.
///             Example: {"D2": {"type": "bar", "data_range": "Sheet1!$B$2:$B$10",
///                       "categories_range": "Sheet1!$A$2:$A$10", "title": "Monthly Activity"}}
///     sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config (default: None).
///                 A single-cell key (e.g. "D2") places one sparkline; a range key (e.g. "D2:D10")
///                 places a grouped sparkline, one per row of the data range.
///                 Required key: "range" (the data to plot, sheet-qualified like a chart range,
///                 e.g. "Sheet1!A2:C10"). Options: type ("line", "column", "win_loss"),
///                 style (1-36), markers, high_point, low_point, first_point, last_point,
///                 negative_points, show_axis, color and the *_point colors, line_weight,
///                 custom_max, custom_min, group_max, group_min, date_range.
///                 Example: {"D2:D10": {"range": "Sheet1!A2:C10", "type": "line", "markers": True}}
///     startrow: Zero-based sheet row of the header, as in pandas' to_excel; the data
///               starts on the row below it (default: 0). The rows above stay free for
///               cells, merged_ranges and rich_text. Cell references in every option stay
///               absolute sheet positions, and so do row_heights keys. The table, freeze
///               panes, formula columns, conditional formats and validations move with
///               the frame. The returned row count does not include the offset.
///     autofilter: Put filter buttons on the header row of the written range, without
///                 making it an Excel table (default: False). The range is the header row
///                 and the data rows, formula columns included, and it follows startrow.
///                 Only the buttons are added; no rows are hidden. Raises
///                 ConfigurationError together with table_style (a table has its own
///                 filter) and with header=False.
///     tab_color: Colour of the sheet tab: "#RRGGBB" or a named colour such as "red"
///                (default: None).
///     compression_level: Zip compression level of the file, 1 (fastest, largest)
///                        to 9 (default: None, which is level 6).
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (None, str, int, float, bool, date, datetime) or dicts with
///            "value" and an optional "format" dict (the keys column_formats accepts), plus the
///            shorthand keys "num_format", "font_name", "quote_prefix", "align_horizontal",
///            "align_vertical" and "wrap_text", which override the same key in "format".
///            Cells are written after all DataFrame data, so they can overwrite data cells.
///            Example: {"B9": "Label", "D6": {"value": "934728173849", "num_format": "@"}}
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     XlsxTurboError: If the conversion fails. The subclass says what kind of
///         failure it was -- ConfigurationError / ConfigurationTypeError for an
///         option, InputDataError for the frame, FileError for the filesystem.
///         Every one of them is also the builtin its failures raised before
///         0.19.0, so `except ValueError` still works; note
///         ConfigurationTypeError is a TypeError rather than a ValueError.
///         Full contract: https://tstone-1.github.io/xlsxturbo/errors/
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [30, 25]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
///     >>> # With table formatting and auto-width columns:
///     >>> xlsxturbo.df_to_xlsx(df, "styled.xlsx", table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With custom column widths and row heights:
///     >>> xlsxturbo.df_to_xlsx(df, "custom.xlsx", column_widths={0: 25, 1: 10}, row_heights={0: 20})
///     >>> # For very large files, use constant_memory mode:
///     >>> xlsxturbo.df_to_xlsx(large_df, "big.xlsx", constant_memory=True)
#[pyfunction]
#[pyo3(signature = (
    df,
    output_path,
    sheet_name = "Sheet1",
    header = true,
    autofit = false,
    table_style = None,
    freeze_panes = None,
    column_widths = None,
    table_name = None,
    header_format = None,
    row_heights = None,
    constant_memory = false,
    column_formats = None,
    conditional_formats = None,
    formula_columns = None,
    merged_ranges = None,
    hyperlinks = None,
    comments = None,
    validations = None,
    rich_text = None,
    images = None,
    checkboxes = None,
    textboxes = None,
    charts = None,
    defined_names = None,
    cells = None,
    sparklines = None,
    startrow = None,
    autofilter = false,
    tab_color = None,
    compression_level = None,
))]
#[allow(clippy::too_many_arguments)]
fn df_to_xlsx<'py>(
    py: Python<'py>,
    df: &Bound<'py, PyAny>,
    output_path: &Bound<'py, PyAny>,
    sheet_name: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: Option<&Bound<'py, PyAny>>,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<&Bound<'py, PyAny>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
    comments: Option<&Bound<'py, PyAny>>,
    validations: Option<&Bound<'py, PyAny>>,
    rich_text: Option<&Bound<'py, PyAny>>,
    images: Option<&Bound<'py, PyAny>>,
    checkboxes: Option<&Bound<'py, PyAny>>,
    textboxes: Option<&Bound<'py, PyAny>>,
    charts: Option<&Bound<'py, PyAny>>,
    defined_names: Option<&Bound<'py, PyAny>>,
    cells: Option<&Bound<'py, PyAny>>,
    sparklines: Option<&Bound<'py, PyAny>>,
    startrow: Option<&Bound<'py, PyAny>>,
    autofilter: bool,
    tab_color: Option<&Bound<'py, PyAny>>,
    compression_level: Option<&Bound<'py, PyAny>>,
) -> PyResult<(u32, u16)> {
    let output = OutputArg::from_py(output_path)?;
    require_supported_dataframe(df, None)?;
    let compression_level = extract_compression_level_kwarg(compression_level)?;
    let startrow = extract_startrow_kwarg(startrow)?;
    let freeze_panes = extract_freeze_panes_kwarg(freeze_panes)?;
    let tab_color = extract_tab_color_kwarg(tab_color)?;
    let defined_names = extract_defined_names_kwarg(defined_names)?;
    let row_heights = extract_row_heights_kwarg(row_heights)?;
    let opts = extract_options(&RawOptions {
        column_widths,
        header_format,
        column_formats,
        conditional_formats,
        formula_columns,
        merged_ranges,
        hyperlinks,
        comments,
        validations,
        rich_text,
        images,
        checkboxes,
        textboxes,
        charts,
        sparklines,
        cells,
    })?;

    // A table name colliding with a defined name is refused before anything is
    // written, so it reports as a configuration failure rather than arriving
    // from the save as a `FileError`. Skipped entirely without defined names,
    // since nothing else here needs the shape. The `1` is this workbook's only
    // possible table id, so a sheet asking for a style without a name claims
    // the `Table1` the writer will give it.
    if defined_names.is_some() {
        let mut tables: HashMap<String, ClaimedTableName> = HashMap::new();
        if let Some(name) = claimed_table_name(
            df,
            constant_memory,
            header,
            table_style,
            table_name.as_deref(),
            1,
        )
        .map_err(errors::input_data)?
        {
            tables.insert(
                name.to_lowercase(),
                ClaimedTableName {
                    name,
                    sheet: sheet_name.to_string(),
                    auto: table_name.is_none(),
                },
            );
        }
        reject_table_name_collisions(&tables, defined_names.as_ref())
            .map_err(errors::workbook_validation)?;
    }

    let config = WriteConfig {
        include_header: header,
        autofit,
        table_style,
        freeze_panes,
        autofilter,
        tab_color,
        startrow,
        table_name: table_name.as_deref(),
        row_heights: row_heights.as_ref(),
        constant_memory,
    };

    let mut buffer = Vec::new();
    let result = convert_dataframe_to_xlsx(
        py,
        df,
        output.destination(&mut buffer),
        sheet_name,
        &config,
        &opts,
        WorkbookOptions {
            defined_names: defined_names.as_ref(),
            compression_level,
        },
    )?;
    output.deliver(&buffer)?;
    Ok(result)
}

/// Get the version of the xlsxturbo library
#[pyfunction]
fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

/// Write multiple DataFrames to separate sheets in a single workbook.
///
/// This is a convenience function that writes multiple DataFrames to
/// separate sheets in one workbook, which is more efficient than
/// calling df_to_xlsx multiple times.
///
/// Args:
///     sheets: List of tuples. Each tuple can be:
///             - (DataFrame, sheet_name) - uses global defaults
///             - (DataFrame, sheet_name, options_dict) - per-sheet overrides
///             Several tuples may name the same sheet: each writes its frame at
///             its own startrow. The frames may not overlap, and the first tuple
///             for a sheet sets autofit, freeze_panes, tab_color, column_widths
///             and row_heights.
///             Options dict keys: header, autofit, table_style, freeze_panes,
///             autofilter, tab_color, column_widths, row_heights, table_name, header_format, column_formats,
///             conditional_formats, formula_columns, merged_ranges, hyperlinks,
///             comments, validations, rich_text, images, checkboxes, textboxes, charts,
///             sparklines, cells, startrow
///     output_path: Path for the output XLSX file, or a binary file-like
///         object (io.BytesIO, a file opened 'wb') to write the workbook into
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///              Combined with column_widths: explicit widths win for the columns
///              they name; every other column is still autofitted (rather than
///              left at Excel's default width). Add an "_all" entry in
///              column_widths to cap the autofit width instead of overriding it.
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: True freezes the header row (default: False). A (row, col) tuple
///                   freezes at that zero-based cell instead, as XlsxWriter's
///                   freeze_panes(row, col) does: it is the top-left cell of the scrolling
///                   area, so (4, 1) keeps four rows and one column in view. True follows
///                   startrow and does nothing with header=False; a tuple is taken as given.
///     column_widths: Dict mapping column index or "_all" to width in characters (default: None)
///                    Example: {0: 20, "_all": 50} sets col A to 20, caps others at 50. An
///                    integer key must be a non-negative index within Excel's column range
///                    (0..=16383); a negative key, a key beyond 16383, or a
///                    non-integer/non-"_all" key raises. With autofit=True and no "_all"
///                    key: listed columns get the explicit width, unlisted columns are
///                    autofitted. With autofit=True and an "_all" key: "_all" caps the
///                    autofit width for unlisted columns instead of overriding it.
///     table_name: Name for Excel table (requires table_style; default: auto-generated,
///                 `Table1`, `Table2`, ... in sheet order).
///                 Must be alphanumeric/underscore, max 255 chars; anything else is
///                 rewritten. Effective names must be unique across the workbook after
///                 that rewrite, and must not collide with a defined_names entry.
///     header_format: Dict with header cell formatting options (default: None)
///                    Example: {"bold": True, "bg_color": "#4F81BD", "font_color": "white"}
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Emits a RuntimeWarning and disables: table_style,
///                      row_heights, autofit, column_widths with autofit cap, conditional_formats,
///                      formula_columns, merged_ranges, hyperlinks, comments, validations,
///                      rich_text, images, checkboxes, textboxes, charts, sparklines, and cells.
///                      Plain column_widths, header_format, column_formats, freeze_panes,
///                      autofilter and tab_color remain supported.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline, border.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names/patterns to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     formula_columns: Dict mapping column names to Excel formula templates (default: None).
///                      Use {row} as placeholder for the current row number.
///     merged_ranges: List of merge specs: (range, text) or (range, text, format_dict) (default: None).
///     hyperlinks: List of link specs: (cell_ref, url) or (cell_ref, url, display_text) (default: None).
///     comments: Dict mapping cell refs to note text or config dict (default: None).
///     validations: Dict mapping column names/patterns to validation configs (default: None).
///                  Types: list, whole_number, decimal, text_length
///                  (aliases accepted, e.g. integer/number/length — the full table is at
///                  https://tstone-1.github.io/xlsxturbo/data-validation/).
///                  For "whole_number", min/max are bounded to the i32 range
///                  (-2147483648..=2147483647); a value outside that range raises
///                  ValueError naming the field and range.
///     rich_text: Dict mapping cell refs to lists of formatted text segments (default: None).
///     images: Dict mapping cell refs to image paths or config dicts (default: None).
///     checkboxes: Dict mapping cell refs to checkbox state (bool) or config dict (default: None).
///                 Example: {"A1": True} or {"A1": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}
///     textboxes: Dict mapping cell refs to floating text shapes (default: None).
///                Example: {"B2": "text"} or {"B2": {"text": "Note", "width": 200, "font": {"bold": True}}}
///     charts: Dict mapping cell refs to native Excel chart configs (default: None).
///             "data_range"/"values_range"/"values" and "categories_range"/"categories"
///             (including a chart-level fallback used by series without their own)
///             must include a sheet name (e.g. "Sheet1!$B$2:$B$10"); a bare range
///             raises ValueError.
///     sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config (default: None).
///                 Range key (e.g. "D2:D10") makes a grouped sparkline; single cell makes one.
///                 "range" must be sheet-qualified, e.g. "Sheet1!A2:C10".
///                 Example: {"D2:D10": {"range": "Sheet1!A2:C10", "type": "line", "markers": True}}
///     startrow: Zero-based sheet row of the header, as in pandas' to_excel; the data
///               starts on the row below it (default: 0). The rows above stay free for
///               cells, merged_ranges and rich_text. Cell references in every option stay
///               absolute sheet positions, and so do row_heights keys. The table, freeze
///               panes, formula columns, conditional formats and validations move with
///               the frame. The returned row count does not include the offset.
///     autofilter: Put filter buttons on the header row of the written range, without
///                 making it an Excel table (default: False). The range is the header row
///                 and the data rows, formula columns included, and it follows startrow.
///                 Only the buttons are added; no rows are hidden. Raises
///                 ConfigurationError together with table_style (a table has its own
///                 filter) and with header=False.
///     tab_color: Colour of the sheet tab: "#RRGGBB" or a named colour such as "red"
///                (default: None).
///     compression_level: Zip compression level of the file, 1 (fastest, largest)
///                        to 9 (default: None, which is level 6).
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (None, str, int, float, bool, date, datetime) or dicts with
///            "value" and an optional "format" dict (the keys column_formats accepts), plus the
///            shorthand keys "num_format", "font_name", "quote_prefix", "align_horizontal",
///            "align_vertical" and "wrap_text", which override the same key in "format".
///            Cells are written after all DataFrame data, so they can overwrite data cells.
///            Example: {"B9": "Label", "D6": {"value": "934728173849", "num_format": "@"}}
///
/// Returns:
///     List of (rows, columns) tuples for each sheet
///
/// Raises:
///     XlsxTurboError: If the conversion fails. The subclass says what kind of
///         failure it was -- ConfigurationError / ConfigurationTypeError for an
///         option, InputDataError for the frame, FileError for the filesystem.
///         Every one of them is also the builtin its failures raised before
///         0.19.0, so `except ValueError` still works; note
///         ConfigurationTypeError is a TypeError rather than a ValueError.
///         Full contract: https://tstone-1.github.io/xlsxturbo/errors/
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df1 = pd.DataFrame({'a': [1, 2]})
///     >>> df2 = pd.DataFrame({'b': [3, 4]})
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sheet1"), (df2, "Sheet2")], "out.xlsx")
///     >>> # With styling applied to all sheets:
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sales"), (df2, "Regions")], "report.xlsx",
///     ...                       table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With per-sheet options (header=False for one sheet):
///     >>> xlsxturbo.dfs_to_xlsx([
///     ...     (df1, "Data", {"header": True, "table_style": "Medium2"}),
///     ...     (df2, "Instructions", {"header": False})
///     ... ], "report.xlsx", autofit=True)
#[pyfunction]
#[pyo3(signature = (
    sheets,
    output_path,
    header = true,
    autofit = false,
    table_style = None,
    freeze_panes = None,
    column_widths = None,
    table_name = None,
    header_format = None,
    row_heights = None,
    constant_memory = false,
    column_formats = None,
    conditional_formats = None,
    formula_columns = None,
    merged_ranges = None,
    hyperlinks = None,
    comments = None,
    validations = None,
    rich_text = None,
    images = None,
    checkboxes = None,
    textboxes = None,
    charts = None,
    defined_names = None,
    cells = None,
    sparklines = None,
    startrow = None,
    autofilter = false,
    tab_color = None,
    compression_level = None,
))]
#[allow(clippy::too_many_arguments)]
fn dfs_to_xlsx<'py>(
    py: Python<'py>,
    sheets: Vec<Bound<'py, PyAny>>,
    output_path: &Bound<'py, PyAny>,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: Option<&Bound<'py, PyAny>>,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<&Bound<'py, PyAny>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
    comments: Option<&Bound<'py, PyAny>>,
    validations: Option<&Bound<'py, PyAny>>,
    rich_text: Option<&Bound<'py, PyAny>>,
    images: Option<&Bound<'py, PyAny>>,
    checkboxes: Option<&Bound<'py, PyAny>>,
    textboxes: Option<&Bound<'py, PyAny>>,
    charts: Option<&Bound<'py, PyAny>>,
    defined_names: Option<&Bound<'py, PyAny>>,
    cells: Option<&Bound<'py, PyAny>>,
    sparklines: Option<&Bound<'py, PyAny>>,
    startrow: Option<&Bound<'py, PyAny>>,
    autofilter: bool,
    tab_color: Option<&Bound<'py, PyAny>>,
    compression_level: Option<&Bound<'py, PyAny>>,
) -> PyResult<Vec<(u32, u16)>> {
    let output = OutputArg::from_py(output_path)?;
    if sheets.is_empty() {
        return Err(errors::configuration(
            "dfs_to_xlsx requires at least one sheet, got an empty list",
        ));
    }
    let row_heights = extract_row_heights_kwarg(row_heights)?;
    let startrow = extract_startrow_kwarg(startrow)?;
    let freeze_panes = extract_freeze_panes_kwarg(freeze_panes)?;
    let tab_color = extract_tab_color_kwarg(tab_color)?;
    let defined_names = extract_defined_names_kwarg(defined_names)?;
    let compression_level = extract_compression_level_kwarg(compression_level)?;
    let mut workbook = Workbook::new();
    let mut stats = Vec::new();
    let mut table_names: HashMap<String, ClaimedTableName> = HashMap::new();

    let opts = extract_options(&RawOptions {
        column_widths,
        header_format,
        column_formats,
        conditional_formats,
        formula_columns,
        merged_ranges,
        hyperlinks,
        comments,
        validations,
        rich_text,
        images,
        checkboxes,
        textboxes,
        charts,
        sparklines,
        cells,
    })?;

    // The id rust_xlsxwriter will give the next table it names itself. It counts
    // tables, not sheets, so it advances only where one is actually created --
    // including for an explicitly named table, which consumes an id without
    // using the generated name. See `claimed_table_name` for the measurement.
    let mut next_table_id: u32 = 1;

    // First pass: resolve every entry and check all of them, so that nothing
    // below can fail on how the entries relate to each other once rows are
    // being written.
    let mut entries: Vec<SheetEntry<'py>> = Vec::with_capacity(sheets.len());
    let mut first_by_name: HashMap<String, usize> = HashMap::new();

    for (index, sheet_tuple) in sheets.into_iter().enumerate() {
        let (df, sheet_name, sheet_config) = extract_sheet_info(&sheet_tuple, index)?;
        require_supported_dataframe(&df, Some(&sheet_name))?;

        // Several entries may name one sheet: each writes a frame to it. The
        // fold is the writer's own, which compares sheet names with
        // `to_lowercase` -- 'Summary' and 'summary' are one sheet to Excel.
        let first = *first_by_name
            .entry(sheet_name.to_lowercase())
            .or_insert(index);
        let is_first = first == index;
        if !is_first {
            reject_whole_sheet_options(&sheet_config, index, &sheet_name, first)?;
        }

        // Merge per-sheet scalar options with global defaults
        let effective_header = sheet_config.header.unwrap_or(header);
        // A later frame needs the sheet's autofit only to know whether to
        // estimate its widths for the capped fit.
        let effective_autofit = if is_first {
            sheet_config.autofit.unwrap_or(autofit)
        } else {
            entries[first].autofit
        };
        let effective_table_style: Option<String> = match &sheet_config.table_style {
            Some(style_opt) => style_opt.clone(),
            None => table_style.map(|s| s.to_string()),
        };
        let effective_table_name: Option<String> = sheet_config
            .table_name
            .as_ref()
            .cloned()
            .or_else(|| table_name.clone());

        entries.push(SheetEntry {
            df,
            header: effective_header,
            autofit: effective_autofit,
            table_style: effective_table_style,
            freeze_panes: sheet_config.freeze_panes.unwrap_or(freeze_panes),
            autofilter: sheet_config.autofilter.unwrap_or(autofilter),
            tab_color: sheet_config.tab_color.unwrap_or(tab_color),
            startrow: sheet_config.startrow.unwrap_or(startrow),
            table_name: effective_table_name,
            name: sheet_name,
            config: sheet_config,
            first,
        });
    }

    // The writer assigns IDs in worksheet order, then frame order within a
    // worksheet, even when the caller interleaves entries for different sheets.
    // Stable sorting by the first entry reproduces that order without changing
    // the order in which frames are written.
    let mut table_entries: Vec<_> = entries.iter().collect();
    table_entries.sort_by_key(|entry| entry.first);
    for entry in table_entries {
        // A table is only actually created when there's at least one data row
        // and one column (see the gate in `apply_frame_features`), so an
        // empty DataFrame never claims a table name here either — otherwise
        // two empty sheets sharing a table name would false-positive as a
        // conflict.
        if let Some(sanitized) = claimed_table_name(
            &entry.df,
            constant_memory,
            entry.header,
            entry.table_style.as_deref(),
            entry.table_name.as_deref(),
            next_table_id,
        )
        .map_err(|e| errors::input_data(format!("sheet '{}': {}", entry.name, e)))?
        {
            next_table_id += 1;
            let auto = entry.table_name.is_none();
            // `to_lowercase`, not `to_ascii_lowercase`: this fold has to match
            // the writer's own uniqueness check exactly, and rust_xlsxwriter
            // folds names with `to_lowercase`. A narrower fold under-reaches --
            // sanitization keeps every alphanumeric codepoint, so 'TÄBLE' and
            // 'täble' both reach the workbook, pass this check, and collide
            // during the save, where the message names neither sheet. A wider
            // one would refuse names Excel accepts, which is the more expensive
            // direction.
            let key = sanitized.to_lowercase();
            let claim = ClaimedTableName {
                name: sanitized,
                sheet: entry.name.clone(),
                auto,
            };
            if let Some(previous) = table_names.insert(key, claim) {
                return Err(errors::workbook_validation(format!(
                    "Duplicate table name '{}' for sheets '{}' and '{}'. Excel table names must be unique within a workbook.{}",
                    previous.name,
                    previous.sheet,
                    entry.name,
                    workbook::auto_table_name_hint(previous.auto || auto)
                )));
            }
        }
    }

    check_frame_layout(&entries, constant_memory)?;

    // Merge per-sheet complex options with global defaults (references, no
    // cloning needed). A second or later frame on a sheet keeps the frame-bound
    // ones and only its own cell-addressed ones: see `for_later_frame`.
    let mut effective: Vec<EffectiveOpts<'_>> = Vec::with_capacity(entries.len());
    for (index, entry) in entries.iter().enumerate() {
        let merged = entry.config.merge_with(&opts);
        effective.push(if entry.first == index {
            merged
        } else {
            let sheet_column_widths = effective[entry.first].column_widths;
            merged.for_later_frame(&entry.config, sheet_column_widths)
        });
    }
    let write_config = |index: usize| {
        let entry = &entries[index];
        let sheet_row_heights = entry.config.row_heights.as_ref().or(row_heights.as_ref());
        entry.write_config(sheet_row_heights, constant_memory)
    };
    let last_on_sheet = |index: usize| {
        let first = entries[index].first;
        !entries[index + 1..]
            .iter()
            .any(|later| later.first == first)
    };

    // Second pass: write the frames in list order, and finish a sheet as soon
    // as its last frame is on it.
    let mut frames: Vec<FrameOutcome> = Vec::with_capacity(entries.len());
    for (index, entry) in entries.iter().enumerate() {
        let in_sheet = |e: String| errors::configuration(format!("sheet '{}': {}", entry.name, e));
        // The sheet carries the first entry's spelling of its name.
        let sheet = entries[entry.first].name.as_str();
        let config = write_config(index);

        let worksheet = if entry.first == index {
            add_named_worksheet(&mut workbook, sheet, constant_memory).map_err(|e| match e {
                ConvertError::Config(message) => in_sheet(message),
                ConvertError::File(failure) => errors::file(failure),
            })?
        } else {
            workbook
                .worksheet_from_name(sheet)
                .map_err(|e| in_sheet(e.to_string()))?
        };
        let frame = write_frame(py, worksheet, sheet, &entry.df, &config, &effective[index])
            .map_err(in_sheet)?;
        stats.push((frame.rows, frame.total_cols));
        frames.push(frame);

        if last_on_sheet(index) {
            let members: Vec<usize> = (entry.first..=index)
                .filter(|&member| entries[member].first == entry.first)
                .collect();
            let on_sheet: Vec<&FrameOutcome> = members.iter().map(|&m| &frames[m]).collect();
            let (col_count, content_widths) = sheet_extent(&on_sheet);

            // The whole-sheet options are the first entry's.
            finish_sheet(
                worksheet,
                &write_config(entry.first),
                &effective[entry.first],
                col_count,
                &content_widths,
            )
            .map_err(in_sheet)?;
            // The cell-addressed options come last, entry by entry, so that
            // `cells` can overwrite the data of any frame on the sheet.
            for &member in &members {
                apply_cell_options(py, worksheet, &write_config(member), &effective[member])
                    .map_err(|e| {
                        errors::configuration(format!("sheet '{}': {}", entries[member].name, e))
                    })?;
            }
        }
    }

    reject_table_name_collisions(&table_names, defined_names.as_ref())
        .map_err(errors::workbook_validation)?;

    // The same "finish" as the single-sheet path, and the same code: the
    // defined names and the detached save were written out twice until 1.3.x,
    // which is how the GIL release first shipped in a draft that covered only
    // one of them. See `convert::finish_workbook`.
    let mut buffer = Vec::new();
    finish_workbook(
        py,
        &mut workbook,
        WorkbookOptions {
            defined_names: defined_names.as_ref(),
            compression_level,
        },
        output.destination(&mut buffer),
    )?;
    output.deliver(&buffer)?;

    Ok(stats)
}

/// xlsxturbo - High-performance Excel writer
///
/// A Rust-powered library for converting DataFrames and CSV files to Excel XLSX format.
/// Substantially faster than pandas + openpyxl; the machine-labeled benchmark tables
/// (the durable, reproducible source for performance numbers) are on the performance
/// page: https://tstone-1.github.io/xlsxturbo/performance/
///
/// Features:
/// - Direct DataFrame support (pandas and polars)
/// - Automatic type detection (numbers, booleans, dates, datetimes)
/// - Proper Excel formatting for dates and times
/// - Handles NaN/Inf/None gracefully
/// - Memory-efficient for large files
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'a': [1, 2], 'b': [3.14, 2.71]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
#[pymodule]
fn xlsxturbo(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(csv_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(df_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(dfs_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    errors::register(m)?;
    Ok(())
}
