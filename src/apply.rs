//! Excel application functions for applying features to worksheets.
//!
//! The feature implementations live in focused submodules; this facade keeps
//! the public crate-local API stable for the conversion pipeline.

mod annotations;
mod cells;
mod charts;
mod conditional_formats;
mod dimensions;
mod formulas;
mod media;
mod rich_text;
mod validations;

pub(crate) use annotations::{apply_comments, apply_hyperlinks, apply_merged_ranges};
pub(crate) use cells::apply_cells;
pub(crate) use charts::apply_charts;
pub(crate) use conditional_formats::apply_conditional_formats;
pub(crate) use dimensions::{apply_column_widths, apply_column_widths_with_autofit_cap};
pub(crate) use formulas::apply_formula_columns;
pub(crate) use media::{apply_checkboxes, apply_images, apply_textboxes};
pub(crate) use rich_text::apply_rich_text;
pub(crate) use validations::apply_validations;
