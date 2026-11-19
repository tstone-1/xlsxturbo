//! Native Excel pivot tables over the frame they are given with.
//!
//! A pivot table's source is always its frame's written range, header row and
//! formula columns included, so every field is a header the frame wrote. Its
//! location is `"A3"` (the frame's own sheet) or `"Summary!A3"` (another sheet,
//! created empty after every other sheet when the workbook has none of that
//! name). This runs once per workbook, after every frame is written, because a
//! pivot table may go on a sheet that a later entry writes.
//!
//! rust_xlsxwriter writes the definition without cached data: Excel builds the
//! table from the source when the file is opened, so its size is unknown here.
//! The crate checks the first column of the table for overlaps; this module
//! adds what xlsxturbo knows and the crate does not: where every frame is.
//!
//! The field names are always passed with `set_field_names()`, the header texts
//! as `write_frame` wrote them. A constant-memory sheet needs that (the crate
//! cannot read its flushed header row). On any other sheet the crate compares
//! the names with the header cells when the workbook is saved, and the
//! cell-addressed options (`cells`, `hyperlinks`, `merged_ranges`, `rich_text`,
//! `checkboxes`) run after `write_frame` and may overwrite a header cell. So a
//! write into the header row of a frame with pivot tables is refused here,
//! before the save; left to the crate it fails the save as a `FileError`.
//!
//! The same holds for the crate's other save-time rule, a value field caption
//! equal to a column name without regard to case: `build_pivot_table` refuses it.
//!
//! The caller creates the sheets a plan needs ([`PivotPlan::new_sheets`]) with
//! its own error mapping, then adds the tables ([`PivotPlan::add_to`]).

use crate::parse::{parse_cell_range, parse_cell_ref, MAX_ROW_1BASED};
use crate::types::{
    pydict_to_hashmap_str, sequence_as_list, EffectiveOpts, OptionMap, PivotTableConfig,
};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{
    PivotTable, PivotTableFunction, PivotTableLayout, PivotTableStyle, PivotTableValueField,
    Workbook,
};
use std::collections::HashMap;

const PIVOT_KEYS: &[&str] = &[
    "rows",
    "columns",
    "filters",
    "values",
    "style",
    "layout",
    "row_grand_totals",
    "column_grand_totals",
];

const VALUE_FIELD_KEYS: &[&str] = &["field", "function", "name", "num_format"];

/// The cells one frame wrote, header and formula columns included.
pub(crate) struct FrameArea {
    /// The sheet's name as the workbook spells it.
    pub(crate) sheet: String,
    pub(crate) first_row: u32,
    pub(crate) first_col: u16,
    pub(crate) last_row: u32,
    pub(crate) last_col: u16,
}

impl FrameArea {
    fn contains(&self, row: u32, col: u16) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_col..=self.last_col).contains(&col)
    }

    fn range_text(&self) -> String {
        format!(
            "{}:{}",
            rust_xlsxwriter::utility::row_col_to_cell(self.first_row, self.first_col),
            rust_xlsxwriter::utility::row_col_to_cell(self.last_row, self.last_col)
        )
    }
}

/// One frame's `pivot_tables`, with what they need to know about the frame.
pub(crate) struct PivotRequest<'a> {
    /// The frame's written cells; `None` when it wrote nothing at all.
    pub(crate) area: Option<&'a FrameArea>,
    /// The frame's sheet, for messages, as the workbook spells it.
    pub(crate) sheet: &'a str,
    pub(crate) header: bool,
    /// The header texts, in column order: frame columns, then formula columns.
    pub(crate) field_names: &'a [String],
    pub(crate) tables: &'a IndexMap<String, PivotTableConfig>,
}

/// The cells one cell-addressed option entry writes a value into.
pub(crate) struct CellWriteSpan {
    /// The sheet, as the workbook spells it.
    sheet: String,
    /// The option and key as the caller wrote them, e.g. `cells['A1']`.
    option: String,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
}

impl CellWriteSpan {
    fn new(sheet: &str, option: String, span: (u32, u16, u32, u16)) -> Self {
        let (first_row, first_col, last_row, last_col) = span;
        CellWriteSpan {
            sheet: sheet.to_string(),
            option,
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }
}

/// Every cell value one entry's cell-addressed options write on `sheet`:
/// `cells`, `hyperlinks`, `merged_ranges` (every cell of the range), `rich_text`
/// and `checkboxes`. Comments, images, textboxes, charts and sparklines write
/// no cell value. A reference that does not parse is left out; the option's own
/// step has refused it before this runs.
pub(crate) fn cell_writes(sheet: &str, opts: &EffectiveOpts<'_>) -> Vec<CellWriteSpan> {
    let mut writes = Vec::new();
    let mut one_cell = |option: String, cell: &str| {
        if let Ok((row, col)) = parse_cell_ref(cell) {
            writes.push(CellWriteSpan::new(sheet, option, (row, col, row, col)));
        }
    };
    for (cell, _, _) in opts.hyperlinks.into_iter().flatten() {
        one_cell(format!("hyperlinks['{}']", cell), cell);
    }
    for cell in opts.rich_text.into_iter().flat_map(|m| m.keys()) {
        one_cell(format!("rich_text['{}']", cell), cell);
    }
    for cell in opts.checkboxes.into_iter().flat_map(|m| m.keys()) {
        one_cell(format!("checkboxes['{}']", cell), cell);
    }
    for (range, _, _) in opts.merged_ranges.into_iter().flatten() {
        if let Ok(span) = parse_cell_range(range) {
            writes.push(CellWriteSpan::new(
                sheet,
                format!("merged_ranges['{}']", range),
                span,
            ));
        }
    }
    for cell in opts.cells.into_iter().flatten() {
        writes.push(CellWriteSpan::new(
            sheet,
            cell.context.clone(),
            (cell.row, cell.col, cell.row, cell.col),
        ));
    }
    writes
}

/// Refuse a cell-addressed write into the frame's header row: the header texts
/// are the pivot table's field names, and the crate compares them with the
/// cells only when the workbook is saved.
fn reject_header_writes(
    context: &str,
    area: &FrameArea,
    writes: &[CellWriteSpan],
) -> Result<(), String> {
    let header_row = area.first_row;
    let sheet = area.sheet.to_lowercase();
    let hit = writes.iter().find(|w| {
        w.sheet.to_lowercase() == sheet
            && (w.first_row..=w.last_row).contains(&header_row)
            && w.first_col <= area.last_col
            && area.first_col <= w.last_col
    });
    let Some(write) = hit else {
        return Ok(());
    };
    let cell = |col: u16| rust_xlsxwriter::utility::row_col_to_cell(header_row, col);
    Err(format!(
        "{}: {} writes cell {}, in the header row {}:{} of the frame the pivot table \
         is built from; the header texts are its field names, so keep that write \
         out of the header row",
        context,
        write.option,
        cell(write.first_col.max(area.first_col)),
        cell(area.first_col),
        cell(area.last_col),
    ))
}

/// A location key split into its sheet (if any) and its cell.
struct Location {
    sheet: Option<String>,
    row: u32,
    col: u16,
}

fn parse_location(key: &str) -> Result<Location, String> {
    let (sheet, cell) = match key.rsplit_once('!') {
        None => (None, key),
        Some((sheet_part, cell)) => {
            let sheet = match sheet_part
                .strip_prefix('\'')
                .and_then(|s| s.strip_suffix('\''))
            {
                Some(quoted) => quoted.replace("''", "'"),
                // `sheet_name` is taken as given, never trimmed, so white space
                // around an unquoted sheet part is refused rather than trimmed:
                // trimming would pick another sheet than the one the text names.
                None if sheet_part.trim() != sheet_part => {
                    return Err(format!(
                        "sheet name '{}' has white space around it; write a sheet \
                         name with leading or trailing spaces in quotes, as '{}'!{}",
                        sheet_part,
                        sheet_part.replace('\'', "''"),
                        cell.trim()
                    ));
                }
                None => sheet_part.to_string(),
            };
            if sheet.is_empty() {
                return Err("no sheet name before '!'".to_string());
            }
            (Some(sheet), cell)
        }
    };
    if cell.contains(':') {
        return Err("a pivot table is placed at one cell, not at a range".to_string());
    }
    let (row, col) = parse_cell_ref(cell)?;
    // `parse_cell_ref` leaves the grid to its callers; XFD1048576 is the last cell.
    if row >= MAX_ROW_1BASED {
        return Err(format!(
            "cell '{}' is below Excel's last row, 1048576",
            cell
        ));
    }
    Ok(Location { sheet, row, col })
}

/// Parse a pivot table style name: the gallery name without its `PivotStyle` prefix.
fn parse_pivot_style(style: &str) -> Result<PivotTableStyle, String> {
    macro_rules! styles {
        ($($name:ident),+ $(,)?) => {
            match style {
                "None" => Ok(PivotTableStyle::None),
                $( stringify!($name) => Ok(PivotTableStyle::$name), )+
                _ => Err(format!(
                    "unknown style '{}'. Valid styles: Light1-Light28, Medium1-Medium28, Dark1-Dark28, None",
                    style
                )),
            }
        };
    }
    styles!(
        Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10, Light11,
        Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20, Light21,
        Light22, Light23, Light24, Light25, Light26, Light27, Light28, Medium1, Medium2, Medium3,
        Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10, Medium11, Medium12,
        Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20, Medium21,
        Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28, Dark1, Dark2, Dark3,
        Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11, Dark12, Dark13, Dark14, Dark15,
        Dark16, Dark17, Dark18, Dark19, Dark20, Dark21, Dark22, Dark23, Dark24, Dark25, Dark26,
        Dark27, Dark28,
    )
}

fn parse_layout(layout: &str) -> Result<PivotTableLayout, String> {
    match layout {
        "compact" => Ok(PivotTableLayout::Compact),
        "outline" => Ok(PivotTableLayout::Outline),
        "tabular" => Ok(PivotTableLayout::Tabular),
        _ => Err(format!(
            "unknown layout '{}'. Valid: compact, outline, tabular",
            layout
        )),
    }
}

fn parse_function(function: &str) -> Result<PivotTableFunction, String> {
    match function {
        "sum" => Ok(PivotTableFunction::Sum),
        "count" => Ok(PivotTableFunction::Count),
        "average" => Ok(PivotTableFunction::Average),
        "max" => Ok(PivotTableFunction::Max),
        "min" => Ok(PivotTableFunction::Min),
        "product" => Ok(PivotTableFunction::Product),
        "count_numbers" => Ok(PivotTableFunction::CountNumbers),
        "std_dev" => Ok(PivotTableFunction::StdDev),
        "std_dev_p" => Ok(PivotTableFunction::StdDevP),
        "var" => Ok(PivotTableFunction::Var),
        "var_p" => Ok(PivotTableFunction::VarP),
        _ => Err(format!(
            "unknown function '{}'. Valid: sum, count, average, max, min, product, \
             count_numbers, std_dev, std_dev_p, var, var_p",
            function
        )),
    }
}

/// A list option as its items, refusing a `str` (which Python would iterate
/// as one-letter column names) and anything else that is not a sequence.
fn list_items<'py>(view: &OptionMap<'py, '_>, key: &str) -> Result<Vec<Bound<'py, PyAny>>, String> {
    let Some(value) = view.get(key) else {
        return Ok(Vec::new());
    };
    let value = value.bind(view.py());
    if value.is_none() {
        return Ok(Vec::new());
    }
    let list = sequence_as_list(value)
        .map_err(|e| format!("{}: '{}': {}", view.context(), key, e))?
        .ok_or_else(|| {
            format!(
                "{}: '{}' must be a list of column names, got {}",
                view.context(),
                key,
                crate::types::pytype_name(value)
            )
        })?;
    Ok(list.iter().collect())
}

fn field_name(
    view: &OptionMap<'_, '_>,
    key: &str,
    item: &Bound<'_, PyAny>,
) -> Result<String, String> {
    item.extract::<String>().map_err(|_| {
        format!(
            "{}: '{}' must hold column names (str), got {}",
            view.context(),
            key,
            crate::types::pytype_name(item)
        )
    })
}

/// Check that `name` is one of the frame's headers.
fn require_field(context: &str, name: &str, field_names: &[String]) -> Result<(), String> {
    if field_names.iter().any(|f| f == name) {
        return Ok(());
    }
    Err(format!(
        "{}: column '{}' is not in the frame. Columns: {}",
        context,
        name,
        field_names.join(", ")
    ))
}

/// The frame must be a source Excel can build a pivot table from.
fn check_source(context: &str, request: &PivotRequest<'_>) -> Result<(), String> {
    if !request.header {
        return Err(format!(
            "{}: a pivot table needs the frame's header row for its field names; \
             it is written with header=False",
            context
        ));
    }
    if request.field_names.is_empty() {
        return Err(format!(
            "{}: a pivot table needs at least one column; the frame has none",
            context
        ));
    }
    let data_rows = request
        .area
        .as_ref()
        .map_or(0, |a| a.last_row - a.first_row);
    if data_rows == 0 {
        return Err(format!(
            "{}: a pivot table needs at least one data row; the frame has none",
            context
        ));
    }
    let mut seen: HashMap<String, &str> = HashMap::new();
    for name in request.field_names {
        if name.is_empty() {
            return Err(format!(
                "{}: a pivot table needs a name for every column; the frame has an empty column name",
                context
            ));
        }
        // Excel compares field names without case, and so does the crate.
        if let Some(previous) = seen.insert(name.to_lowercase(), name) {
            return Err(format!(
                "{}: a pivot table needs unique column names; the frame has '{}' and '{}'",
                context, previous, name
            ));
        }
    }
    Ok(())
}

/// Build one pivot table from its option dict, every field checked against the frame.
fn build_pivot_table(
    py: Python<'_>,
    context: &str,
    config: &PivotTableConfig,
    request: &PivotRequest<'_>,
    area: &FrameArea,
) -> Result<PivotTable, String> {
    let view = OptionMap::new(py, config, context.to_string());
    view.reject_unknown(PIVOT_KEYS)?;

    let mut table = PivotTable::new()
        .set_data_source((
            area.sheet.as_str(),
            area.first_row,
            area.first_col,
            area.last_row,
            area.last_col,
        ))
        .set_field_names(request.field_names);

    for key in ["rows", "columns", "filters"] {
        for item in list_items(&view, key)? {
            let name = field_name(&view, key, &item)?;
            require_field(context, &name, request.field_names)?;
            table = match key {
                "rows" => table.add_row_field(&name),
                "columns" => table.add_column_field(&name),
                _ => table.add_filter_field(&name),
            };
        }
    }

    for item in list_items(&view, "values")? {
        let value_field = if let Ok(name) = item.extract::<String>() {
            require_field(context, &name, request.field_names)?;
            PivotTableValueField::new(&name)
        } else if let Ok(dict) = item.cast::<pyo3::types::PyDict>() {
            let inner_context = format!("{}: values", context);
            let map = pydict_to_hashmap_str(dict, &inner_context)?;
            let inner = OptionMap::new(py, &map, inner_context);
            inner.reject_unknown(VALUE_FIELD_KEYS)?;
            let name = inner.required_string("field")?;
            require_field(context, &name, request.field_names)?;
            let mut value_field = PivotTableValueField::new(&name);
            if let Some(function) = inner.string("function")? {
                value_field = value_field.set_function(
                    parse_function(&function).map_err(|e| format!("{}: {}", context, e))?,
                );
            }
            if let Some(caption) = inner.string("name")? {
                // The crate refuses this only when the workbook is saved, as a
                // file error; Excel compares field names without case.
                let folded = caption.to_lowercase();
                if let Some(column) = request
                    .field_names
                    .iter()
                    .find(|f| f.to_lowercase() == folded)
                {
                    return Err(format!(
                        "{}: value field name '{}' is the column name '{}'; a caption \
                         must differ from every column name, case aside",
                        context, caption, column
                    ));
                }
                value_field = value_field.set_custom_name(caption);
            }
            if let Some(num_format) = inner.string("num_format")? {
                value_field = value_field.set_num_format(num_format);
            }
            value_field
        } else {
            return Err(format!(
                "{}: 'values' must hold column names or value field dicts, got {}",
                context,
                crate::types::pytype_name(&item)
            ));
        };
        table = table.add_value_field(&value_field);
    }

    // `style=None` turns the style off, as `table_style=None` does; a missing
    // key keeps Excel's default (Light16).
    match view.get("style") {
        None => {}
        Some(value) if value.bind(py).is_none() => table = table.set_style(PivotTableStyle::None),
        Some(_) => {
            let style = view.string("style")?.unwrap_or_default();
            table = table
                .set_style(parse_pivot_style(&style).map_err(|e| format!("{}: {}", context, e))?);
        }
    }
    if let Some(layout) = view.string("layout")? {
        table = table.set_layout(parse_layout(&layout).map_err(|e| format!("{}: {}", context, e))?);
    }
    if let Some(on) = view.bool("row_grand_totals")? {
        table = table.set_row_grand_totals(on);
    }
    if let Some(on) = view.bool("column_grand_totals")? {
        table = table.set_column_grand_totals(on);
    }
    Ok(table)
}

/// One pivot table, checked and built, waiting for its sheet.
struct Planned {
    /// `sheet '<frame's sheet>': pivot_tables['<key>']`, for every later error.
    context: String,
    sheet: String,
    row: u32,
    col: u16,
    table: PivotTable,
}

/// Every pivot table of a workbook, checked and built; nothing is added yet.
pub(crate) struct PivotPlan {
    /// Sheets to create, in the spelling of their first mention, each with the
    /// context of the key that first named it.
    new_sheets: Vec<(String, String)>,
    tables: Vec<Planned>,
}

impl PivotPlan {
    /// The sheets the caller must create, empty and in this order, before
    /// [`PivotPlan::add_to`]: `(name, context)`, the context naming the key.
    pub(crate) fn new_sheets(&self) -> &[(String, String)] {
        &self.new_sheets
    }

    /// Add every pivot table to its sheet. Each error names the key it came from.
    pub(crate) fn add_to(&self, workbook: &mut Workbook) -> Result<(), String> {
        for plan in &self.tables {
            let worksheet = workbook
                .worksheet_from_name(&plan.sheet)
                .map_err(|e| format!("{}: sheet '{}': {}", plan.context, plan.sheet, e))?;
            worksheet
                .add_pivot_table(plan.row, plan.col, &plan.table)
                .map_err(|e| format!("{}: {}", plan.context, e))?;
        }
        Ok(())
    }
}

/// Check, place and build every requested pivot table.
///
/// `areas` holds every frame in the workbook, with or without pivot tables, so
/// that an anchor inside any of them is refused. `sheets` is every sheet the
/// workbook has, in its own spelling. `writes` holds the cell values every
/// cell-addressed option writes, so that one in a header row is refused.
/// Nothing is created or added here: a refusal raises before the save, so no
/// file is written.
pub(crate) fn plan_pivot_tables(
    py: Python<'_>,
    sheets: &[&str],
    areas: &[&FrameArea],
    requests: &[PivotRequest<'_>],
    writes: &[CellWriteSpan],
) -> Result<PivotPlan, String> {
    let existing: HashMap<String, &str> = sheets.iter().map(|s| (s.to_lowercase(), *s)).collect();
    let mut new_sheets: IndexMap<String, (String, String)> = IndexMap::new();
    let mut anchors: HashMap<(String, u32, u16), String> = HashMap::new();
    let mut planned: Vec<Planned> = Vec::new();

    for request in requests {
        for (key, config) in request.tables {
            let context = format!("sheet '{}': pivot_tables['{}']", request.sheet, key);
            let location = parse_location(key).map_err(|e| format!("{}: {}", context, e))?;
            check_source(&context, request)?;
            let area = request
                .area
                .expect("check_source refuses a frame without an area");
            reject_header_writes(&context, area, writes)?;

            // The target sheet in the workbook's spelling; a new one is created later.
            let target = match &location.sheet {
                None => request.sheet.to_string(),
                Some(name) => {
                    let folded = name.to_lowercase();
                    match existing.get(&folded) {
                        Some(&spelled) => spelled.to_string(),
                        None => new_sheets
                            .entry(folded)
                            .or_insert_with(|| (name.clone(), context.clone()))
                            .0
                            .clone(),
                    }
                }
            };
            let folded_target = target.to_lowercase();
            let cell = rust_xlsxwriter::utility::row_col_to_cell(location.row, location.col);

            if let Some(frame) = areas.iter().find(|a| {
                a.sheet.to_lowercase() == folded_target && a.contains(location.row, location.col)
            }) {
                return Err(format!(
                    "{}: cell {} on sheet '{}' is inside the frame written at {}; \
                     place the pivot table outside every frame",
                    context,
                    cell,
                    frame.sheet,
                    frame.range_text()
                ));
            }
            let anchor = (folded_target, location.row, location.col);
            if let Some(previous) = anchors.insert(anchor, context.clone()) {
                return Err(format!(
                    "{}: cell {} on sheet '{}' already holds {}",
                    context, cell, target, previous
                ));
            }

            let table = build_pivot_table(py, &context, config, request, area)?;
            planned.push(Planned {
                context,
                sheet: target,
                row: location.row,
                col: location.col,
                table,
            });
        }
    }

    Ok(PivotPlan {
        new_sheets: new_sheets.into_values().collect(),
        tables: planned,
    })
}

#[cfg(test)]
mod tests {
    use super::{parse_layout, parse_location, parse_pivot_style};

    #[test]
    fn location_forms() {
        let plain = parse_location("F3").unwrap();
        assert_eq!((plain.sheet, plain.row, plain.col), (None, 2, 5));
        let other = parse_location("Summary!B4").unwrap();
        assert_eq!(
            (other.sheet.as_deref(), other.row, other.col),
            (Some("Summary"), 3, 1)
        );
        let quoted = parse_location("'It''s here'!A1").unwrap();
        assert_eq!(quoted.sheet.as_deref(), Some("It's here"));
        // Quoted, white space belongs to the name; the cell part is trimmed.
        let spaced = parse_location("' Data '!A3 ").unwrap();
        assert_eq!(
            (spaced.sheet.as_deref(), spaced.row, spaced.col),
            (Some(" Data "), 2, 0)
        );
    }

    #[test]
    fn bad_locations_are_refused() {
        for key in [
            "",
            "A0",
            "XFE1",
            "A1048577",
            "A1:B2",
            "Pivot!",
            "!A3",
            "''!A3",
            " Data !A3",
            "Data !A3",
            " Data!A3",
            " 'Data'!A3",
        ] {
            assert!(parse_location(key).is_err(), "{key:?}");
        }
    }

    #[test]
    fn every_gallery_style_parses_and_nothing_else() {
        for family in ["Light", "Medium", "Dark"] {
            for n in 1..=28 {
                assert!(
                    parse_pivot_style(&format!("{family}{n}")).is_ok(),
                    "{family}{n}"
                );
            }
            for bad in [0, 29] {
                assert!(
                    parse_pivot_style(&format!("{family}{bad}")).is_err(),
                    "{family}{bad}"
                );
            }
        }
        assert!(parse_pivot_style("None").is_ok());
        for bad in ["", "light16", "PivotStyleLight16", "TableStyleMedium9"] {
            assert!(parse_pivot_style(bad).is_err(), "{bad:?}");
        }
    }

    #[test]
    fn layouts() {
        for layout in ["compact", "outline", "tabular"] {
            assert!(parse_layout(layout).is_ok());
        }
        assert!(parse_layout("flat").is_err());
    }
}
