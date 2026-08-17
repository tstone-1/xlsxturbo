/// Excel's last row, 1-based. `parse_cell_ref` bounds the column at XFD but
/// deliberately not the row (a row past the grid is caught by the writer), so
/// the name checks below apply this bound themselves.
pub(crate) const MAX_ROW_1BASED: u32 = 1_048_576;

/// Excel's last column, 1-based (XFD).
const MAX_COLUMN_1BASED: u32 = 16_384;

/// True when `name` is a string Excel would read as a cell reference, and is
/// therefore not usable as a table name or a defined name.
///
/// Excel forbids a name that collides with a reference to a cell **that
/// exists**, and the grid bound is the whole subtlety: `"AAAA1"` and
/// `"A1048577"` name no cell (the grid stops at XFD1048576), so Excel accepts
/// both as ordinary names and refusing them would reject input Excel takes. The
/// check is therefore shape-then-grid — the A1 shape is confirmed by
/// [`parse_cell_ref`], which owns the XFD column bound, plus the row bound
/// applied here — rather than a letters-then-digits regex, which would over-reach
/// by exactly those cases.
///
/// Also covered are bare `"R"`/`"C"` in either case — reserved by Excel as
/// row/column selection shortcuts, not references themselves — and the R1C1
/// forms `R<row>C<col>` with either index optional (`"RC"`, `"R1C1"`, `"R1C"`,
/// `"RC1"`), again only within the grid.
///
/// Case-insensitive throughout (`"q1"` is as much a reference as `"Q1"`). A
/// zero-padded row is still a reference: a table named `"A01"` draws the same
/// recovery prompt from Excel as `"Q1"` (measured, Excel 16.112), so the digits
/// are judged by the row they parse to, not their spelling. Row 0 (`"A0"`)
/// parses to no row and stays a legal name.
pub(crate) fn looks_like_cell_reference(name: &str) -> bool {
    is_a1_cell_reference(name) || is_reserved_r1c1_name(name)
}

/// The `A1` half of [`looks_like_cell_reference`].
fn is_a1_cell_reference(name: &str) -> bool {
    let letters = name.chars().take_while(|c| c.is_ascii_alphabetic()).count();
    // Every counted char is ASCII, so the char count is also the byte offset.
    if letters == 0 || letters > 3 {
        return false;
    }
    let digits = &name[letters..];
    if digits.is_empty() || digits.len() > 7 || !digits.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }
    // Delegating the column bound keeps one definition of "which columns
    // exist"; `parse_cell_ref` errors past XFD.
    match parse_cell_ref(name) {
        Ok((row_0based, _)) => row_0based < MAX_ROW_1BASED,
        Err(_) => false,
    }
}

/// The R1C1 half of [`looks_like_cell_reference`].
fn is_reserved_r1c1_name(name: &str) -> bool {
    if name.eq_ignore_ascii_case("r") || name.eq_ignore_ascii_case("c") {
        return true;
    }

    let mut chars = name.chars();
    if !chars.next().is_some_and(|c| c.eq_ignore_ascii_case(&'r')) {
        return false;
    }
    let rest: &str = &name[1..];
    let row_digits: &str = &rest[..rest.chars().take_while(|c| c.is_ascii_digit()).count()];
    let after_row = &rest[row_digits.len()..];

    let mut after = after_row.chars();
    if !after.next().is_some_and(|c| c.eq_ignore_ascii_case(&'c')) {
        return false;
    }
    let col_digits = &after_row[1..];
    if !col_digits.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }

    // An absent index is the relative form ("RC", "R1C"), which is still a
    // reference. A present one only counts inside the grid, for the same reason
    // the A1 half is bounded.
    within(row_digits, MAX_ROW_1BASED) && within(col_digits, MAX_COLUMN_1BASED)
}

/// An empty index is the relative form and always qualifies; a present one must
/// address a row/column that exists.
fn within(digits: &str, max_1based: u32) -> bool {
    if digits.is_empty() {
        return true;
    }
    matches!(digits.parse::<u32>(), Ok(n) if n >= 1 && n <= max_1based)
}

/// Parse a cell reference like "A1" into (row, col) - 0-based
pub(crate) fn parse_cell_ref(cell_ref: &str) -> Result<(u32, u16), String> {
    let cell_ref = cell_ref.trim().to_uppercase();
    if cell_ref.is_empty() {
        return Err("Empty cell reference".to_string());
    }

    // Find where letters end and numbers begin
    let col_end = cell_ref
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .count();
    if col_end == 0 {
        return Err(format!(
            "Invalid cell reference '{}': no column letters",
            cell_ref
        ));
    }

    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    if row_str.is_empty() {
        return Err(format!(
            "Invalid cell reference '{}': no row number",
            cell_ref
        ));
    }

    // Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, etc.)
    // Use checked arithmetic to detect overflow on adversarial input
    let col_u32: u32 = col_str
        .chars()
        .try_fold(0u32, |acc, c| {
            acc.checked_mul(26)?.checked_add(c as u32 - 'A' as u32 + 1)
        })
        .ok_or_else(|| format!("Column '{}' is too large", col_str))?
        .saturating_sub(1);
    if col_u32 > 16383 {
        return Err(format!(
            "Column '{}' exceeds Excel's maximum column (XFD = 16384)",
            col_str
        ));
    }
    let col = col_u32 as u16;

    // Parse row number (Excel rows are 1-based, so must be >= 1)
    let row_1based: u32 = row_str
        .parse::<u32>()
        .map_err(|_| format!("Invalid row number in cell reference '{}'", cell_ref))?;

    if row_1based == 0 {
        return Err(format!(
            "Invalid cell reference '{}': row number must be >= 1 (Excel rows are 1-based)",
            cell_ref
        ));
    }

    // Convert to 0-based index
    let row = row_1based - 1;

    Ok((row, col))
}

/// Parse a cell range like "A1:D1" into (first_row, first_col, last_row, last_col) - 0-based
pub(crate) fn parse_cell_range(range_str: &str) -> Result<(u32, u16, u32, u16), String> {
    let parts: Vec<&str> = range_str.split(':').collect();
    if parts.len() != 2 {
        return Err(format!(
            "Invalid cell range '{}': expected format 'A1:B2'",
            range_str
        ));
    }

    let (first_row, first_col) = parse_cell_ref(parts[0])?;
    let (last_row, last_col) = parse_cell_ref(parts[1])?;

    if first_row > last_row || first_col > last_col {
        return Err(format!(
            "Invalid cell range '{}': first cell must precede the last cell (top-left to bottom-right)",
            range_str
        ));
    }

    Ok((first_row, first_col, last_row, last_col))
}

#[cfg(test)]
mod name_shape_tests {
    use super::looks_like_cell_reference;

    #[test]
    fn a1_shaped_names_inside_the_grid_are_references() {
        // "A01" is here because it was measured, not assumed: Excel 16.112
        // offers to repair a workbook whose table is named "A01" exactly as it
        // does for "Q1", so a zero-padded row counts as a reference.
        for name in ["A1", "Q1", "a1", "q1", "ABC123", "XFD1048576", "Z9", "A01"] {
            assert!(
                looks_like_cell_reference(name),
                "'{}' should read as a cell reference",
                name
            );
        }
    }

    #[test]
    fn shapes_beyond_the_grid_are_legal_names() {
        // The control that makes the bound mean something: each of these is
        // letters-then-digits and would be caught by a naive regex, but names
        // no cell, so Excel accepts it as an ordinary name.
        for name in [
            "AAAA1",     // 4 letters: past XFD
            "XFE1",      // one column past XFD
            "A1048577",  // one row past the last
            "A0",        // row 0 does not exist
            "A00",       // zero-padded row 0 still does not exist
            "A12345678", // 8 digits, past the row space
        ] {
            assert!(
                !looks_like_cell_reference(name),
                "'{}' names no cell and must stay a legal name",
                name
            );
        }
    }

    #[test]
    fn ordinary_names_are_not_references() {
        for name in ["Sales", "Q1_Sales", "_Q1", "Table1x", "R1C1D", "data2024x"] {
            assert!(
                !looks_like_cell_reference(name),
                "'{}' should be a legal name",
                name
            );
        }
    }

    #[test]
    fn reserved_r1c1_forms_are_references() {
        for name in ["R", "C", "r", "c", "RC", "R1C1", "r1c1", "R1C", "RC1"] {
            assert!(
                looks_like_cell_reference(name),
                "'{}' is reserved by Excel",
                name
            );
        }
    }

    #[test]
    fn r1c1_indices_outside_the_grid_are_legal_names() {
        for name in ["R1048577C1", "R1C16385", "R0C1", "R1C0"] {
            assert!(
                !looks_like_cell_reference(name),
                "'{}' addresses nothing and must stay a legal name",
                name
            );
        }
    }

    #[test]
    fn empty_and_whitespace_are_not_references() {
        // `parse_cell_ref` trims its input, so these say the shape check runs
        // on the raw string before delegating.
        for name in ["", " ", " A1", "A1 "] {
            assert!(!looks_like_cell_reference(name), "'{}' matched", name);
        }
    }
}
