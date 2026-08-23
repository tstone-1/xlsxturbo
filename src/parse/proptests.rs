//! Property tests over the parsers in this module.
//!
//! The unit tests in `mod.rs` pin specific inputs. These pin the *rules* those
//! inputs are examples of, over the whole input space proptest can reach. They
//! were chosen for the invariants that are cheap to state and expensive to get
//! wrong: a round-trip, an idempotence, an equivalence to a standard-library
//! predicate, and a documented range that the code and its own error message
//! must agree on.
//!
//! Two habits this file sticks to, both learned the hard way elsewhere in the
//! project:
//!
//! - **Both directions of a comparison.** A property saying "a prefix pattern
//!   matches a string that starts with the prefix" is satisfied by an
//!   implementation that matches *everything*. Each pattern property is stated
//!   as an equivalence against the `str` method it claims to implement, so it
//!   discriminates in both directions.
//! - **A totality check is not a property.** `..._never_panics` tests earn
//!   their place only because these parsers slice by byte offset on input that
//!   may be arbitrary UTF-8; they say nothing about correctness, and every one
//!   of them is paired with a property that does.

use super::tables::is_logical_constant;
use super::{
    looks_like_cell_reference, matches_pattern, naive_date_to_excel, naive_datetime_to_excel,
    parse_cell_range, parse_cell_ref, parse_color, parse_table_style, parse_value,
    sanitize_table_name,
};
use crate::types::{CellValue, DateOrder};
use proptest::prelude::*;

/// Excel's last row, 0-based (1,048,576 rows).
const MAX_ROW: u32 = 1_048_575;
/// Excel's last column, 0-based (XFD).
const MAX_COL: u32 = 16_383;

/// Encode a 0-based column index as Excel's bijective base-26 column letters.
///
/// `parse_cell_ref` has no inverse in `src/` because nothing in the library
/// needs one, so the round-trip property below needs one here. Writing it in
/// the test buys independence -- a bug in `parse_cell_ref` will not be mirrored
/// by an identical bug here -- but it does not buy immunity from a *shared*
/// misunderstanding of the encoding. `column_letters_are_anchored` pins the four
/// values nobody disputes, so this function cannot quietly drift into agreeing
/// with a broken parser.
fn column_letters(col: u32) -> String {
    let mut out = Vec::new();
    let mut n = col;
    loop {
        out.push(b'A' + (n % 26) as u8);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).expect("only ASCII letters are pushed")
}

#[test]
fn column_letters_are_anchored() {
    assert_eq!(column_letters(0), "A");
    assert_eq!(column_letters(25), "Z");
    assert_eq!(column_letters(26), "AA");
    assert_eq!(column_letters(MAX_COL), "XFD");
}

/// The named colors `parse_color` accepts, and the value each must produce.
///
/// This is a second copy of the table in `colors.rs`, which is normally the
/// thing to avoid -- but here the duplication *is* the test. A property
/// generated from the implementation's own list could only prove the list
/// agrees with itself.
const NAMED_COLORS: &[(&str, u32)] = &[
    ("white", 0xFFFFFF),
    ("black", 0x000000),
    ("red", 0xFF0000),
    ("green", 0x00FF00),
    ("blue", 0x0000FF),
    ("yellow", 0xFFFF00),
    ("cyan", 0x00FFFF),
    ("magenta", 0xFF00FF),
    ("gray", 0x808080),
    ("grey", 0x808080),
    ("silver", 0xC0C0C0),
    ("orange", 0xFFA500),
    ("purple", 0x800080),
    ("navy", 0x000080),
    ("teal", 0x008080),
    ("maroon", 0x800000),
];

/// The three table-style families and the highest index each one goes up to,
/// exactly as `parse_table_style`'s error message advertises them.
const TABLE_STYLE_FAMILIES: &[(&str, u32)] = &[("Light", 21), ("Medium", 28), ("Dark", 11)];

/// The first date Excel numbers correctly, and its serial.
///
/// Everything before 1900-03-01 sits on the wrong side of Excel's phantom
/// 1900-02-29, which is why `parse_value` refuses to emit a date below 61.
const FIRST_SANE_SERIAL: f64 = 61.0;

proptest! {
    // --- Cell references -------------------------------------------------

    /// Every addressable cell survives being written as A1 and read back.
    ///
    /// Meant to catch: the `saturating_sub(1)` that converts bijective base-26
    /// to a 0-based index, and the `row_1based - 1` beside it. Either off by
    /// one and this fails on the first case.
    #[test]
    fn cell_ref_round_trips(row in 0u32..=MAX_ROW, col in 0u32..=MAX_COL) {
        let a1 = format!("{}{}", column_letters(col), row + 1);
        prop_assert_eq!(parse_cell_ref(&a1), Ok((row, col as u16)));
    }

    /// Case and surrounding whitespace are not part of the address.
    #[test]
    fn cell_ref_ignores_case_and_padding(
        row in 0u32..=MAX_ROW,
        col in 0u32..=MAX_COL,
        pad in "[ \t]{0,3}",
    ) {
        let a1 = format!("{}{}{}{}", pad, column_letters(col).to_lowercase(), row + 1, pad);
        prop_assert_eq!(parse_cell_ref(&a1), Ok((row, col as u16)));
    }

    /// A column past XFD is refused rather than silently truncated to `u16`.
    ///
    /// Meant to catch: widening or deleting the 16,383 bound. Without it the
    /// `as u16` cast below would wrap, and column 16,384 would land on column 0.
    #[test]
    fn cell_ref_rejects_columns_past_xfd(col in (MAX_COL + 1)..1_000_000u32) {
        let a1 = format!("{}1", column_letters(col));
        prop_assert!(parse_cell_ref(&a1).is_err());
    }

    /// Row 0 does not exist: Excel rows are 1-based.
    #[test]
    fn cell_ref_rejects_row_zero(col in 0u32..=MAX_COL, zeros in "0{1,4}") {
        let a1 = format!("{}{}", column_letters(col), zeros);
        prop_assert!(parse_cell_ref(&a1).is_err());
    }

    /// Arbitrary input produces an `Err`, never a panic.
    ///
    /// This one is not decoration: `parse_cell_ref` uppercases its input and
    /// then slices it at a byte offset counted in `char`s. That is only sound
    /// because the prefix it counts is ASCII, and `to_uppercase` can change a
    /// string's byte length (`ß` becomes `SS`) and can turn a non-ASCII char
    /// into an ASCII one (`ı` becomes `I`).
    #[test]
    fn cell_ref_never_panics(s in ".*") {
        let _ = parse_cell_ref(&s);
    }

    // --- Cell ranges -----------------------------------------------------

    /// A range built from two well-ordered corners parses back to those corners.
    #[test]
    fn cell_range_round_trips(
        row in 0u32..=MAX_ROW,
        col in 0u32..=MAX_COL,
        dr in 0u32..1_000,
        dc in 0u32..1_000,
    ) {
        let (last_row, last_col) = ((row + dr).min(MAX_ROW), (col + dc).min(MAX_COL));
        let range = format!(
            "{}{}:{}{}",
            column_letters(col), row + 1,
            column_letters(last_col), last_row + 1,
        );
        prop_assert_eq!(
            parse_cell_range(&range),
            Ok((row, col as u16, last_row, last_col as u16))
        );
    }

    /// Corners given bottom-right first are refused, not silently swapped.
    ///
    /// Silently normalising would be the tempting fix and the wrong one: a
    /// reversed range is far more often a typo than an intent.
    #[test]
    fn cell_range_rejects_reversed_corners(
        row in 0u32..=MAX_ROW,
        col in 0u32..=MAX_COL,
        dr in 0u32..1_000,
        dc in 0u32..1_000,
    ) {
        prop_assume!(dr > 0 || dc > 0);
        let (last_row, last_col) = ((row + dr).min(MAX_ROW), (col + dc).min(MAX_COL));
        prop_assume!(last_row > row || last_col > col);
        let reversed = format!(
            "{}{}:{}{}",
            column_letters(last_col), last_row + 1,
            column_letters(col), row + 1,
        );
        prop_assert!(parse_cell_range(&reversed).is_err());
    }

    /// Anything that is not exactly two colon-separated cells is refused.
    #[test]
    fn cell_range_requires_exactly_one_colon(s in "[A-Z0-9:]{0,12}") {
        prop_assume!(s.matches(':').count() != 1);
        prop_assert!(parse_cell_range(&s).is_err());
    }

    // --- Colors ----------------------------------------------------------

    /// Every 24-bit color survives `#RRGGBB` in either case.
    #[test]
    fn hex_color_round_trips(rgb in 0u32..0x0100_0000) {
        prop_assert_eq!(parse_color(&format!("#{:06X}", rgb)), Ok(rgb));
        prop_assert_eq!(parse_color(&format!("#{:06x}", rgb)), Ok(rgb));
    }

    /// Surrounding whitespace is trimmed before parsing.
    #[test]
    fn hex_color_ignores_padding(rgb in 0u32..0x0100_0000, pad in "[ \t\r\n]{0,4}") {
        prop_assert_eq!(parse_color(&format!("{}#{:06X}{}", pad, rgb, pad)), Ok(rgb));
    }

    /// Six characters after `#` parse **iff** all six are hex digits.
    ///
    /// The alphabet is deliberately narrow, and that is the whole design of
    /// this property. Written first over all printable ASCII, it passed under
    /// the mutation it exists to catch: the interesting inputs are a ~1 in
    /// 70,000 slice of that space, so 256 cases never reached one. A generator
    /// that cannot produce the failing case makes a property indistinguishable
    /// from a test with no assertion. Hex digits plus the two sign characters
    /// puts a near-miss in roughly one case in nine.
    #[test]
    fn six_chars_are_a_color_iff_all_hex(s in "[0-9a-fA-F+\\-]{6}") {
        let all_hex = s.chars().all(|c| c.is_ascii_hexdigit());
        prop_assert_eq!(parse_color(&format!("#{}", s)).is_ok(), all_hex);
    }

    /// A signed hex literal is not a color, however well-formed it looks.
    ///
    /// `u32::from_str_radix` accepts a leading `+`, so `#+12345` passes a
    /// length check and parses as 0x12345. `colors.rs` carries an explicit
    /// `is_ascii_hexdigit` guard for exactly that. Here the generator *is* the
    /// case rather than a region it might wander into, so this cannot go quiet
    /// the way the property above did.
    #[test]
    fn sign_prefixed_hex_is_not_a_color(sign in "[+\\-]", rest in "[0-9a-fA-F]{5}") {
        let candidate = format!("#{}{}", sign, rest);
        prop_assert!(
            parse_color(&candidate).is_err(),
            "'{}' was accepted as a color", candidate
        );
    }

    /// A hex color is accepted **iff** it is exactly six digits long.
    #[test]
    fn hex_color_requires_exactly_six_digits(hex in "[0-9a-fA-F]{0,20}") {
        prop_assert_eq!(parse_color(&format!("#{}", hex)).is_ok(), hex.len() == 6);
    }

    /// Named colors are case-insensitive and produce their documented value.
    #[test]
    fn named_colors_ignore_case(
        idx in 0usize..NAMED_COLORS.len(),
        mask in any::<u32>(),
        pad in "[ \t]{0,3}",
    ) {
        let (name, expected) = NAMED_COLORS[idx];
        let cased: String = name
            .chars()
            .enumerate()
            .map(|(i, c)| if (mask >> (i % 32)) & 1 == 1 { c.to_ascii_uppercase() } else { c })
            .collect();
        prop_assert_eq!(parse_color(&format!("{}{}{}", pad, cased, pad)), Ok(expected));
    }

    /// Arbitrary input produces an `Err`, never a panic.
    #[test]
    fn parse_color_never_panics(s in ".*") {
        let _ = parse_color(&s);
    }

    // --- Table styles ----------------------------------------------------

    /// Each family accepts exactly the range its own error message advertises.
    ///
    /// The message promises "Light1-Light21, Medium1-Medium28, Dark1-Dark11".
    /// That is a claim about the code, and until now nothing checked it. Meant
    /// to catch drift in either direction: a variant dropped from the match, or
    /// a message that over-promises.
    #[test]
    fn table_style_families_match_their_documented_ranges(
        idx in 0usize..TABLE_STYLE_FAMILIES.len(),
        n in 1u32..40,
    ) {
        let (family, highest) = TABLE_STYLE_FAMILIES[idx];
        prop_assert_eq!(
            parse_table_style(&format!("{}{}", family, n)).is_ok(),
            n <= highest
        );
    }

    /// The lookup is case-sensitive; `medium9` is not `Medium9`.
    #[test]
    fn table_style_lookup_is_case_sensitive(idx in 0usize..TABLE_STYLE_FAMILIES.len(), n in 1u32..12) {
        let (family, _) = TABLE_STYLE_FAMILIES[idx];
        let lowercased = format!("{}{}", family.to_lowercase(), n);
        prop_assert!(
            parse_table_style(&lowercased).is_err(),
            "'{}' was accepted", lowercased
        );
    }

    // --- Table names -----------------------------------------------------

    /// Sanitising produces something that satisfies every rule the function
    /// claims to enforce, for any input at all.
    #[test]
    fn sanitized_table_names_are_always_valid(name in ".*") {
        let sanitized = sanitize_table_name(&name);
        prop_assert!(!sanitized.is_empty());
        prop_assert!(sanitized.chars().all(|c| c.is_alphanumeric() || c == '_'));
        prop_assert!(!sanitized.chars().next().expect("non-empty").is_ascii_digit());
        prop_assert!(sanitized.chars().count() <= 255);
    }

    /// The same rules, over names long enough to reach the 255-character cap.
    ///
    /// This exists because the property above does not test what it appears
    /// to. `".*"` generates short strings, so the truncation branch was never
    /// entered and swapping the truncate and prepend steps -- which pushes a
    /// 255-character digit-leading name to 256 -- survived it untouched. The
    /// generator here starts with a digit *and* straddles the cap, so both the
    /// prepend and the truncation happen on every case.
    #[test]
    fn long_table_names_are_capped_after_the_prefix_is_added(
        // At least 256 characters, so adding the prefix always overshoots the
        // cap and the truncation branch is entered on every case.
        name in "[0-9][A-Za-z0-9_]{255,300}",
    ) {
        let sanitized = sanitize_table_name(&name);
        prop_assert_eq!(sanitized.chars().count(), 255);
        prop_assert!(sanitized.starts_with('_'));
    }

    /// A multibyte character sitting on the 255-character boundary is not a
    /// panic.
    ///
    /// The cap counts `char`s precisely so this cannot become a byte-boundary
    /// slice, and this is the test that says so.
    #[test]
    fn table_name_truncation_is_char_safe(reps in 250usize..300) {
        let name: String = "é".repeat(reps);
        let sanitized = sanitize_table_name(&name);
        prop_assert!(sanitized.chars().count() <= 255);
    }

    /// Sanitising an already-sanitised name changes nothing.
    ///
    /// A name that survives one pass and is altered by a second means the
    /// function's own output is not in its accepted set -- which is how a
    /// "sanitised" name still ends up rejected by Excel.
    #[test]
    fn sanitizing_a_table_name_is_idempotent(name in ".*") {
        let once = sanitize_table_name(&name);
        prop_assert_eq!(sanitize_table_name(&once), once);
    }

    /// Names that are already valid pass through untouched.
    ///
    /// The generator reaches cell-reference shapes ("A1", "R", "Q1") often —
    /// measured at roughly one case in 200 — and those are *not* already valid:
    /// they collide with a cell address, so sanitising prefixes them. Excluding
    /// them here rather than narrowing the regex keeps the generator wide, and
    /// the reference-shape behaviour has its own assertions in `tables.rs`.
    ///
    /// The logical constants are excluded for the same reason and are worth a
    /// word, because this property was *false* for a while without anyone
    /// noticing: "TRUE" matches the regex, is not a cell reference, and is
    /// prefixed — but the odds of `".*"`-style generation spelling it are so
    /// low that the suite would have stayed green indefinitely. Both exclusions
    /// call the real predicate rather than restating it, so a change to either
    /// rule moves this property with it.
    #[test]
    fn valid_table_names_are_left_alone(name in "[A-Za-z_][A-Za-z0-9_]{0,40}") {
        prop_assume!(!looks_like_cell_reference(&name));
        prop_assume!(!is_logical_constant(&name));
        prop_assert_eq!(sanitize_table_name(&name), name);
    }

    /// A sanitised name never reads as a cell reference, for any input.
    ///
    /// The property the prefix exists for, stated over the whole input space
    /// rather than the handful of named cases in `tables.rs`.
    #[test]
    fn sanitized_table_names_are_never_cell_references(name in "[A-Za-z_0-9]{0,12}") {
        prop_assert!(!looks_like_cell_reference(&sanitize_table_name(&name)));
    }

    // --- Column patterns -------------------------------------------------

    /// `prefix*` is exactly `str::starts_with`.
    #[test]
    fn prefix_pattern_is_starts_with(name in "[^*]{0,10}", prefix in "[^*]{0,6}") {
        prop_assert_eq!(
            matches_pattern(&name, &format!("{}*", prefix)),
            name.starts_with(&prefix)
        );
    }

    /// `*suffix` is exactly `str::ends_with`.
    #[test]
    fn suffix_pattern_is_ends_with(name in "[^*]{0,10}", suffix in "[^*]{0,6}") {
        prop_assert_eq!(
            matches_pattern(&name, &format!("*{}", suffix)),
            name.ends_with(&suffix)
        );
    }

    /// `*inner*` is exactly `str::contains`.
    #[test]
    fn contains_pattern_is_contains(name in "[^*]{0,10}", inner in "[^*]{1,6}") {
        prop_assert_eq!(
            matches_pattern(&name, &format!("*{}*", inner)),
            name.contains(&inner)
        );
    }

    /// A pattern with no `*` is exactly equality.
    ///
    /// Stated as an equivalence so it discriminates in both directions: a
    /// mutation from `==` to `starts_with` is caught by a name that *extends*
    /// the pattern, which a one-directional test would miss entirely.
    #[test]
    fn exact_pattern_is_equality(name in "[^*]{0,10}", pattern in "[^*]{0,10}") {
        prop_assert_eq!(matches_pattern(&name, &pattern), name == pattern);
    }

    /// A bare `*` matches everything.
    #[test]
    fn a_lone_star_matches_anything(name in ".*") {
        prop_assert!(matches_pattern(&name, "*"));
        prop_assert!(matches_pattern(&name, "**"));
    }

    /// Arbitrary input produces a verdict, never a panic.
    ///
    /// `matches_pattern` slices its pattern at byte offsets 1 and `len - 1`.
    /// Those are only guaranteed to be character boundaries because the
    /// characters being stripped are ASCII `*`; this is the test that says so.
    #[test]
    fn matches_pattern_never_panics(name in ".*", pattern in ".*") {
        let _ = matches_pattern(&name, &pattern);
    }

    // --- Value detection -------------------------------------------------

    /// Every `i64` is detected as an integer, with its value intact.
    #[test]
    fn integers_round_trip(v in any::<i64>()) {
        let parsed = parse_value(&v.to_string(), DateOrder::Auto);
        prop_assert!(
            matches!(parsed, CellValue::Integer(got) if got == v),
            "expected Integer({}), got {:?}", v, parsed
        );
    }

    /// Every finite non-integral `f64` is detected as a float, bit for bit.
    #[test]
    fn finite_floats_round_trip(v in any::<f64>()) {
        prop_assume!(v.is_finite());
        let text = format!("{:?}", v);
        // A value whose shortest representation is integral is an Integer by
        // design -- `parse_value` tries `i64` first -- so exclude those here.
        prop_assume!(text.parse::<i64>().is_err());
        let parsed = parse_value(&text, DateOrder::Auto);
        prop_assert!(
            matches!(parsed, CellValue::Float(got) if got == v),
            "expected Float({}), got {:?}", v, parsed
        );
    }

    /// Whitespace-only input is an empty cell, not a string of spaces.
    #[test]
    fn whitespace_only_is_empty(s in "[ \t\r\n]{0,8}") {
        prop_assert!(matches!(parse_value(&s, DateOrder::Auto), CellValue::Empty));
    }

    /// A value detected as a string keeps its original padding.
    ///
    /// Trimming is a detection aid only; the stored value must be what the
    /// caller supplied, or round-tripping a CSV silently rewrites its data.
    ///
    /// The `prop_assume!` is not a convenience. Written without it, this
    /// property failed on `"\tNan\t"` -- because `"Nan"` is a perfectly good
    /// `f64` literal, and a NaN is deliberately written as an empty cell. The
    /// assumption is what makes "a word, not a number" the actual population,
    /// and `float_literal_words_are_emptied_not_stored` below pins the case it
    /// excludes rather than letting it fall out of coverage.
    #[test]
    fn strings_keep_their_original_padding(core in "[a-zA-Z]{1,8}", pad in "[ \t]{1,4}") {
        prop_assume!(core.parse::<f64>().is_err());
        let padded = format!("{}{}{}", pad, core, pad);
        let parsed = parse_value(&padded, DateOrder::Auto);
        prop_assert!(
            matches!(&parsed, CellValue::String(got) if got == &padded),
            "expected the padding preserved, got {:?}", parsed
        );
    }

    /// Every spelling Rust accepts for a non-finite float becomes an empty cell.
    ///
    /// Documented behaviour (`docs/dataframe-export.md`: `NaN` -> Empty), and
    /// the *only* unit test for it used the single spelling `"NaN"`. Rust's
    /// `f64` parser accepts a dozen more -- `inf`, `Infinity`, `+nan`, `-INF`
    /// -- and each one silently empties a cell, which is worth knowing when a
    /// text column happens to contain one.
    #[test]
    fn float_literal_words_are_emptied_not_stored(
        sign in "[+\\-]?",
        word in prop::sample::select(vec!["nan", "NaN", "NAN", "inf", "Inf", "INF", "infinity", "Infinity"]),
        pad in "[ \t]{0,3}",
    ) {
        let text = format!("{}{}{}{}", pad, sign, word, pad);
        prop_assert!(
            matches!(parse_value(&text, DateOrder::Auto), CellValue::Empty),
            "{:?} should be an empty cell", text
        );
    }

    /// Arbitrary input produces a value, never a panic.
    #[test]
    fn parse_value_never_panics(s in ".*", order in prop::sample::select(
        vec![DateOrder::Auto, DateOrder::DMY, DateOrder::MDY]
    )) {
        let _ = parse_value(&s, order);
    }

    // --- Excel serial dates ----------------------------------------------

    /// The serial scale is linear, and anchored at 1900-03-01 = 61.
    ///
    /// One anchor plus linearity pins the whole scale. Meant to catch a shifted
    /// epoch, which a single spot-check date would also catch -- and a
    /// *non-uniform* shift, which it would not.
    #[test]
    fn excel_serials_are_linear_from_a_known_anchor(days in 0i64..80_000) {
        let anchor = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).expect("valid date");
        let date = anchor + chrono::Duration::days(days);
        prop_assert_eq!(naive_date_to_excel(date), FIRST_SANE_SERIAL + days as f64);
    }

    /// An ISO date on or after the anchor is detected as a date, with the
    /// serial `naive_date_to_excel` would give it.
    #[test]
    fn iso_dates_agree_with_the_serial_conversion(days in 0i64..80_000) {
        let anchor = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).expect("valid date");
        let date = anchor + chrono::Duration::days(days);
        let expected = naive_date_to_excel(date);
        let parsed = parse_value(&date.format("%Y-%m-%d").to_string(), DateOrder::Auto);
        prop_assert!(
            matches!(parsed, CellValue::Date(got) if got == expected),
            "expected Date({}), got {:?}", expected, parsed
        );
    }

    /// A datetime serial splits cleanly: whole part the date, fraction the time.
    #[test]
    fn datetime_serials_split_into_date_and_time(days in 0i64..80_000, secs in 0u32..86_400) {
        let anchor = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).expect("valid date");
        let date = anchor + chrono::Duration::days(days);
        let time = chrono::NaiveTime::from_num_seconds_from_midnight_opt(secs, 0)
            .expect("seconds are within a day");
        let serial = naive_datetime_to_excel(date.and_time(time));
        prop_assert_eq!(serial.floor(), naive_date_to_excel(date));
        prop_assert!((serial.fract() - secs as f64 / 86_400.0).abs() < 1e-9);
    }

    /// No date Excel cannot represent is ever emitted as a date.
    ///
    /// Everything in the 1900 leap-year gap must fall back to a string. This is
    /// the property the two `excel_dt < 61.0` guards in `values.rs` exist for.
    #[test]
    fn dates_in_the_1900_gap_fall_back_to_string(day in 1u32..=59) {
        let date = chrono::NaiveDate::from_ymd_opt(1900, 1, 1).expect("valid date")
            + chrono::Duration::days(i64::from(day) - 1);
        let text = date.format("%Y-%m-%d").to_string();
        let parsed = parse_value(&text, DateOrder::Auto);
        prop_assert!(
            matches!(&parsed, CellValue::String(got) if got == &text),
            "expected the raw string back for {}, got {:?}", text, parsed
        );
    }

    /// Whatever a date parses to, it is never a serial Excel renders wrongly.
    ///
    /// Broader than the property above and stated over the output rather than
    /// the input: any `Date`/`DateTime` this function emits is at or past the
    /// first correctly-numbered day, whatever the input looked like.
    #[test]
    fn no_emitted_date_is_below_the_first_sane_serial(s in "[0-9]{1,4}[-/][0-9]{1,2}[-/][0-9]{1,4}") {
        match parse_value(&s, DateOrder::Auto) {
            CellValue::Date(serial) | CellValue::DateTime(serial) => {
                prop_assert!(serial >= FIRST_SANE_SERIAL, "emitted serial {}", serial);
            }
            _ => {}
        }
    }
}
