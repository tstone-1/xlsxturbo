use super::cell_refs::looks_like_cell_reference;
use rust_xlsxwriter::TableStyle;
use unicode_normalization::UnicodeNormalization;

/// Generate a table style lookup match from a list of (string, variant) pairs.
macro_rules! table_style_match {
    ($style:expr, $( $name:literal => $variant:ident ),+ $(,)?) => {
        match $style {
            $( $name => Ok(TableStyle::$variant), )+
            _ => Err(format!(
                "Unknown table_style '{}'. Valid styles: Light1-Light21, Medium1-Medium28, Dark1-Dark11, None",
                $style
            )),
        }
    };
}

/// Parse a table style string into a `TableStyle` enum value.
/// Synced with rust_xlsxwriter TableStyle variants.
pub(crate) fn parse_table_style(style: &str) -> Result<TableStyle, String> {
    table_style_match!(style,
        "None" => None,
        "Light1" => Light1, "Light2" => Light2, "Light3" => Light3, "Light4" => Light4,
        "Light5" => Light5, "Light6" => Light6, "Light7" => Light7, "Light8" => Light8,
        "Light9" => Light9, "Light10" => Light10, "Light11" => Light11, "Light12" => Light12,
        "Light13" => Light13, "Light14" => Light14, "Light15" => Light15, "Light16" => Light16,
        "Light17" => Light17, "Light18" => Light18, "Light19" => Light19, "Light20" => Light20,
        "Light21" => Light21,
        "Medium1" => Medium1, "Medium2" => Medium2, "Medium3" => Medium3, "Medium4" => Medium4,
        "Medium5" => Medium5, "Medium6" => Medium6, "Medium7" => Medium7, "Medium8" => Medium8,
        "Medium9" => Medium9, "Medium10" => Medium10, "Medium11" => Medium11, "Medium12" => Medium12,
        "Medium13" => Medium13, "Medium14" => Medium14, "Medium15" => Medium15, "Medium16" => Medium16,
        "Medium17" => Medium17, "Medium18" => Medium18, "Medium19" => Medium19, "Medium20" => Medium20,
        "Medium21" => Medium21, "Medium22" => Medium22, "Medium23" => Medium23, "Medium24" => Medium24,
        "Medium25" => Medium25, "Medium26" => Medium26, "Medium27" => Medium27, "Medium28" => Medium28,
        "Dark1" => Dark1, "Dark2" => Dark2, "Dark3" => Dark3, "Dark4" => Dark4,
        "Dark5" => Dark5, "Dark6" => Dark6, "Dark7" => Dark7, "Dark8" => Dark8,
        "Dark9" => Dark9, "Dark10" => Dark10, "Dark11" => Dark11,
    )
}

/// True for `TRUE` and `FALSE` in any case.
///
/// Excel reserves its two logical constants, and a table named `"true"` draws
/// the same recovery prompt as one named `"Q1"` (measured on Excel 16.112,
/// rust_xlsxwriter#189). Folded with `to_uppercase`, which is the fold
/// rust_xlsxwriter applies, so the two cannot disagree about a name and leave a
/// rewritten one still being refused at `add_table`.
pub(crate) fn is_logical_constant(name: &str) -> bool {
    let upper = name.to_uppercase();
    upper == "TRUE" || upper == "FALSE"
}

/// Sanitize a string for use as an Excel table name
///
/// The name is normalised to NFC first, then invalid characters become `_` and
/// a leading digit gains a `_` prefix. A name Excel would read as a cell
/// reference (`"Q1"`, `"A1"`, `"XFD1048576"`, `"R2D2"`), an R1C1 form
/// (`"R1C1"`) or a reserved selection shortcut (`"R"`, `"C"`) gets the same `_`
/// prefix — `"Q1"` becomes `"_Q1"`, which addresses no cell. So does one of
/// Excel's logical constants, `"TRUE"` or `"FALSE"` in any case. See
/// [`looks_like_cell_reference`] for why the beyond-the-grid shapes
/// (`"AAAA1"`) are deliberately left alone.
///
/// # Why here, when rust_xlsxwriter also refuses these
///
/// It refuses them: 0.99.0 validates both rules and `add_table` returns
/// `NameError` (rust_xlsxwriter#189; before that the name was stored verbatim
/// and Excel offered to repair the workbook). Rewriting still happens here
/// because refusing is the wrong answer for a table name — nothing references
/// it, `table_name` has always fixed up invalid characters and a leading digit
/// silently, and a hard error for a name we can repair would break callers that
/// pass a column heading straight through. The screen has to stay at least as
/// wide as the crate's, or a name we let past becomes an error the caller
/// cannot act on.
///
/// # Why NFC first
///
/// The character screen below is an allowlist of `is_alphanumeric() || '_'`,
/// and a combining mark is neither, so decomposed text was silently rewritten:
/// `"Verkäufe"` typed as `"Verka" + U+0308` came out as `"Verka_ufe"`, and
/// `"がくせい"` in NFD as `"か_くせい"`. Excel accepts both spellings, so that
/// was data loss with no warning. Composing first repairs every name whose
/// marks have a precomposed form.
///
/// It does not repair the rest, and that is a known limit rather than an
/// oversight: `"ไม่"` (Thai tone mark U+0E48) and `"हिन्दी"` (Hindi virama
/// U+094D) have no composed form, so they are still rewritten even though
/// Excel accepts them. Closing that gap means widening the allowlist or
/// inverting it to a denylist, which changes what is *accepted* and needs its
/// own audit in that direction — see `AGENTS.md`.
pub(crate) fn sanitize_table_name(name: &str) -> String {
    // Compose before screening, so a decomposed name is not mangled by the
    // allowlist below. NFC and not NFKC: NFKC would fold U+FF21 FULLWIDTH A to
    // ASCII "A", turning names Excel treats as distinct into the same name.
    let composed: String = name.nfc().collect();
    let mut sanitized: String = composed
        .chars()
        .map(|c| {
            if c.is_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect();

    // Must start with letter or underscore
    if sanitized.chars().next().is_none_or(|c| c.is_ascii_digit()) {
        sanitized = format!("_{}", sanitized);
    }

    // Must not collide with a cell reference or a logical constant. Done before
    // the cap so the 255-character invariant holds unconditionally: the prefix
    // adds one character, and truncating afterwards cannot reintroduce either
    // shape, because what survives still starts with the `_`.
    if looks_like_cell_reference(&sanitized) || is_logical_constant(&sanitized) {
        sanitized = format!("_{}", sanitized);
    }

    // Max 255 chars. Build by chars (not bytes) so a multibyte codepoint at the
    // boundary can never trigger a `truncate` mid-char-boundary panic.
    if sanitized.chars().count() > 255 {
        sanitized = sanitized.chars().take(255).collect();
    }
    sanitized
}

/// Excel's limit for a table column name, in UTF-16 code units.
const TABLE_HEADER_MAX_UTF16: usize = 255;

/// Whether `name` holds an `_xHHHH_` sequence: `_x`, four hex digits, `_`.
///
/// Excel reads that shape as an escaped character. The `x` is lower case only
/// (`_X000D_` opens), the hex digits may be either case.
fn has_escape_shape(name: &str) -> bool {
    name.as_bytes().windows(7).any(|w| {
        w[0] == b'_' && w[1] == b'x' && w[6] == b'_' && w[2..6].iter().all(u8::is_ascii_hexdigit)
    })
}

/// Why `name` cannot be the header of a table column, or `None` when it can.
///
/// rust_xlsxwriter 0.102.0 writes the name into `<tableColumn name="...">`
/// without the `_xHHHH_` escaping Excel uses there, so each of these headers
/// gives a workbook Excel refuses to open, with `Ok` from the crate. Measured
/// against Excel with the crate alone on 2026-11-13; the cases and their
/// controls are in `dev-docs/upstream-rust-xlsxwriter.md`. A line feed is not
/// in the list: the crate writes it as `&#xA;` and Excel opens that.
///
/// The same header is fine without a table, which is why this is a check on
/// the table and not on the column name. Remove a rule here only when a crate
/// release writes that header in a form Excel opens.
pub(crate) fn unwritable_table_header(name: &str) -> Option<String> {
    if let Some(c) = name
        .chars()
        .find(|&c| (c < '\u{20}' && c != '\n') || c == '\u{fffe}' || c == '\u{ffff}')
    {
        let what = match c {
            '\t' => "a tab".to_string(),
            '\r' => "a carriage return".to_string(),
            other => format!("the control character U+{:04X}", other as u32),
        };
        return Some(format!("it contains {what}"));
    }
    if has_escape_shape(name) {
        return Some(
            "it contains text of the form '_xHHHH_', which Excel reads as an escaped character"
                .to_string(),
        );
    }
    let units = name.encode_utf16().count();
    if units > TABLE_HEADER_MAX_UTF16 {
        return Some(format!(
            "it is {units} characters long and Excel allows {TABLE_HEADER_MAX_UTF16}"
        ));
    }
    None
}

#[cfg(test)]
mod table_header_tests {
    use super::unwritable_table_header;

    #[test]
    fn headers_excel_refuses_in_a_table_are_named() {
        for (header, reason) in [
            ("a\tb", "a tab"),
            ("a\rb", "a carriage return"),
            ("a\r\nb", "a carriage return"),
            ("a\u{1}b", "U+0001"),
            ("a\u{1f}b", "U+001F"),
            ("a\u{0}b", "U+0000"),
            ("a\u{fffe}b", "U+FFFE"),
            ("a\u{ffff}b", "U+FFFF"),
            ("_x000D_", "_xHHHH_"),
            ("_x000d_", "_xHHHH_"),
            ("a_x0041_b", "_xHHHH_"),
        ] {
            let found = unwritable_table_header(header);
            assert!(
                found.as_deref().is_some_and(|r| r.contains(reason)),
                "{header:?}: {found:?}"
            );
        }
    }

    #[test]
    fn headers_excel_opens_in_a_table_pass() {
        // Each of these was opened in Excel as a table header.
        for header in [
            "a", "a\nb", "  a  ", "a\u{7f}b", "a\u{85}b", "_X000D_", "_x00G0_", "_x000D",
            "_x0000D_",
        ] {
            assert_eq!(unwritable_table_header(header), None, "{header:?}");
        }
    }

    #[test]
    fn the_length_limit_counts_utf16_units() {
        assert_eq!(unwritable_table_header(&"x".repeat(255)), None);
        assert!(unwritable_table_header(&"x".repeat(256)).is_some());
        // One emoji is two UTF-16 units: 127 fit, 128 do not.
        assert_eq!(unwritable_table_header(&"\u{1F600}".repeat(127)), None);
        let reason = unwritable_table_header(&"\u{1F600}".repeat(128));
        assert!(reason.is_some_and(|r| r.contains("256")));
    }
}

#[cfg(test)]
mod cell_reference_name_tests {
    use super::sanitize_table_name;

    #[test]
    fn reference_shaped_names_gain_an_underscore() {
        assert_eq!(sanitize_table_name("Q1"), "_Q1");
        assert_eq!(sanitize_table_name("A1"), "_A1");
        assert_eq!(sanitize_table_name("abc123"), "_abc123");
        assert_eq!(sanitize_table_name("XFD1048576"), "_XFD1048576");
    }

    #[test]
    fn reserved_r1c1_names_gain_an_underscore() {
        assert_eq!(sanitize_table_name("R"), "_R");
        assert_eq!(sanitize_table_name("c"), "_c");
        assert_eq!(sanitize_table_name("R1C1"), "_R1C1");
    }

    #[test]
    fn trailing_text_after_a_reference_still_gains_an_underscore() {
        // Excel ignores what follows a complete R1C1 reference, so these are
        // references to it and were repaired on opening (measured).
        assert_eq!(sanitize_table_name("R2D2"), "_R2D2");
        assert_eq!(sanitize_table_name("C3PO"), "_C3PO");
        assert_eq!(sanitize_table_name("R1_total"), "_R1_total");
    }

    #[test]
    fn logical_constants_gain_an_underscore() {
        // Excel repairs a workbook whose table is named "TRUE" in any case
        // (measured); rust_xlsxwriter 0.99.0 refuses it at `add_table`.
        assert_eq!(sanitize_table_name("TRUE"), "_TRUE");
        assert_eq!(sanitize_table_name("true"), "_true");
        assert_eq!(sanitize_table_name("False"), "_False");
    }

    #[test]
    fn ordinary_names_are_untouched() {
        // The control: without it the assertions above are satisfied by a
        // function that prefixes everything.
        assert_eq!(sanitize_table_name("Sales"), "Sales");
        assert_eq!(sanitize_table_name("Q1_Sales"), "Q1_Sales");
        assert_eq!(sanitize_table_name("Table1x"), "Table1x");
        // The controls for the two tests above: "RCx" carries no index, and
        // "TRUEX" is not the constant. Excel opens both without a prompt.
        assert_eq!(sanitize_table_name("RCx"), "RCx");
        assert_eq!(sanitize_table_name("Rate1"), "Rate1");
        assert_eq!(sanitize_table_name("TRUEX"), "TRUEX");
        assert_eq!(sanitize_table_name("Falsehood"), "Falsehood");
    }

    /// Decomposed names survive, because the allowlist below cannot see a
    /// combining mark as part of a letter.
    ///
    /// The expected values are Excel's, not ours: every input here was put to
    /// Excel's own name validator during the rust_xlsxwriter#189 work and
    /// accepted, so a rewrite is data loss. That external oracle is the point
    /// — `sanitized_table_names_are_always_valid` in `proptests.rs` asserts
    /// the same allowlist the function branches on, so it holds by
    /// construction and cannot fail on any of these.
    #[test]
    fn decomposed_names_are_composed_not_mangled() {
        // "Verkäufe" written as "Verka" + U+0308 COMBINING DIAERESIS.
        assert_eq!(
            sanitize_table_name("Verka\u{308}ufe"),
            "Verk\u{E4}ufe",
            "NFD umlaut must compose, not become an underscore"
        );
        // "é" as "e" + U+0301 COMBINING ACUTE.
        assert_eq!(sanitize_table_name("Caf\u{65}\u{301}"), "Caf\u{E9}");
        // "がくせい" with U+3099 COMBINING VOICED SOUND MARK.
        assert_eq!(
            sanitize_table_name("\u{304B}\u{3099}\u{304F}\u{305B}\u{3044}"),
            "\u{304C}\u{304F}\u{305B}\u{3044}"
        );
    }

    /// The controls for the test above: names that were already composed, or
    /// that contain no marks at all, must come through byte for byte.
    ///
    /// Without these, a `sanitize_table_name` that simply returned its input
    /// would satisfy every assertion up there.
    #[test]
    fn already_composed_names_are_unchanged_by_normalisation() {
        assert_eq!(sanitize_table_name("Verk\u{E4}ufe"), "Verk\u{E4}ufe");
        assert_eq!(
            sanitize_table_name("\u{65E5}\u{672C}\u{8A9E}"),
            "\u{65E5}\u{672C}\u{8A9E}"
        );
        assert_eq!(
            sanitize_table_name("\u{304C}\u{304F}\u{305B}\u{3044}"),
            "\u{304C}\u{304F}\u{305B}\u{3044}"
        );
    }

    /// Composition is NFC and deliberately not NFKC.
    ///
    /// NFKC folds U+FF21 FULLWIDTH LATIN CAPITAL A to ASCII "A", which would
    /// turn the distinct name "Ａ1" into the cell reference "A1" and then into
    /// "_A1". Excel treats the two as different names, so the compatibility
    /// forms must survive.
    #[test]
    fn compatibility_forms_are_not_folded() {
        assert_eq!(sanitize_table_name("\u{FF21}1"), "\u{FF21}1");
        assert_eq!(sanitize_table_name("\u{FB00}1"), "\u{FB00}1");
    }

    /// The known limit, pinned so it is a decision on record rather than a
    /// surprise: marks with no composed form are still rewritten.
    ///
    /// Excel accepts both of these. Widening the allowlist or inverting it to a
    /// denylist is what would fix them; see `AGENTS.md`. This test exists to go
    /// red when someone does that, so the change cannot land unnoticed.
    #[test]
    fn marks_without_a_composed_form_are_still_rewritten() {
        // Thai "\u{E44}\u{E21}\u{E48}", tone mark U+0E48.
        assert_eq!(
            sanitize_table_name("\u{E44}\u{E21}\u{E48}"),
            "\u{E44}\u{E21}_"
        );
        // Hindi "\u{939}\u{93F}\u{928}\u{94D}\u{926}\u{940}", virama U+094D.
        assert_eq!(
            sanitize_table_name("\u{939}\u{93F}\u{928}\u{94D}\u{926}\u{940}"),
            "\u{939}\u{93F}\u{928}_\u{926}\u{940}"
        );
    }

    #[test]
    fn beyond_the_grid_shapes_are_untouched() {
        assert_eq!(sanitize_table_name("AAAA1"), "AAAA1");
        assert_eq!(sanitize_table_name("A1048577"), "A1048577");
    }

    #[test]
    fn the_prefix_is_applied_after_the_digit_prefix_not_instead_of_it() {
        // "1Q1" -> "_1Q1": the digit prefix already made it a non-reference, so
        // the second prefix must not fire.
        assert_eq!(sanitize_table_name("1Q1"), "_1Q1");
        // A sanitised character can create the reference shape, and the check
        // has to see the *sanitised* string: "Q-1" -> "Q_1" is not a reference.
        assert_eq!(sanitize_table_name("Q-1"), "Q_1");
    }

    #[test]
    fn prefixing_stays_idempotent_and_capped() {
        let once = sanitize_table_name("Q1");
        assert_eq!(sanitize_table_name(&once), once);
        let long = "9".repeat(300);
        assert_eq!(sanitize_table_name(&long).chars().count(), 255);
    }
}
