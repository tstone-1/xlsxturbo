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

/// Sanitize a string for use as an Excel table name
///
/// The name is normalised to NFC first, then invalid characters become `_` and
/// a leading digit gains a `_` prefix. A name Excel would read as a cell
/// reference (`"Q1"`, `"A1"`, `"XFD1048576"`), an R1C1 form (`"R1C1"`) or a
/// reserved selection shortcut (`"R"`, `"C"`) gets the same `_` prefix —
/// `"Q1"` becomes `"_Q1"`, which addresses no cell. rust_xlsxwriter 0.98.2
/// stores the name verbatim and Excel then offers to repair the workbook
/// (measured, rust_xlsxwriter#189), so the screen has to happen here. See
/// [`looks_like_cell_reference`] for why the beyond-the-grid shapes
/// (`"AAAA1"`) are deliberately left alone.
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

    // Must not collide with a cell reference. Done before the cap so the
    // 255-character invariant holds unconditionally; a reference-shaped name is
    // at most 10 characters, so the two branches never both fire.
    if looks_like_cell_reference(&sanitized) {
        sanitized = format!("_{}", sanitized);
    }

    // Max 255 chars. Build by chars (not bytes) so a multibyte codepoint at the
    // boundary can never trigger a `truncate` mid-char-boundary panic.
    if sanitized.chars().count() > 255 {
        sanitized = sanitized.chars().take(255).collect();
    }
    sanitized
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
    fn ordinary_names_are_untouched() {
        // The control: without it the assertions above are satisfied by a
        // function that prefixes everything.
        assert_eq!(sanitize_table_name("Sales"), "Sales");
        assert_eq!(sanitize_table_name("Q1_Sales"), "Q1_Sales");
        assert_eq!(sanitize_table_name("Table1x"), "Table1x");
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
