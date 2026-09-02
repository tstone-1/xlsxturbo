//! Parsing and utility functions.

mod cell_refs;
mod colors;
mod context;
mod formats;
mod patterns;
mod tables;
mod values;

pub(crate) use cell_refs::{looks_like_cell_reference, parse_cell_range, parse_cell_ref};
pub(crate) use colors::{parse_color, parse_color_enum};
pub(crate) use context::WithOptionContext;
pub(crate) use formats::{
    build_column_formats, parse_column_format, parse_header_format, parse_horizontal_alignment,
    parse_icon_type, parse_rich_text_format, parse_vertical_alignment,
};
pub(crate) use patterns::matches_pattern;
pub(crate) use tables::{parse_table_style, sanitize_table_name};
pub(crate) use values::{naive_date_to_excel, naive_datetime_to_excel, parse_value};

#[cfg(test)]
mod boundaries;

#[cfg(test)]
mod proptests;

#[cfg(test)]
mod tests {
    use super::formats::parse_border_style;
    use super::{
        matches_pattern, naive_date_to_excel, parse_cell_range, parse_cell_ref, parse_color,
        parse_horizontal_alignment, parse_table_style, parse_value, parse_vertical_alignment,
        sanitize_table_name,
    };
    use crate::types::{CellValue, DateOrder};

    #[test]
    fn test_parse_integer() {
        assert!(matches!(
            parse_value("123", DateOrder::Auto),
            CellValue::Integer(123)
        ));
        assert!(matches!(
            parse_value("-456", DateOrder::Auto),
            CellValue::Integer(-456)
        ));
    }

    #[test]
    fn integers_beyond_i64_parse_to_text_not_to_a_rounded_float() {
        // The f64 branch parses any run of digits, so before this screen a CSV
        // cell of 20 digits was written as a rounded number -- against the
        // documented "above 2^53 becomes text" guarantee that the DataFrame
        // path keeps. Each case here is a value f64 cannot hold exactly.
        for text in [
            "9223372036854775808",      // i64::MAX + 1
            "18446744073709551615",     // u64::MAX
            "18446744073709551616",     // u64::MAX + 1, beyond every Rust int
            "-9223372036854775809",     // i64::MIN - 1
            "-99999999999999999999999", // far below i64::MIN
            "123456789012345678901234567890",
        ] {
            let parsed = parse_value(text, DateOrder::Auto);
            assert!(
                matches!(&parsed, CellValue::String(s) if s == text),
                "expected the text {:?} back, got {:?}",
                text,
                parsed
            );
        }

        // A leading `+` is dropped while the value still fits a Rust integer,
        // matching what the i64 arm does with "+42"; past u64 the digits are
        // written through as given because nothing can re-parse them.
        assert!(
            matches!(parse_value("+18446744073709551615", DateOrder::Auto),
                     CellValue::String(ref s) if s == "18446744073709551615")
        );

        // A digit string long enough that f64 parses it as infinity used to
        // become an empty cell, which loses the value outright.
        let huge = "1".repeat(400);
        assert!(
            matches!(parse_value(&huge, DateOrder::Auto), CellValue::String(ref s) if *s == huge)
        );
    }

    #[test]
    fn the_integer_text_screen_does_not_swallow_floats_or_i64_values() {
        // Controls for the screen above: everything an f64 should still own,
        // and the i64 range it must not reach at all.
        assert!(matches!(
            parse_value("9007199254740992", DateOrder::Auto), // 2^53, fits i64
            CellValue::Integer(9_007_199_254_740_992)
        ));
        assert!(matches!(
            parse_value("007", DateOrder::Auto),
            CellValue::Integer(7)
        ));
        // Numbers the float branch owns must still reach it as numbers.
        for text in ["1e3", "1.0", "-2.5", ".5", "1e400"] {
            let parsed = parse_value(text, DateOrder::Auto);
            assert!(
                matches!(parsed, CellValue::Float(_) | CellValue::Empty),
                "{:?} must stay on the float branch, got {:?}",
                text,
                parsed
            );
        }

        // Text that only looks numeric keeps falling through to the string
        // default, untrimmed, as it did before.
        for text in ["1_000", "٣٤", "12a", "+", "-", ""] {
            let parsed = parse_value(text, DateOrder::Auto);
            let expected = if text.is_empty() {
                CellValue::Empty
            } else {
                CellValue::String(text.to_string())
            };
            assert_eq!(
                format!("{:?}", parsed),
                format!("{:?}", expected),
                "unexpected parse of {:?}",
                text
            );
        }
    }

    #[test]
    fn the_length_gate_on_the_integer_screens_falls_exactly_on_19_bytes() {
        // The two overflow screens sit behind `trimmed.len() >= 19` so that
        // ordinary cells skip them. 19 is the shortest an integer literal that
        // fails `i64` can be (`i64::MAX` + 1 is 9223372036854775808), so the
        // gate must let 19 bytes through and may stop at 18. These cases pin
        // both sides of it; widening the gate to 20 loses the first, and
        // narrowing it to 18 would only cost time, not behaviour.
        //
        // 19 bytes and not an integer: the float branch still owns it.
        let float_19 = "1234567890.12345678";
        assert_eq!(float_19.len(), 19);
        assert!(
            matches!(parse_value(float_19, DateOrder::Auto), CellValue::Float(_)),
            "19-byte float changed branch: {:?}",
            parse_value(float_19, DateOrder::Auto)
        );

        // 19 bytes and an integer above i64::MAX: text, via the screens.
        let int_19 = "9999999999999999999";
        assert_eq!(int_19.len(), 19);
        assert!(
            matches!(parse_value(int_19, DateOrder::Auto), CellValue::String(ref s) if s == int_19),
            "19-digit integer above i64::MAX must stay text, got {:?}",
            parse_value(int_19, DateOrder::Auto)
        );

        // 18 bytes: below the gate, and unchanged either way -- a float stays
        // a float, and every 18-digit integer fits i64.
        let float_18 = "123456789.12345678";
        assert_eq!(float_18.len(), 18);
        assert!(matches!(
            parse_value(float_18, DateOrder::Auto),
            CellValue::Float(_)
        ));
        let int_18 = "999999999999999999";
        assert_eq!(int_18.len(), 18);
        assert!(matches!(
            parse_value(int_18, DateOrder::Auto),
            CellValue::Integer(999_999_999_999_999_999)
        ));
    }

    #[test]
    fn test_parse_float() {
        let value = parse_value("3.25", DateOrder::Auto);
        assert!(
            matches!(value, CellValue::Float(_)),
            "Expected CellValue::Float, got {:?}",
            value
        );
        if let CellValue::Float(v) = value {
            assert!((v - 3.25).abs() < 0.001);
        }
    }

    #[test]
    fn test_parse_boolean() {
        assert!(matches!(
            parse_value("true", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("TRUE", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("false", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
        assert!(matches!(
            parse_value("False", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_parse_empty() {
        assert!(matches!(parse_value("", DateOrder::Auto), CellValue::Empty));
        assert!(matches!(
            parse_value("   ", DateOrder::Auto),
            CellValue::Empty
        ));
        assert!(matches!(
            parse_value("NaN", DateOrder::Auto),
            CellValue::Empty
        ));
    }

    #[test]
    fn test_parse_date() {
        assert!(matches!(
            parse_value("2024-01-15", DateOrder::Auto),
            CellValue::Date(_)
        ));
        assert!(matches!(
            parse_value("2024/01/15", DateOrder::Auto),
            CellValue::Date(_)
        ));
    }

    #[test]
    fn test_parse_datetime() {
        assert!(matches!(
            parse_value("2024-01-15T10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
        assert!(matches!(
            parse_value("2024-01-15 10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
    }

    #[test]
    fn test_parse_datetime_preserves_fractional_seconds() {
        let value = parse_value("2024-01-15T10:30:00.250", DateOrder::Auto);
        let CellValue::DateTime(serial) = value else {
            panic!("expected datetime");
        };
        let expected = super::naive_datetime_to_excel(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
                .unwrap()
                .and_hms_milli_opt(10, 30, 0, 250)
                .unwrap(),
        );
        assert!((serial - expected).abs() < 0.000000001);
    }

    #[test]
    fn test_parse_string() {
        assert!(matches!(
            parse_value("hello", DateOrder::Auto),
            CellValue::String(_)
        ));
    }

    #[test]
    fn test_matches_pattern_exact() {
        assert!(matches_pattern("column_name", "column_name"));
        assert!(!matches_pattern("column_name", "other"));
    }

    #[test]
    fn test_matches_pattern_prefix() {
        assert!(matches_pattern("price_usd", "price_*"));
        assert!(matches_pattern("price_", "price_*"));
        assert!(!matches_pattern("cost_usd", "price_*"));
    }

    #[test]
    fn test_matches_pattern_suffix() {
        assert!(matches_pattern("col_weight", "*_weight"));
        assert!(matches_pattern("_weight", "*_weight"));
        assert!(!matches_pattern("col_height", "*_weight"));
    }

    #[test]
    fn test_matches_pattern_contains() {
        assert!(matches_pattern("leadframe_difference", "*difference*"));
        assert!(matches_pattern("difference", "*difference*"));
        assert!(matches_pattern("my_difference_col", "*difference*"));
        assert!(!matches_pattern("other_column", "*difference*"));
    }

    #[test]
    fn test_matches_pattern_wildcard() {
        // Single "*" matches everything
        assert!(matches_pattern("anything", "*"));
        assert!(matches_pattern("", "*"));
        // Double "**" also matches everything
        assert!(matches_pattern("anything", "**"));
        assert!(matches_pattern("", "**"));
    }

    // --- parse_cell_ref tests ---

    #[test]
    fn test_parse_cell_ref_basic() {
        assert_eq!(parse_cell_ref("A1").unwrap(), (0, 0));
        assert_eq!(parse_cell_ref("B2").unwrap(), (1, 1));
        assert_eq!(parse_cell_ref("Z1").unwrap(), (0, 25));
        assert_eq!(parse_cell_ref("AA1").unwrap(), (0, 26));
        assert_eq!(parse_cell_ref("AZ1").unwrap(), (0, 51));
    }

    #[test]
    fn test_parse_cell_ref_case_insensitive() {
        assert_eq!(parse_cell_ref("a1").unwrap(), (0, 0));
        assert_eq!(parse_cell_ref("aa1").unwrap(), (0, 26));
    }

    #[test]
    fn test_parse_cell_ref_max_column() {
        // XFD = 16384th column = index 16383
        assert_eq!(parse_cell_ref("XFD1").unwrap(), (0, 16383));
    }

    #[test]
    fn test_parse_cell_ref_overflow_column() {
        assert!(parse_cell_ref("ZZZZ1").is_err());
    }

    #[test]
    fn test_parse_cell_ref_exceeds_excel_max() {
        // XFE = 16385th column, exceeds Excel max
        assert!(parse_cell_ref("XFE1").is_err());
    }

    #[test]
    fn test_parse_cell_ref_row_zero() {
        assert!(parse_cell_ref("A0").is_err());
    }

    #[test]
    fn test_parse_cell_ref_empty() {
        assert!(parse_cell_ref("").is_err());
    }

    #[test]
    fn test_parse_cell_ref_no_row() {
        assert!(parse_cell_ref("A").is_err());
    }

    #[test]
    fn test_parse_cell_ref_no_column() {
        assert!(parse_cell_ref("1").is_err());
    }

    // --- parse_cell_range tests ---

    #[test]
    fn test_parse_cell_range_basic() {
        assert_eq!(parse_cell_range("A1:B2").unwrap(), (0, 0, 1, 1));
        assert_eq!(parse_cell_range("A1:D1").unwrap(), (0, 0, 0, 3));
    }

    #[test]
    fn test_parse_cell_range_invalid_format() {
        assert!(parse_cell_range("A1").is_err()); // no colon
        assert!(parse_cell_range("A1:B2:C3").is_err()); // too many colons
    }

    // --- parse_color tests ---

    #[test]
    fn test_parse_color_hex() {
        assert_eq!(parse_color("#FF0000").unwrap(), 0xFF0000);
        assert_eq!(parse_color("#000000").unwrap(), 0x000000);
        assert_eq!(parse_color("#FFFFFF").unwrap(), 0xFFFFFF);
        assert_eq!(parse_color("#4F81BD").unwrap(), 0x4F81BD);
    }

    #[test]
    fn test_parse_color_named() {
        assert_eq!(parse_color("red").unwrap(), 0xFF0000);
        assert_eq!(parse_color("Red").unwrap(), 0xFF0000);
        assert_eq!(parse_color("WHITE").unwrap(), 0xFFFFFF);
        assert_eq!(parse_color("gray").unwrap(), 0x808080);
        assert_eq!(parse_color("grey").unwrap(), 0x808080);
    }

    #[test]
    fn test_parse_color_invalid() {
        assert!(parse_color("#FFF").is_err()); // too short
        assert!(parse_color("#GGGGGG").is_err()); // invalid hex
        assert!(parse_color("chartreuse").is_err()); // unsupported name
    }

    #[test]
    fn test_parse_color_whitespace() {
        assert_eq!(parse_color("  #FF0000  ").unwrap(), 0xFF0000);
        assert_eq!(parse_color("  red  ").unwrap(), 0xFF0000);
    }

    #[test]
    fn test_parse_color_hex_rejects_sign_characters() {
        // u32::from_str_radix accepts a leading '+'/'-', which would otherwise
        // let a 6-character-looking string like "+12345" slip through as a
        // valid (but wrong) color instead of being rejected.
        assert!(parse_color("#+12345").is_err());
        assert!(parse_color("#-12345").is_err());
        // A genuine 6-digit hex color is unaffected.
        assert_eq!(parse_color("#A1B2C3").unwrap(), 0xA1B2C3);
        // Case-insensitivity of hex digits is preserved.
        assert_eq!(parse_color("#a1b2c3").unwrap(), 0xA1B2C3);
    }

    // --- sanitize_table_name tests ---

    #[test]
    fn test_sanitize_table_name_valid() {
        assert_eq!(sanitize_table_name("MyTable"), "MyTable");
        assert_eq!(sanitize_table_name("_table1"), "_table1");
    }

    #[test]
    fn test_sanitize_table_name_special_chars() {
        assert_eq!(sanitize_table_name("My Table!"), "My_Table_");
        assert_eq!(sanitize_table_name("data-2024"), "data_2024");
    }

    #[test]
    fn test_sanitize_table_name_starts_with_digit() {
        assert_eq!(sanitize_table_name("123Data"), "_123Data");
    }

    #[test]
    fn test_sanitize_table_name_truncation() {
        let long_name = "a".repeat(300);
        let sanitized = sanitize_table_name(&long_name);
        assert_eq!(sanitized.len(), 255);
    }

    #[test]
    fn test_sanitize_table_name_multibyte_truncation() {
        let long_name = "é".repeat(300);
        let sanitized = sanitize_table_name(&long_name);
        assert_eq!(sanitized.chars().count(), 255);
    }

    #[test]
    fn test_sanitize_table_name_empty() {
        assert_eq!(sanitize_table_name(""), "_");
    }

    // --- parse_table_style tests ---

    #[test]
    fn test_parse_table_style_valid() {
        assert!(parse_table_style("None").is_ok());
        assert!(parse_table_style("Light1").is_ok());
        assert!(parse_table_style("Medium14").is_ok());
        assert!(parse_table_style("Dark11").is_ok());
    }

    #[test]
    fn test_parse_table_style_invalid() {
        assert!(parse_table_style("light1").is_err()); // case-sensitive
        assert!(parse_table_style("Medium29").is_err()); // out of range
        assert!(parse_table_style("Dark12").is_err()); // out of range
        assert!(parse_table_style("").is_err());
    }

    // --- naive_date_to_excel tests ---

    #[test]
    fn test_naive_date_to_excel_epoch_formula() {
        // The raw epoch-based formula (1899-12-30) returns 2.0 for
        // 1900-01-01, one more than the real Excel serial (1). This is only
        // correct from 1900-03-01 (serial 61) onward; see
        // `test_parse_value_pre_march_1900_falls_back_to_string` for the
        // caller-facing guard that rejects this range instead of writing an
        // off-by-one date.
        let date = chrono::NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
        assert_eq!(naive_date_to_excel(date), 2.0);
    }

    #[test]
    fn test_naive_date_to_excel_known_date() {
        // 2024-01-15 is a known Excel serial date
        let date = chrono::NaiveDate::from_ymd_opt(2024, 1, 15).unwrap();
        assert_eq!(naive_date_to_excel(date), 45306.0);
    }

    #[test]
    fn test_naive_date_to_excel_boundary() {
        // 1900-03-01 is the first date the epoch-based formula gets right
        // (serial 61); everything before it is rejected by parse_value.
        let date = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).unwrap();
        assert_eq!(naive_date_to_excel(date), 61.0);
    }

    #[test]
    fn test_parse_value_pre_march_1900_falls_back_to_string() {
        // 1900-01-01 previously parsed as serial 2.0, which Excel renders as
        // 1900-01-02 (one day late) because of the 1900 leap-year bug. It
        // must now fall back to a string instead of writing a wrong date.
        let result = parse_value("1900-01-01", DateOrder::Auto);
        match result {
            CellValue::String(s) => assert_eq!(s, "1900-01-01"),
            other => panic!("expected String fallback, got {:?}", other),
        }

        // 1900-02-28 is the last date affected by the bug; still a string.
        let result = parse_value("1900-02-28", DateOrder::Auto);
        match result {
            CellValue::String(s) => assert_eq!(s, "1900-02-28"),
            other => panic!("expected String fallback, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_value_first_correct_1900_date() {
        // 1900-03-01 (serial 61) is the first date the formula gets right,
        // so it should parse as a real Date, not fall back to string.
        let result = parse_value("1900-03-01", DateOrder::Auto);
        match result {
            CellValue::Date(v) => assert_eq!(v, 61.0),
            other => panic!("expected Date(61.0), got {:?}", other),
        }
    }

    #[test]
    fn test_parse_value_modern_date_unaffected() {
        // Modern dates are well past the 1900-03-01 boundary and must keep
        // parsing as dates exactly as before.
        let result = parse_value("2024-01-15", DateOrder::Auto);
        match result {
            CellValue::Date(v) => assert_eq!(v, 45306.0),
            other => panic!("expected Date(45306.0), got {:?}", other),
        }
    }

    // --- DateOrder tests ---

    #[test]
    fn test_date_order_parse() {
        assert_eq!(DateOrder::parse("auto"), Some(DateOrder::Auto));
        assert_eq!(DateOrder::parse("mdy"), Some(DateOrder::MDY));
        assert_eq!(DateOrder::parse("us"), Some(DateOrder::MDY));
        assert_eq!(DateOrder::parse("dmy"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("eu"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("european"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("AUTO"), Some(DateOrder::Auto));
        assert_eq!(DateOrder::parse("invalid"), None);
        assert_eq!(DateOrder::parse(""), None);
    }

    // --- parse_border_style tests ---

    #[test]
    fn test_parse_border_style_valid() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(parse_border_style("thin").unwrap(), FormatBorder::Thin);
        assert_eq!(parse_border_style("medium").unwrap(), FormatBorder::Medium);
        assert_eq!(parse_border_style("thick").unwrap(), FormatBorder::Thick);
        assert_eq!(parse_border_style("dashed").unwrap(), FormatBorder::Dashed);
        assert_eq!(parse_border_style("dotted").unwrap(), FormatBorder::Dotted);
        assert_eq!(parse_border_style("double").unwrap(), FormatBorder::Double);
        assert_eq!(parse_border_style("hair").unwrap(), FormatBorder::Hair);
    }

    #[test]
    fn test_parse_border_style_case_insensitive() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(parse_border_style("THIN").unwrap(), FormatBorder::Thin);
        assert_eq!(parse_border_style("Thick").unwrap(), FormatBorder::Thick);
        assert_eq!(parse_border_style("Medium").unwrap(), FormatBorder::Medium);
    }

    #[test]
    fn test_parse_border_style_aliases() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(
            parse_border_style("medium_dashed").unwrap(),
            FormatBorder::MediumDashed
        );
        assert_eq!(
            parse_border_style("mediumdashed").unwrap(),
            FormatBorder::MediumDashed
        );
        assert_eq!(
            parse_border_style("dash_dot").unwrap(),
            FormatBorder::DashDot
        );
        assert_eq!(
            parse_border_style("dashdot").unwrap(),
            FormatBorder::DashDot
        );
        assert_eq!(
            parse_border_style("slant_dash_dot").unwrap(),
            FormatBorder::SlantDashDot
        );
        assert_eq!(
            parse_border_style("slantdashdot").unwrap(),
            FormatBorder::SlantDashDot
        );
    }

    #[test]
    fn test_parse_border_style_invalid() {
        assert!(parse_border_style("").is_err());
        assert!(parse_border_style("bold").is_err());
        assert!(parse_border_style("heavy").is_err());
    }

    // --- parse_horizontal_alignment tests ---

    #[test]
    fn test_parse_horizontal_alignment_valid() {
        use rust_xlsxwriter::FormatAlign;
        assert_eq!(
            parse_horizontal_alignment("left").unwrap(),
            FormatAlign::Left
        );
        assert_eq!(
            parse_horizontal_alignment("center").unwrap(),
            FormatAlign::Center
        );
        assert_eq!(
            parse_horizontal_alignment("right").unwrap(),
            FormatAlign::Right
        );
        assert_eq!(
            parse_horizontal_alignment("fill").unwrap(),
            FormatAlign::Fill
        );
        assert_eq!(
            parse_horizontal_alignment("justify").unwrap(),
            FormatAlign::Justify
        );
        assert_eq!(
            parse_horizontal_alignment("CENTER").unwrap(),
            FormatAlign::Center
        );
    }

    #[test]
    fn test_parse_horizontal_alignment_invalid() {
        assert!(parse_horizontal_alignment("").is_err());
        assert!(parse_horizontal_alignment("top").is_err());
        assert!(parse_horizontal_alignment("middle").is_err());
    }

    // --- parse_vertical_alignment tests ---

    #[test]
    fn test_parse_vertical_alignment_valid() {
        use rust_xlsxwriter::FormatAlign;
        assert_eq!(parse_vertical_alignment("top").unwrap(), FormatAlign::Top);
        assert_eq!(
            parse_vertical_alignment("center").unwrap(),
            FormatAlign::VerticalCenter
        );
        assert_eq!(
            parse_vertical_alignment("bottom").unwrap(),
            FormatAlign::Bottom
        );
        assert_eq!(
            parse_vertical_alignment("justify").unwrap(),
            FormatAlign::VerticalJustify
        );
        assert_eq!(parse_vertical_alignment("TOP").unwrap(), FormatAlign::Top);
    }

    #[test]
    fn test_parse_vertical_alignment_invalid() {
        assert!(parse_vertical_alignment("").is_err());
        assert!(parse_vertical_alignment("left").is_err());
        assert!(parse_vertical_alignment("right").is_err());
        assert!(parse_vertical_alignment("general").is_err());
    }

    // --- naive_datetime_to_excel tests ---

    #[test]
    fn test_naive_datetime_to_excel_noon() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(12, 0, 0)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        // 2024-01-15 = 45306.0, noon = 0.5
        assert!((result - 45306.5).abs() < 0.001);
    }

    #[test]
    fn test_naive_datetime_to_excel_midnight() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        assert!((result - 45306.0).abs() < 0.001);
    }

    #[test]
    fn test_naive_datetime_to_excel_end_of_day() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(23, 59, 59)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        assert!((result - 45307.0).abs() < 0.001); // just under next day
    }

    #[test]
    fn test_naive_datetime_to_excel_fractional_seconds() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_micro_opt(12, 0, 0, 500_000)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        assert!((result - 45306.50000578704).abs() < 0.000000001);
    }

    // --- parse_icon_type tests ---

    #[test]
    fn test_parse_icon_type_valid() {
        assert!(super::parse_icon_type("3_arrows").is_ok());
        assert!(super::parse_icon_type("3arrows").is_ok());
        assert!(super::parse_icon_type("3_flags").is_ok());
        assert!(super::parse_icon_type("3_traffic_lights").is_ok());
        assert!(super::parse_icon_type("4_arrows").is_ok());
        assert!(super::parse_icon_type("5_quarters").is_ok());
        assert!(super::parse_icon_type("5_rating").is_ok());
    }

    #[test]
    fn test_parse_icon_type_case_insensitive() {
        assert!(super::parse_icon_type("3_ARROWS").is_ok());
        assert!(super::parse_icon_type("5_Quarters").is_ok());
    }

    #[test]
    fn test_parse_icon_type_invalid() {
        assert!(super::parse_icon_type("").is_err());
        assert!(super::parse_icon_type("6_arrows").is_err());
        assert!(super::parse_icon_type("invalid").is_err());
    }

    // --- naive_date_to_excel pre-epoch guard tests ---

    #[test]
    fn test_naive_date_to_excel_pre_epoch() {
        // Dates before 1900-03-01 (serial 61) should be treated as strings,
        // not invalid or off-by-one serial numbers.
        let result = parse_value("1899-01-01", DateOrder::Auto);
        assert!(matches!(result, CellValue::String(_)));
    }

    // --- parse_value whitespace-preservation tests (Fix W5) ---

    #[test]
    fn test_parse_value_preserves_padded_string() {
        // Type detection operates on the trimmed text, but a genuine string
        // fallback must preserve the original, untrimmed value.
        let result = parse_value(" padded ", DateOrder::Auto);
        match result {
            CellValue::String(s) => assert_eq!(s, " padded "),
            other => panic!("expected String(\" padded \"), got {:?}", other),
        }
    }

    #[test]
    fn test_parse_value_padded_number_still_detected() {
        // Trimming still applies to type detection: a padded numeric string
        // is recognized as a number (its numeric value, not text).
        let result = parse_value(" 123 ", DateOrder::Auto);
        assert!(matches!(result, CellValue::Integer(123)));
    }

    #[test]
    fn test_parse_value_whitespace_only_is_empty() {
        // Whitespace-only input trims to empty and keeps the existing
        // empty-cell behavior (CellValue::Empty), not a padded empty string.
        let result = parse_value("   ", DateOrder::Auto);
        assert!(matches!(result, CellValue::Empty));
    }
}
