//! Parsing and utility functions.

mod cell_refs;
mod colors;
mod formats;
mod patterns;
mod tables;
mod values;

pub(crate) use cell_refs::{parse_cell_range, parse_cell_ref};
pub(crate) use colors::{parse_color, parse_color_enum};
pub(crate) use formats::{
    build_column_formats, parse_column_format, parse_header_format, parse_horizontal_alignment,
    parse_icon_type, parse_rich_text_format, parse_vertical_alignment,
};
pub(crate) use patterns::matches_pattern;
pub(crate) use tables::{parse_table_style, sanitize_table_name};
pub(crate) use values::{naive_date_to_excel, naive_datetime_to_excel, parse_value};

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
    fn test_naive_date_to_excel_epoch() {
        // Excel epoch is 1899-12-30, so 1900-01-01 = day 2
        let date = chrono::NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
        assert_eq!(naive_date_to_excel(date), 2.0);
    }

    #[test]
    fn test_naive_date_to_excel_known_date() {
        // 2024-01-15 is a known Excel serial date
        let date = chrono::NaiveDate::from_ymd_opt(2024, 1, 15).unwrap();
        assert_eq!(naive_date_to_excel(date), 45306.0);
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
        // Dates before 1900-01-01 should be treated as strings, not invalid serial numbers
        let result = super::parse_value("1899-01-01", crate::types::DateOrder::Auto);
        assert!(matches!(result, crate::types::CellValue::String(_)));
    }
}
