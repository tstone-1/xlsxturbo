//! Pins a rust_xlsxwriter defect that xlsxturbo works around.
//!
//! A `data_bar` conditional format and a sparkline on the same worksheet make
//! rust_xlsxwriter 0.97.1 emit unbalanced `<ext>` elements in the worksheet XML:
//! three opened, two closed. The result is not well-formed, so Excel reports the
//! workbook as damaged.
//!
//! `apply::reject_databar_with_sparklines` refuses that combination, which means
//! the corrupt file is no longer reachable through xlsxturbo's own API -- and
//! therefore no Python test can detect the day upstream fixes it. This module
//! uses rust_xlsxwriter **directly**, with no xlsxturbo code in the path, for two
//! reasons: it is the evidence that the defect is upstream rather than ours, and
//! it is the only place the fix can be noticed.
//!
//! The assertions are deliberately the wrong way round: they assert the bug is
//! *still present*. When it goes away this fails, which is the signal to delete
//! the guard, this file, and the note in `docs/errors.md`. A workaround nobody
//! removes outlives its reason and quietly becomes a bug of its own.

use rust_xlsxwriter::{ConditionalFormatDataBar, Sparkline, Workbook, Worksheet, XlsxError};
use std::io::Read;

/// Build a worksheet with the requested features and return its XML.
fn sheet_xml(with_data_bar: bool, with_sparkline: bool) -> Result<String, XlsxError> {
    let mut workbook = Workbook::new();
    let sheet: &mut Worksheet = workbook.add_worksheet();
    sheet.write(0, 0, 1)?;
    sheet.write(1, 0, 2)?;
    if with_data_bar {
        sheet.add_conditional_format(0, 0, 1, 0, &ConditionalFormatDataBar::new())?;
    }
    if with_sparkline {
        sheet.add_sparkline(5, 5, &Sparkline::new().set_range(("Sheet1", 0, 0, 1, 0)))?;
    }
    let buffer = workbook.save_to_buffer()?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).expect("valid zip");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet present")
        .read_to_string(&mut xml)
        .expect("readable xml");
    Ok(xml)
}

/// `<ext ` opens and `</ext>` closes, counted independently.
fn ext_balance(xml: &str) -> (usize, usize) {
    (xml.matches("<ext ").count(), xml.matches("</ext>").count())
}

#[test]
fn databar_with_sparkline_is_still_unbalanced() {
    let xml = sheet_xml(true, true).expect("write succeeds; it is the output that is wrong");
    let (opened, closed) = ext_balance(&xml);
    assert_ne!(
        opened, closed,
        "rust_xlsxwriter now balances <ext> for a data bar beside a sparkline \
         ({opened} opened, {closed} closed). The upstream defect is FIXED: remove \
         `reject_databar_with_sparklines` and its call in convert.rs, restore the \
         data_bar sample in tests/test_options.py, drop the note in docs/errors.md, \
         and delete this file."
    );
}

#[test]
fn each_feature_alone_is_balanced() {
    // The control. Without it, this module would keep passing if sparklines or
    // data bars became malformed on their own -- a far worse defect, reported as
    // the known one.
    for (data_bar, sparkline, label) in [
        (true, false, "data bar alone"),
        (false, true, "sparkline alone"),
        (false, false, "neither"),
    ] {
        let xml = sheet_xml(data_bar, sparkline).expect("write succeeds");
        let (opened, closed) = ext_balance(&xml);
        assert_eq!(
            opened, closed,
            "{label}: {opened} <ext> opened, {closed} closed"
        );
    }
}
