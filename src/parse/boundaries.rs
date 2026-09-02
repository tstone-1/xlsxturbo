//! Named edges of the date and serial conversions.
//!
//! `proptests.rs` states the rules; this file pins the specific points where a
//! rule changes. They are separated because they fail differently: a property
//! tells you an invariant broke somewhere in a space, a boundary test tells you
//! which exact input moved.
//!
//! Every value here is a date Excel itself treats specially. The 1900
//! leap-year bug is the reason most of them exist: Excel numbers dates as if
//! 1900-02-29 happened, and it did not, so serials 1 through 60 are off by one
//! from any correct calendar and serial 60 denotes a day that never existed.

use super::{naive_date_to_excel, naive_datetime_to_excel, parse_value};
use crate::types::{CellValue, DateOrder};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

/// Parse an ISO date/datetime string with the default ordering.
fn parse(text: &str) -> CellValue {
    parse_value(text, DateOrder::Auto)
}

/// Build a date, panicking on an invalid one (every literal here is checked).
fn date(y: i32, m: u32, d: u32) -> NaiveDate {
    NaiveDate::from_ymd_opt(y, m, d).expect("test date literal is valid")
}

#[test]
fn excel_serial_anchors() {
    // The epoch itself. Serial 0 is 1899-12-30 only because of the phantom
    // leap day; the "real" Excel day 1 is 1900-01-01.
    assert_eq!(naive_date_to_excel(date(1899, 12, 30)), 0.0);
    assert_eq!(naive_date_to_excel(date(1900, 1, 1)), 2.0);
    // 1900-02-28 is the last day before the gap. Excel calls it 59; this
    // formula calls it 60, which is exactly the off-by-one that makes every
    // serial below 61 unusable.
    assert_eq!(naive_date_to_excel(date(1900, 2, 28)), 60.0);
    // The first day both agree on.
    assert_eq!(naive_date_to_excel(date(1900, 3, 1)), 61.0);
    // A modern anchor, and Excel's last representable day.
    assert_eq!(naive_date_to_excel(date(2024, 1, 15)), 45_306.0);
    assert_eq!(naive_date_to_excel(date(9999, 12, 31)), 2_958_465.0);
}

#[test]
fn the_phantom_leap_day_cannot_be_constructed() {
    // Not a quirk of this crate -- 1900 was not a leap year, so the date Excel
    // reserves serial 60 for does not exist in any correct calendar.
    assert!(NaiveDate::from_ymd_opt(1900, 2, 29).is_none());
}

#[test]
fn dates_on_the_wrong_side_of_the_gap_stay_strings() {
    for text in [
        "1899-12-29", // negative serial
        "1899-12-30", // serial 0
        "1900-01-01", // Excel's day 1
        "1900-02-28", // last day before the gap
    ] {
        let parsed = parse(text);
        assert!(
            matches!(&parsed, CellValue::String(got) if got == text),
            "{} should stay a string, got {:?}",
            text,
            parsed
        );
    }
}

#[test]
fn the_first_correctly_numbered_day_is_a_date() {
    let parsed = parse("1900-03-01");
    assert!(
        matches!(parsed, CellValue::Date(serial) if serial == 61.0),
        "expected Date(61.0), got {:?}",
        parsed
    );
}

#[test]
fn the_boundary_also_applies_to_datetimes() {
    // A datetime one second before the boundary carries a serial below 61 and
    // must fall back too. Both guards in `values.rs` are needed; neither
    // covers the other's branch.
    let parsed = parse("1900-02-28 23:59:59");
    assert!(
        matches!(&parsed, CellValue::String(_)),
        "expected a string, got {:?}",
        parsed
    );
    let parsed = parse("1900-03-01 00:00:00");
    assert!(
        matches!(parsed, CellValue::DateTime(serial) if serial == 61.0),
        "expected DateTime(61.0), got {:?}",
        parsed
    );
}

#[test]
fn excels_last_representable_day_is_still_a_date() {
    let parsed = parse("9999-12-31");
    assert!(
        matches!(parsed, CellValue::Date(serial) if serial == 2_958_465.0),
        "expected Date(2958465.0), got {:?}",
        parsed
    );
}

#[test]
fn midnight_carries_no_fraction() {
    let d = date(2024, 1, 15);
    let midnight = NaiveTime::from_hms_opt(0, 0, 0).expect("valid time");
    // A date and a datetime at midnight are the same number, so a column that
    // mixes the two does not sort into two groups.
    assert_eq!(
        naive_datetime_to_excel(NaiveDateTime::new(d, midnight)),
        naive_date_to_excel(d)
    );
}

#[test]
fn a_microsecond_before_midnight_is_still_the_same_day() {
    let d = date(2024, 1, 15);
    let last = NaiveTime::from_hms_micro_opt(23, 59, 59, 999_999).expect("valid time");
    let serial = naive_datetime_to_excel(NaiveDateTime::new(d, last));
    assert!(serial < naive_date_to_excel(d) + 1.0);
    assert!(serial > naive_date_to_excel(d) + 0.999_999);
}

#[test]
fn sub_microsecond_precision_near_midnight_rolls_into_the_next_day() {
    // An Excel serial is an `f64` count of days, so its time resolution
    // depends on the magnitude of the date. At 2024 (serial ~45,306) one ulp
    // is 2^-37 days -- about 630 nanoseconds -- and any instant closer to
    // midnight than half of that rounds up to the next day's serial.
    //
    // Not a defect in this crate and not fixable here: it is what the file
    // format stores. Pinned because the first assertion written for this test
    // was the intuitive one -- that the last nanosecond of a day stays inside
    // the day -- and it is false.
    let d = date(2024, 1, 15);
    let last_nano = NaiveTime::from_hms_nano_opt(23, 59, 59, 999_999_999).expect("valid time");
    let serial = naive_datetime_to_excel(NaiveDateTime::new(d, last_nano));
    assert_eq!(serial, naive_date_to_excel(date(2024, 1, 16)));

    // The resolution is a function of the date's magnitude, so an early date
    // keeps more of it. Serial 61 has an ulp of 2^-46 days (~1.2 ns), and the
    // same instant stays inside its own day.
    let early = date(1900, 3, 1);
    let early_serial = naive_datetime_to_excel(NaiveDateTime::new(early, last_nano));
    assert!(early_serial < naive_date_to_excel(early) + 1.0);
}

#[test]
fn a_leap_second_lands_on_the_following_midnight() {
    // `%S` accepts 60, so a CSV really can contain one. chrono represents it as
    // 23:59:59 plus a full extra second in the nanosecond field, which makes
    // the time fraction exactly 1.0 -- the leap second becomes the next day's
    // midnight.
    //
    // Excel has no representation for a leap second, so *some* lie is
    // unavoidable, and rolling forward is the one that keeps the values
    // ordered. Pinned here because it is surprising enough to be "fixed" into
    // something worse by someone who meets it without context.
    let parsed = parse("2016-12-31 23:59:60");
    let expected = naive_date_to_excel(date(2017, 1, 1));
    assert!(
        matches!(parsed, CellValue::DateTime(serial) if serial == expected),
        "expected DateTime({}), got {:?}",
        expected,
        parsed
    );
}

#[test]
fn a_real_leap_day_is_unremarkable() {
    // The control for the 1900 tests above: 2000 *was* a leap year (divisible
    // by 400), so this is an ordinary date and must not be caught by any of
    // the fallbacks.
    let parsed = parse("2000-02-29");
    assert!(
        matches!(parsed, CellValue::Date(serial) if serial == naive_date_to_excel(date(2000, 2, 29))),
        "expected a date, got {:?}",
        parsed
    );
}

#[test]
fn integer_boundaries_are_detected_as_integers() {
    // Type *detection* has no 2^53 cutoff -- that belongs to the write path
    // (`src/write.rs`), which falls back to a string. Detection must hand the
    // exact value over for that decision to be possible at all, so a value
    // that arrives as a float here is already lossy whatever the writer does.
    for v in [
        i64::MIN,
        i64::MAX,
        1i64 << 53,
        (1i64 << 53) + 1,
        -(1i64 << 53),
    ] {
        let text = v.to_string();
        let parsed = parse(&text);
        assert!(
            matches!(parsed, CellValue::Integer(got) if got == v),
            "expected Integer({}), got {:?}",
            v,
            parsed
        );
    }
}

#[test]
fn integers_beyond_i64_become_text_carrying_every_digit() {
    // Past `i64::MAX` the integer branch stops matching. This used to accept a
    // `Float` as well, which is how a 20-digit CSV cell shipped rounded for
    // several releases: the assertion `f > 9.223e18` is satisfied by a value
    // that has already lost its low digits, so it could not see the defect.
    // Text is now the only correct answer, and the digits are compared
    // literally rather than to a tolerance.
    for beyond in [
        "9223372036854775808",  // i64::MAX + 1
        "-9223372036854775809", // i64::MIN - 1
        "18446744073709551616", // u64::MAX + 1
    ] {
        match parse(beyond) {
            CellValue::Integer(v) => panic!("wrapped to Integer({})", v),
            CellValue::String(s) => assert_eq!(s, beyond),
            other => panic!("unexpected {:?} for {}", other, beyond),
        }
    }
}
