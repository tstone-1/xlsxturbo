//! Calendar dates and times of day without a time zone, and the text patterns
//! that produce them.
//!
//! This module replaced the `chrono` crate, which is no longer maintained
//! (RUSTSEC-2026-0543). xlsxturbo needs three things from a date library: is
//! this year-month-day a real date, which day number is it, and does this text
//! match one of ten fixed patterns. They are written out here so that the
//! result for every input stays what it was: the accepted year range, the
//! leniencies of the text patterns and the arithmetic are the ones the 1.x
//! releases shipped with. Before `chrono` was removed, both implementations
//! ran side by side over every day of the year range and over 80 million
//! generated texts, with no difference; `dev-docs/testing.md` has the method.

/// First and last year a date may have. The range `chrono` accepted.
const MIN_YEAR: i32 = -262_143;
const MAX_YEAR: i32 = 262_142;

/// Day numbers (days since 1970-01-01) of the first and last valid date.
const MIN_UNIX_DAYS: i64 = unix_days_unchecked(MIN_YEAR, 1, 1);
const MAX_UNIX_DAYS: i64 = unix_days_unchecked(MAX_YEAR, 12, 31);

/// Day number of the Excel epoch, 1899-12-30.
const EXCEL_EPOCH_UNIX_DAYS: i64 = unix_days_unchecked(1899, 12, 30);

const SECONDS_PER_DAY: i64 = 86_400;
const NANOS_PER_SECOND: u32 = 1_000_000_000;

/// Days from 1970-01-01 to a date of the proleptic Gregorian calendar.
///
/// The caller has checked that the month is 1..=12 and the day exists. The
/// method counts in 400-year cycles of 146,097 days from a year that starts on
/// 1 March, which puts the leap day at the end of the year.
const fn unix_days_unchecked(year: i32, month: u32, day: u32) -> i64 {
    let year = year as i64 - if month <= 2 { 1 } else { 0 };
    let cycle = year.div_euclid(400);
    let year_of_cycle = year.rem_euclid(400);
    let month_from_march = ((month + 9) % 12) as i64;
    let day_of_year = (153 * month_from_march + 2) / 5 + day as i64 - 1;
    let day_of_cycle = year_of_cycle * 365 + year_of_cycle / 4 - year_of_cycle / 100 + day_of_year;
    cycle * 146_097 + day_of_cycle - 719_468
}

const fn is_leap_year(year: i32) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

const fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    }
}

/// A calendar date, stored as its day number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct NaiveDate {
    unix_days: i64,
}

impl NaiveDate {
    /// A date from year, month and day, or `None` when there is no such date.
    pub(crate) fn from_ymd_opt(year: i32, month: u32, day: u32) -> Option<Self> {
        if !(MIN_YEAR..=MAX_YEAR).contains(&year) || day == 0 || day > days_in_month(year, month) {
            return None;
        }
        Some(Self {
            unix_days: unix_days_unchecked(year, month, day),
        })
    }

    /// The date a number of days after 1970-01-01, or `None` outside the year range.
    pub(crate) fn from_unix_days(unix_days: i64) -> Option<Self> {
        (MIN_UNIX_DAYS..=MAX_UNIX_DAYS)
            .contains(&unix_days)
            .then_some(Self { unix_days })
    }

    /// The date `days` after this one. For tests, which stay inside the year range.
    #[cfg(test)]
    pub(crate) fn plus_days(self, days: i64) -> Self {
        Self::from_unix_days(self.unix_days + days).expect("test date is in range")
    }

    /// Year, month and day of this date.
    #[cfg(test)]
    pub(crate) fn ymd(self) -> (i32, u32, u32) {
        // The inverse of `unix_days_unchecked`.
        let days = self.unix_days + 719_468;
        let cycle = days.div_euclid(146_097);
        let day_of_cycle = days.rem_euclid(146_097);
        let year_of_cycle = (day_of_cycle - day_of_cycle / 1_460 + day_of_cycle / 36_524
            - day_of_cycle / 146_096)
            / 365;
        let day_of_year =
            day_of_cycle - (365 * year_of_cycle + year_of_cycle / 4 - year_of_cycle / 100);
        let month_from_march = (5 * day_of_year + 2) / 153;
        let day = day_of_year - (153 * month_from_march + 2) / 5 + 1;
        let month = if month_from_march < 10 {
            month_from_march + 3
        } else {
            month_from_march - 9
        };
        let year = year_of_cycle + cycle * 400 + i64::from(month <= 2);
        (year as i32, month as u32, day as u32)
    }

    /// This date as `YYYY-MM-DD`.
    #[cfg(test)]
    pub(crate) fn iso(self) -> String {
        let (year, month, day) = self.ymd();
        format!("{year:04}-{month:02}-{day:02}")
    }
}

/// A time of day.
///
/// A leap second is the last second of the day with a nanosecond value of one
/// second or more, so `23:59:60` is `second_of_day` 86,399 and `nanosecond`
/// 1,000,000,000.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct NaiveTime {
    second_of_day: u32,
    nanosecond: u32,
}

impl NaiveTime {
    /// A time from its parts, or `None` when a part is out of range. Only
    /// second 59 may carry a nanosecond value of one second or more.
    pub(crate) fn from_hms_nano_opt(
        hour: u32,
        minute: u32,
        second: u32,
        nanosecond: u32,
    ) -> Option<Self> {
        let nanosecond_limit = if second == 59 {
            2 * NANOS_PER_SECOND
        } else {
            NANOS_PER_SECOND
        };
        if hour >= 24 || minute >= 60 || second >= 60 || nanosecond >= nanosecond_limit {
            return None;
        }
        Some(Self {
            second_of_day: hour * 3_600 + minute * 60 + second,
            nanosecond,
        })
    }
}

/// A date with a time of day.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct NaiveDateTime {
    date: NaiveDate,
    time: NaiveTime,
}

impl NaiveDateTime {
    pub(crate) fn new(date: NaiveDate, time: NaiveTime) -> Self {
        Self { date, time }
    }

    /// The UTC date and time of a count of seconds since 1970-01-01T00:00:00,
    /// or `None` outside the year range.
    pub(crate) fn from_unix_seconds(seconds: i64, nanosecond: u32) -> Option<Self> {
        if nanosecond >= NANOS_PER_SECOND {
            return None;
        }
        let date = NaiveDate::from_unix_days(seconds.div_euclid(SECONDS_PER_DAY))?;
        let time = NaiveTime {
            second_of_day: seconds.rem_euclid(SECONDS_PER_DAY) as u32,
            nanosecond,
        };
        Some(Self { date, time })
    }
}

/// Convert NaiveDate to Excel serial date number.
///
/// The epoch used here (1899-12-30) only produces the correct Excel serial
/// for dates on or after 1900-03-01 (serial 61). Excel's serial numbering
/// assumes a phantom leap day, 1900-02-29, that never actually existed, so
/// for any date before 1900-03-01 this formula returns a value one day
/// ahead of what Excel would show. Callers that accept arbitrary dates
/// must reject results below serial 61 rather than write them as dates.
pub(crate) fn naive_date_to_excel(date: NaiveDate) -> f64 {
    (date.unix_days - EXCEL_EPOCH_UNIX_DAYS) as f64
}

/// Convert NaiveDateTime to Excel serial datetime number
pub(crate) fn naive_datetime_to_excel(dt: NaiveDateTime) -> f64 {
    let date_part = naive_date_to_excel(dt.date);
    let seconds = dt.time.second_of_day as f64;
    let fractional_seconds = dt.time.nanosecond as f64 / 1_000_000_000.0;
    let time_fraction = (seconds + fractional_seconds) / 86400.0;
    date_part + time_fraction
}

/// The numbers one pattern found in a text.
#[derive(Default)]
struct Fields {
    year: Option<i64>,
    month: Option<i64>,
    day: Option<i64>,
    hour: Option<i64>,
    minute: Option<i64>,
    second: Option<i64>,
    nanosecond: Option<u32>,
}

/// Read 1 to `max_digits` ASCII digits from the start of `text`.
///
/// Returns the rest of the text and the number, or `None` when the text does
/// not start with a digit or the number does not fit `i64`.
fn take_number(text: &str, max_digits: usize) -> Option<(&str, i64)> {
    let mut value = 0i64;
    let mut used = 0;
    for byte in text.bytes().take(max_digits) {
        if !byte.is_ascii_digit() {
            break;
        }
        value = value.checked_mul(10)?.checked_add(i64::from(byte - b'0'))?;
        used += 1;
    }
    (used > 0).then(|| (&text[used..], value))
}

/// Match `text` against one pattern and return the numbers it holds.
///
/// A pattern is made of `%Y %m %d %H %M %S`, `%.f`, single characters that must
/// match, and spaces. The leniencies are part of the 1.x output, because they
/// decide which CSV cells become dates:
///
/// - White space, of any kind and length, is skipped before every number and
///   where the pattern has a space.
/// - A number has one digit or more. `%Y` takes up to four digits, or any
///   count after a `+` or `-`; the others take up to two.
/// - `%.f` matches nothing when the text has no `.` there. After a `.` it
///   needs a digit, uses nine and skips the rest.
/// - The whole text has to be used.
fn match_pattern(text: &str, pattern: &str) -> Option<Fields> {
    let mut fields = Fields::default();
    let mut rest = text;
    let mut spec = pattern.chars();
    while let Some(c) = spec.next() {
        if c.is_whitespace() {
            rest = rest.trim_start();
            continue;
        }
        if c != '%' {
            rest = rest.strip_prefix(c)?;
            continue;
        }
        let kind = spec.next()?;
        if kind == '.' {
            // "%.f": an optional fraction of a second.
            spec.next().filter(|&f| f == 'f')?;
            if let Some(digits) = rest.strip_prefix('.') {
                let (after, value) = take_number(digits, 9)?;
                let used = digits.len() - after.len();
                fields.nanosecond = Some(value as u32 * 10u32.pow(9 - used as u32));
                rest = after.trim_start_matches(|d: char| d.is_ascii_digit());
            }
            continue;
        }
        rest = rest.trim_start();
        let (after, value) = if kind == 'Y' {
            if let Some(digits) = rest.strip_prefix('-') {
                let (after, value) = take_number(digits, usize::MAX)?;
                (after, -value)
            } else if let Some(digits) = rest.strip_prefix('+') {
                take_number(digits, usize::MAX)?
            } else {
                take_number(rest, 4)?
            }
        } else {
            take_number(rest, 2)?
        };
        let field = match kind {
            'Y' => &mut fields.year,
            'm' => &mut fields.month,
            'd' => &mut fields.day,
            'H' => &mut fields.hour,
            'M' => &mut fields.minute,
            'S' => &mut fields.second,
            _ => return None,
        };
        *field = Some(value);
        rest = after;
    }
    rest.is_empty().then_some(fields)
}

fn date_from_fields(fields: &Fields) -> Option<NaiveDate> {
    NaiveDate::from_ymd_opt(
        i32::try_from(fields.year?).ok()?,
        u32::try_from(fields.month?).ok()?,
        u32::try_from(fields.day?).ok()?,
    )
}

/// Parse `text` as a date with one of the date patterns.
pub(crate) fn parse_date(text: &str, pattern: &str) -> Option<NaiveDate> {
    date_from_fields(&match_pattern(text, pattern)?)
}

/// Parse `text` as a date and time with one of the datetime patterns.
///
/// Second 60 is accepted as a leap second.
pub(crate) fn parse_datetime(text: &str, pattern: &str) -> Option<NaiveDateTime> {
    let fields = match_pattern(text, pattern)?;
    let date = date_from_fields(&fields)?;
    let hour = u32::try_from(fields.hour?).ok()?;
    let minute = u32::try_from(fields.minute?).ok()?;
    let nanosecond = fields.nanosecond.unwrap_or(0);
    let (second, nanosecond) = match fields.second? {
        60 => (59, nanosecond + NANOS_PER_SECOND),
        second => (u32::try_from(second).ok()?, nanosecond),
    };
    let time = NaiveTime::from_hms_nano_opt(hour, minute, second, nanosecond)?;
    Some(NaiveDateTime::new(date, time))
}

#[cfg(test)]
mod tests {
    //! The rules of this module, one test per rule. They were found by running
    //! the module against `chrono` before that crate was removed; each test
    //! fails when its rule is changed.
    use super::*;

    const ISO: &str = "%Y-%m-%d";
    const ISO_DATETIME: &str = "%Y-%m-%dT%H:%M:%S%.f";

    fn date(year: i32, month: u32, day: u32) -> NaiveDate {
        NaiveDate::from_ymd_opt(year, month, day).expect("test date is valid")
    }

    fn serial(text: &str) -> Option<f64> {
        parse_datetime(text, ISO_DATETIME).map(naive_datetime_to_excel)
    }

    #[test]
    fn leap_years_follow_the_gregorian_rule() {
        assert!(NaiveDate::from_ymd_opt(2024, 2, 29).is_some());
        assert!(NaiveDate::from_ymd_opt(2023, 2, 29).is_none());
        // Divisible by 100 is not a leap year, divisible by 400 is.
        assert!(NaiveDate::from_ymd_opt(1900, 2, 29).is_none());
        assert!(NaiveDate::from_ymd_opt(2100, 2, 29).is_none());
        assert!(NaiveDate::from_ymd_opt(2000, 2, 29).is_some());
    }

    #[test]
    fn a_date_needs_a_real_month_and_day() {
        for (year, month, day) in [
            (2024, 0, 1),
            (2024, 13, 1),
            (2024, 1, 0),
            (2024, 1, 32),
            (2024, 4, 31),
            (2024, 2, 30),
        ] {
            assert!(
                NaiveDate::from_ymd_opt(year, month, day).is_none(),
                "{year}-{month}-{day}"
            );
        }
    }

    #[test]
    fn the_year_range_is_the_one_the_1x_releases_had() {
        assert!(NaiveDate::from_ymd_opt(262_142, 12, 31).is_some());
        assert!(NaiveDate::from_ymd_opt(262_143, 1, 1).is_none());
        assert!(NaiveDate::from_ymd_opt(-262_143, 1, 1).is_some());
        assert!(NaiveDate::from_ymd_opt(-262_144, 12, 31).is_none());
        // The same limits for a count of seconds.
        let last_second = (MAX_UNIX_DAYS + 1) * SECONDS_PER_DAY - 1;
        assert!(NaiveDateTime::from_unix_seconds(last_second, 999_999_999).is_some());
        assert!(NaiveDateTime::from_unix_seconds(last_second + 1, 0).is_none());
        let first_second = MIN_UNIX_DAYS * SECONDS_PER_DAY;
        assert!(NaiveDateTime::from_unix_seconds(first_second, 0).is_some());
        assert!(NaiveDateTime::from_unix_seconds(first_second - 1, 0).is_none());
    }

    #[test]
    fn day_numbers_are_consecutive_and_convert_back() {
        // Every day of 4,000 years: the next date has the next day number, and
        // the day number gives the date back.
        let mut expected = date(1, 1, 1).unix_days;
        for year in 1..=4_000 {
            for month in 1..=12 {
                for day in 1..=days_in_month(year, month) {
                    let d = date(year, month, day);
                    assert_eq!(d.unix_days, expected, "{year}-{month}-{day}");
                    assert_eq!(d.ymd(), (year, month, day));
                    expected += 1;
                }
            }
        }
        // 4,000 years are ten 400-year cycles.
        assert_eq!(expected - date(1, 1, 1).unix_days, 10 * 146_097);
        assert_eq!(date(1970, 1, 1).unix_days, 0);
        assert_eq!(date(-262_143, 1, 1).ymd(), (-262_143, 1, 1));
        assert_eq!(date(262_142, 12, 31).ymd(), (262_142, 12, 31));
    }

    #[test]
    fn seconds_before_1970_count_back_from_midnight() {
        // One second before the epoch is the last second of 1969-12-31, not
        // the first second of 1970-01-01 with a negative time.
        let dt = NaiveDateTime::from_unix_seconds(-1, 0).expect("in range");
        let expected = NaiveDateTime::new(
            date(1969, 12, 31),
            NaiveTime::from_hms_nano_opt(23, 59, 59, 0).expect("valid time"),
        );
        assert_eq!(dt, expected);
        // A whole second does not fit the nanosecond part.
        assert!(NaiveDateTime::from_unix_seconds(0, 999_999_999).is_some());
        assert!(NaiveDateTime::from_unix_seconds(0, 1_000_000_000).is_none());
        assert_eq!(naive_datetime_to_excel(dt), 25_568.0 + 86_399.0 / 86_400.0);
    }

    #[test]
    fn only_second_59_may_hold_a_leap_second() {
        assert!(NaiveTime::from_hms_nano_opt(23, 59, 59, 1_999_999_999).is_some());
        assert!(NaiveTime::from_hms_nano_opt(23, 59, 59, 2_000_000_000).is_none());
        assert!(NaiveTime::from_hms_nano_opt(23, 59, 58, 1_000_000_000).is_none());
        assert!(NaiveTime::from_hms_nano_opt(24, 0, 0, 0).is_none());
        assert!(NaiveTime::from_hms_nano_opt(0, 60, 0, 0).is_none());
        assert!(NaiveTime::from_hms_nano_opt(0, 0, 60, 0).is_none());
    }

    #[test]
    fn a_number_has_one_or_two_digits() {
        assert_eq!(parse_date("2024-1-5", ISO), Some(date(2024, 1, 5)));
        assert_eq!(parse_date("2024-01-05", ISO), Some(date(2024, 1, 5)));
        // A third digit is not part of the number, and nothing else may follow.
        assert_eq!(parse_date("2024-011-05", ISO), None);
        assert_eq!(parse_date("2024-01-005", ISO), None);
        assert_eq!(parse_date("2024--05", ISO), None);
    }

    #[test]
    fn a_year_has_up_to_four_digits_or_a_sign() {
        assert_eq!(parse_date("24-01-15", ISO), Some(date(24, 1, 15)));
        assert_eq!(parse_date("12345-01-15", ISO), None);
        assert_eq!(parse_date("+12345-01-15", ISO), Some(date(12_345, 1, 15)));
        assert_eq!(parse_date("+2024-01-15", ISO), Some(date(2024, 1, 15)));
        assert_eq!(parse_date("-0044-03-15", ISO), Some(date(-44, 3, 15)));
        // The sign belongs to the year only.
        assert_eq!(parse_date("2024-+1-15", ISO), None);
        assert_eq!(parse_date("2024-01--15", ISO), None);
        // A year past the range, and one that does not fit the integer type.
        assert_eq!(parse_date("+262143-01-01", ISO), None);
        assert_eq!(parse_date("+99999999999999999999-01-01", ISO), None);
    }

    #[test]
    fn white_space_is_skipped_before_a_number() {
        assert_eq!(parse_date("2024- 1-15", ISO), Some(date(2024, 1, 15)));
        assert_eq!(parse_date("2024-01-\t15", ISO), Some(date(2024, 1, 15)));
        assert_eq!(parse_date("2024-01-\u{a0}15", ISO), Some(date(2024, 1, 15)));
        // Not before a separator, and not at the end.
        assert_eq!(parse_date("2024 -01-15", ISO), None);
        assert_eq!(parse_date("2024-01-15 ", ISO), None);
    }

    #[test]
    fn a_space_in_the_pattern_matches_any_white_space_or_none() {
        let pattern = "%Y-%m-%d %H:%M:%S";
        let expected = parse_datetime("2024-01-15 10:30:00", pattern);
        assert!(expected.is_some());
        assert_eq!(parse_datetime("2024-01-15   10:30:00", pattern), expected);
        assert_eq!(parse_datetime("2024-01-15\t10:30:00", pattern), expected);
        // None at all is enough: the day takes two digits and the hour follows.
        assert_eq!(parse_datetime("2024-01-1510:30:00", pattern), expected);
        // The "T" of the other pattern is a character that has to be there.
        assert_eq!(parse_datetime("2024-01-15 10:30:00", ISO_DATETIME), None);
    }

    #[test]
    fn the_whole_text_has_to_match() {
        assert_eq!(parse_date("2024-01-15x", ISO), None);
        assert_eq!(parse_date("2024-01-15T10:30:00", ISO), None);
        assert_eq!(parse_date("2024-01", ISO), None);
        assert_eq!(parse_date("", ISO), None);
        assert_eq!(serial("2024-01-15T10:30:00Z"), None);
        assert_eq!(serial("2024-01-15T10:30"), None);
    }

    #[test]
    fn a_fraction_is_optional_and_uses_nine_digits() {
        let whole = serial("2024-01-15T10:30:00").expect("a datetime");
        let half = serial("2024-01-15T10:30:00.5").expect("a datetime");
        assert_eq!(half, serial("2024-01-15T10:30:00.500").expect("a datetime"));
        assert!(half > whole);
        // Digits after the ninth are dropped, not rounded.
        assert_eq!(
            serial("2024-01-15T10:30:00.1234567899999"),
            serial("2024-01-15T10:30:00.123456789")
        );
        // A "." needs a digit, and no other character starts a fraction.
        assert_eq!(serial("2024-01-15T10:30:00."), None);
        assert_eq!(serial("2024-01-15T10:30:00,5"), None);
        assert_eq!(serial("2024-01-15T10:30:00. 5"), None);
    }

    #[test]
    fn second_60_is_the_next_minute() {
        let leap = serial("2016-12-31T23:59:60").expect("a datetime");
        assert_eq!(leap, naive_date_to_excel(date(2017, 1, 1)));
        assert_eq!(serial("2024-01-15T10:30:60"), serial("2024-01-15T10:31:00"));
        assert_eq!(serial("2024-01-15T10:30:61"), None);
        assert_eq!(serial("2024-01-15T24:00:00"), None);
        assert_eq!(serial("2024-01-15T10:60:00"), None);
    }

    #[test]
    fn every_pattern_in_use_is_one_this_module_reads() {
        use crate::types::{DateOrder, DATETIME_PATTERNS};
        // A pattern with a field this module does not know never matches, so
        // a new pattern has to be checked here.
        for pattern in DATETIME_PATTERNS {
            let text = pattern
                .replace("%Y", "2024")
                .replace("%m", "03")
                .replace("%d", "15")
                .replace("%H", "10")
                .replace("%M", "30")
                .replace("%S", "45")
                .replace("%.f", ".25");
            assert!(parse_datetime(&text, pattern).is_some(), "{pattern}");
        }
        for order in [DateOrder::Auto, DateOrder::DMY, DateOrder::MDY] {
            for pattern in order.patterns() {
                let text = pattern
                    .replace("%Y", "2024")
                    .replace("%m", "03")
                    .replace("%d", "15");
                assert_eq!(
                    parse_date(&text, pattern),
                    Some(date(2024, 3, 15)),
                    "{pattern}"
                );
            }
        }
    }
}
