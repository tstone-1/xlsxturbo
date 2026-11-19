//! Integration tests for the `xlsxturbo` CLI binary.
//!
//! Cargo builds the binary for integration tests and exposes its path via the
//! `CARGO_BIN_EXE_xlsxturbo` env var, so these drive the real compiled binary.
//! The `zip` dev-dependency reads the parts of the workbook it writes, so a
//! flag is checked by what reaches the file, not by an exit code.

use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Command;

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_xlsxturbo")
}

/// The text of one part of the workbook at `xlsx`.
fn read_part(xlsx: &Path, part: &str) -> String {
    let file = fs::File::open(xlsx).expect("output xlsx exists");
    let mut archive = zip::ZipArchive::new(file).expect("output is a zip archive");
    let mut entry = archive
        .by_name(part)
        .unwrap_or_else(|e| panic!("{part} missing: {e}"));
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("part is UTF-8");
    text
}

/// The `<c>` element of `cell` in a worksheet part, start tag to end tag.
fn cell_xml<'a>(sheet: &'a str, cell: &str) -> &'a str {
    let open = format!("<c r=\"{cell}\"");
    let start = sheet
        .find(&open)
        .unwrap_or_else(|| panic!("no cell {cell} in {sheet}"));
    let rest = &sheet[start..];
    let end = rest.find("</c>").map(|i| i + 4).unwrap_or_else(|| {
        rest.find("/>")
            .map(|i| i + 2)
            .expect("cell element is closed")
    });
    &rest[..end]
}

/// The `<v>` text of `cell`, or `None` when the cell has no value element.
fn cell_value<'a>(sheet: &'a str, cell: &str) -> Option<&'a str> {
    let xml = cell_xml(sheet, cell);
    let start = xml.find("<v>")? + 3;
    let end = xml[start..].find("</v>")? + start;
    Some(&xml[start..end])
}

/// Run the CLI over `csv` into `xlsx` with `extra` arguments; assert it succeeded.
fn convert(csv: &Path, xlsx: &Path, extra: &[&str]) -> String {
    let output = Command::new(bin())
        .arg(csv)
        .arg(xlsx)
        .args(extra)
        .output()
        .expect("failed to run xlsxturbo binary");
    assert!(
        output.status.success(),
        "{extra:?}: expected exit 0, got {:?}, stderr {:?}",
        output.status,
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8_lossy(&output.stdout).into_owned()
}

/// Unique temp path for this process + tag, so parallel tests don't collide.
fn temp_path(tag: &str, ext: &str) -> PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!(
        "xlsxturbo_cli_{}_{}.{}",
        std::process::id(),
        tag,
        ext
    ));
    p
}

#[test]
fn valid_csv_exits_zero_and_prints_ok() {
    let csv = temp_path("ok", "csv");
    let xlsx = temp_path("ok", "xlsx");
    fs::write(&csv, "a,b,c\n1,2,3\n4,5,6\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .output()
        .expect("failed to run xlsxturbo binary");

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        output.status.success(),
        "expected exit 0, got {:?}",
        output.status
    );
    // Contract: prints "OK {rows} {cols}".
    assert!(stdout.starts_with("OK "), "stdout was: {:?}", stdout);
    assert!(stdout.trim().ends_with("3 3"), "stdout was: {:?}", stdout);
    assert!(xlsx.exists(), "output xlsx was not created");

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

#[test]
fn missing_input_file_exits_nonzero() {
    let missing = temp_path("does_not_exist", "csv");
    let xlsx = temp_path("missing", "xlsx");
    let _ = fs::remove_file(&missing); // ensure absent

    let output = Command::new(bin())
        .arg(&missing)
        .arg(&xlsx)
        .output()
        .expect("failed to run xlsxturbo binary");

    assert_eq!(output.status.code(), Some(1), "expected exit code 1");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("Error"), "stderr was: {:?}", stderr);
    assert!(!xlsx.exists(), "no xlsx should be written on failure");

    let _ = fs::remove_file(&xlsx);
}

/// A bad `--date-order` is a usage error, so it exits 2 — the value clap
/// already uses for the usage errors it rejects itself. It exited 1 until
/// 1.3.x, which made it indistinguishable from a conversion that was asked for
/// correctly and failed, and left the binary disagreeing with its own argument
/// parser about which kind of mistake had been made.
#[test]
fn invalid_date_order_is_a_usage_error() {
    let csv = temp_path("baddate", "csv");
    let xlsx = temp_path("baddate", "xlsx");
    fs::write(&csv, "a\n1\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .arg("--date-order")
        .arg("nonsense")
        .output()
        .expect("failed to run xlsxturbo binary");

    assert_eq!(
        output.status.code(),
        Some(2),
        "expected exit code 2 (usage)"
    );
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Invalid date_order"),
        "stderr was: {:?}",
        stderr
    );

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

/// The control for the test above, and the reason 2 is the right number: clap
/// picks it for an argument *it* rejects. Without this, "our usage errors exit
/// 2" is a value chosen by taste rather than one that agrees with the parser
/// sitting in front of it — and a clap major that changed the convention would
/// silently leave the two halves disagreeing again.
#[test]
fn clap_rejects_an_unknown_flag_with_the_same_usage_code() {
    let output = Command::new(bin())
        .arg("--no-such-flag")
        .output()
        .expect("failed to run xlsxturbo binary");

    assert_eq!(
        output.status.code(),
        Some(2),
        "clap's own usage exit code moved; the hand-rolled one in src/main.rs \
         was chosen to match it, so the two now disagree"
    );
}

#[test]
fn parallel_flag_exits_zero_and_produces_file() {
    let csv = temp_path("parallel", "csv");
    let xlsx = temp_path("parallel", "xlsx");
    fs::write(&csv, "a,b,c\n1,2,3\n4,5,6\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .arg("--parallel")
        .output()
        .expect("failed to run xlsxturbo binary");

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        output.status.success(),
        "expected exit 0, got {:?}",
        output.status
    );
    // Contract: prints "OK {rows} {cols}", same as the non-parallel path.
    assert!(stdout.starts_with("OK "), "stdout was: {:?}", stdout);
    assert!(stdout.trim().ends_with("3 3"), "stdout was: {:?}", stdout);
    assert!(xlsx.exists(), "output xlsx was not created");

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

/// `--parallel` writes the same cells as the sequential run: numbers, text,
/// a date and an empty field, each cell compared by its whole element.
#[test]
fn parallel_run_writes_the_same_cells_as_the_sequential_run() {
    let csv = temp_path("par_same", "csv");
    let seq = temp_path("par_same_seq", "xlsx");
    let par = temp_path("par_same_par", "xlsx");
    fs::write(&csv, "n,s,d\n1,x,2024-03-15\n2.5,,2024-03-16\n-3,y z,\n").unwrap();

    let seq_out = convert(&csv, &seq, &[]);
    let par_out = convert(&csv, &par, &["--parallel"]);
    assert_eq!(seq_out, par_out);

    let (seq_sheet, par_sheet) = (
        read_part(&seq, "xl/worksheets/sheet1.xml"),
        read_part(&par, "xl/worksheets/sheet1.xml"),
    );
    for cell in ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "C3", "A4", "B4"] {
        assert_eq!(
            cell_xml(&seq_sheet, cell),
            cell_xml(&par_sheet, cell),
            "cell {cell}"
        );
    }
    // The values themselves, so two runs that agree on a wrong value fail too.
    assert_eq!(cell_value(&par_sheet, "A2"), Some("1"));
    assert_eq!(cell_value(&par_sheet, "A3"), Some("2.5"));
    assert_eq!(cell_value(&par_sheet, "A4"), Some("-3"));
    assert_eq!(cell_value(&par_sheet, "C2"), Some("45366"));
    // A cell's `s` is an index into the styles in order of first use, so the
    // same index can name another number format: compare the styles too.
    for part in ["xl/sharedStrings.xml", "xl/styles.xml"] {
        assert_eq!(read_part(&seq, part), read_part(&par, part), "{part}");
    }

    for path in [&csv, &seq, &par] {
        let _ = fs::remove_file(path);
    }
}

/// A CSV that `flexible(true)` genuinely refuses, plus the 1-based row it fails on.
///
/// Picking the input is the whole difficulty: the reader is deliberately
/// permissive, so ragged rows are *accepted* (`flexible(true)`) and an unclosed
/// quote is read to EOF rather than rejected. Invalid UTF-8 is what actually
/// fails, because `records()` yields `StringRecord`. The bad byte sits on the
/// third line, so a 1-based report must say 3 and the old 0-based one said 2.
fn write_malformed_csv(path: &PathBuf) -> usize {
    let mut bytes: Vec<u8> = b"a,b\n1,2\nx,".to_vec();
    bytes.push(0xFF); // not valid UTF-8 in any position
    bytes.extend_from_slice(b"y\n");
    fs::write(path, bytes).unwrap();
    3
}

/// Run the CLI over `csv`, returning its stderr.
fn run_cli_expecting_failure(csv: &PathBuf, xlsx: &PathBuf, parallel: bool) -> String {
    let mut command = Command::new(bin());
    command.arg(csv).arg(xlsx);
    if parallel {
        command.arg("--parallel");
    }
    let output = command.output().expect("failed to run xlsxturbo binary");
    assert_eq!(
        output.status.code(),
        Some(1),
        "malformed CSV should exit 1 ({}), stderr was: {:?}",
        if parallel { "parallel" } else { "sequential" },
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8_lossy(&output.stderr).into_owned()
}

/// Both CSV pipelines must name the same row, and count it from 1.
///
/// This is the guard the duplication never had. The row number used to be
/// formatted at two sites from two different expressions -- `row_count` in the
/// sequential loop, `row_count + chunk.len()` in the chunked one -- which agreed
/// only by arithmetic coincidence: nothing failed if the chunked path's
/// bookkeeping changed, and a user comparing the two would have been told a
/// different row for the same file. Both now go through one helper, and this
/// test is what notices if they stop.
///
/// It asserts the *value*, not just agreement: two paths reporting the same
/// wrong number agree perfectly.
#[test]
fn both_csv_paths_report_the_same_one_based_row_for_a_parse_error() {
    let csv = temp_path("badrow", "csv");
    let seq_xlsx = temp_path("badrow_seq", "xlsx");
    let par_xlsx = temp_path("badrow_par", "xlsx");
    let bad_line = write_malformed_csv(&csv);

    let sequential = run_cli_expecting_failure(&csv, &seq_xlsx, false);
    let parallel = run_cli_expecting_failure(&csv, &par_xlsx, true);

    let expected = format!("CSV parse error at row {}:", bad_line);
    assert!(
        sequential.contains(&expected),
        "sequential path should report row {}, stderr was: {:?}",
        bad_line,
        sequential
    );
    assert!(
        parallel.contains(&expected),
        "parallel path should report row {}, stderr was: {:?}",
        bad_line,
        parallel
    );
    // The 0-based report this replaced. Named explicitly so a silent slide back
    // to it fails here rather than only in a user's confusion.
    let zero_based = format!("CSV parse error at row {}:", bad_line - 1);
    assert!(
        !sequential.contains(&zero_based) && !parallel.contains(&zero_based),
        "row numbers must be 1-based; sequential: {:?}, parallel: {:?}",
        sequential,
        parallel
    );
    assert_eq!(
        sequential, parallel,
        "the two CSV pipelines must report the identical error for the identical file"
    );

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&seq_xlsx);
    let _ = fs::remove_file(&par_xlsx);
}

/// Control for the test above: the same fixture *without* the bad byte converts.
///
/// Without this, a CLI that failed on every input -- a broken binary, a path
/// mistake, an argument the parser rejects -- would satisfy every assertion up
/// there, since they only ever check that a failure happened and what it said.
#[test]
fn the_malformed_csv_fixture_is_valid_once_the_bad_byte_is_removed() {
    let csv = temp_path("badrow_control", "csv");
    let xlsx = temp_path("badrow_control", "xlsx");
    fs::write(&csv, "a,b\n1,2\nx,y\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .output()
        .expect("failed to run xlsxturbo binary");

    assert!(
        output.status.success(),
        "the fixture must fail only on the invalid byte, but the clean version failed: {:?}",
        String::from_utf8_lossy(&output.stderr)
    );

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

#[test]
fn sheet_name_flag_is_respected() {
    let csv = temp_path("sheetname", "csv");
    let xlsx = temp_path("sheetname", "xlsx");
    fs::write(&csv, "a,b\n1,2\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .arg("--sheet-name")
        .arg("MySheet")
        .arg("--verbose")
        .output()
        .expect("failed to run xlsxturbo binary");

    assert!(
        output.status.success(),
        "expected exit 0, got {:?}",
        output.status
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    // The CLI CSV reader treats every line as data (no header row), so "a,b\n1,2\n"
    // is 2 rows x 2 cols.
    assert!(stdout.trim().ends_with("2 2"), "stdout was: {:?}", stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Sheet:  MySheet"),
        "stderr was: {:?}",
        stderr
    );
    let workbook = read_part(&xlsx, "xl/workbook.xml");
    assert!(
        workbook.contains("<sheet name=\"MySheet\" sheetId=\"1\""),
        "workbook.xml was: {workbook}"
    );
    assert!(!workbook.contains("Sheet1"), "workbook.xml was: {workbook}");

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

#[test]
fn version_flag_prints_crate_version() {
    let output = Command::new(bin())
        .arg("--version")
        .output()
        .expect("failed to run xlsxturbo binary");

    assert!(
        output.status.success(),
        "expected exit 0, got {:?}",
        output.status
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let expected = format!("xlsxturbo {}", env!("CARGO_PKG_VERSION"));
    assert!(
        stdout.trim() == expected,
        "stdout was: {:?}, expected: {:?}",
        stdout,
        expected
    );
}

#[test]
fn non_default_date_order_exits_zero() {
    let csv = temp_path("dmy", "csv");
    let xlsx = temp_path("dmy", "xlsx");
    fs::write(&csv, "a\n01-02-2024\n").unwrap();

    let output = Command::new(bin())
        .arg(&csv)
        .arg(&xlsx)
        .arg("--date-order")
        .arg("dmy")
        .output()
        .expect("failed to run xlsxturbo binary");

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        output.status.success(),
        "expected exit 0, got {:?}",
        output.status
    );
    assert!(stdout.starts_with("OK "), "stdout was: {:?}", stdout);
    // The CLI CSV reader treats every line as data (no header row), so "a\n01-02-2024\n"
    // is 2 rows x 1 col.
    assert!(stdout.trim().ends_with("2 1"), "stdout was: {:?}", stdout);
    assert!(xlsx.exists(), "output xlsx was not created");

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
}

/// `--date-order` decides what `01-02-2024` means, on both CSV paths: 1 February
/// (serial 45323) for `dmy`, 2 January (serial 45293) for `mdy`.
#[test]
fn date_order_changes_the_serial() {
    let csv = temp_path("dateorder", "csv");
    fs::write(&csv, "a\n01-02-2024\n").unwrap();

    for parallel in [false, true] {
        let mut serials = Vec::new();
        for (order, expected) in [("dmy", "45323"), ("mdy", "45293")] {
            let xlsx = temp_path(&format!("dateorder_{order}_{parallel}"), "xlsx");
            let mut extra = vec!["--date-order", order];
            if parallel {
                extra.push("--parallel");
            }
            convert(&csv, &xlsx, &extra);
            let sheet = read_part(&xlsx, "xl/worksheets/sheet1.xml");
            let serial = cell_value(&sheet, "A2").map(str::to_owned);
            assert_eq!(
                serial.as_deref(),
                Some(expected),
                "{order}, parallel {parallel}"
            );
            serials.push(serial);
            let _ = fs::remove_file(&xlsx);
        }
        assert_ne!(serials[0], serials[1], "parallel {parallel}");
    }

    let _ = fs::remove_file(&csv);
}

/// `--compression-level` reaches the writer: level 1 gives a larger file than 9.
#[test]
fn compression_level_flag_changes_the_file_size() {
    let csv = temp_path("level", "csv");
    let fast = temp_path("level1", "xlsx");
    let small = temp_path("level9", "xlsx");
    let rows: String = (0..5000)
        .map(|i| format!("{},row {}\n", i, i % 50))
        .collect();
    fs::write(&csv, format!("n,s\n{rows}")).unwrap();

    for (path, level) in [(&fast, "1"), (&small, "9")] {
        let output = Command::new(bin())
            .arg(&csv)
            .arg(path)
            .arg("--compression-level")
            .arg(level)
            .output()
            .expect("failed to run xlsxturbo binary");
        assert!(output.status.success(), "level {level}: {:?}", output);
    }
    let (fast_len, small_len) = (
        fs::metadata(&fast).unwrap().len(),
        fs::metadata(&small).unwrap().len(),
    );
    assert!(
        fast_len > small_len,
        "level 1: {fast_len} bytes, level 9: {small_len}"
    );

    for path in [&csv, &fast, &small] {
        let _ = fs::remove_file(path);
    }
}

/// A level outside 1 to 9 is a usage error, exit 2, as in the Python command.
#[test]
fn invalid_compression_level_is_a_usage_error() {
    for level in ["0", "10", "x"] {
        let output = Command::new(bin())
            .arg("in.csv")
            .arg("out.xlsx")
            .arg("-c")
            .arg(level)
            .output()
            .expect("failed to run xlsxturbo binary");
        assert_eq!(output.status.code(), Some(2), "level {level}");
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(stderr.contains("from 1 to 9"), "stderr was: {:?}", stderr);
    }
}
