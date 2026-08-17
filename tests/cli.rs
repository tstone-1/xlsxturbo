//! Integration tests for the `xlsxturbo` CLI binary.
//!
//! Cargo builds the binary for integration tests and exposes its path via the
//! `CARGO_BIN_EXE_xlsxturbo` env var, so these drive the real compiled binary
//! without any extra test dependency.

use std::fs;
use std::path::PathBuf;
use std::process::Command;

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_xlsxturbo")
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

#[test]
fn invalid_date_order_exits_nonzero_with_message() {
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

    assert_eq!(output.status.code(), Some(1), "expected exit code 1");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Invalid date_order"),
        "stderr was: {:?}",
        stderr
    );

    let _ = fs::remove_file(&csv);
    let _ = fs::remove_file(&xlsx);
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
    // --verbose echoes the sheet name to stderr, giving us a cheap way to
    // confirm the flag was actually threaded through without adding an
    // xlsx-reading dependency to the test suite.
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Sheet:  MySheet"),
        "stderr was: {:?}",
        stderr
    );
    assert!(xlsx.exists(), "output xlsx was not created");

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
