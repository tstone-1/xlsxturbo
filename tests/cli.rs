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
