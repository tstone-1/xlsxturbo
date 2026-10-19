//! Open rust_xlsxwriter defects that xlsxturbo works around, pinned so the
//! workaround is noticed when upstream fixes them.
//!
//! Each test asserts the defect is STILL PRESENT. When one goes red (or stops
//! compiling) after a rust_xlsxwriter bump, the defect is fixed: remove the
//! workaround it names, then delete the test. See
//! `dev-docs/upstream-rust-xlsxwriter.md`, "Upstream defects belong upstream".

use rust_xlsxwriter::{Workbook, XlsxError};

/// jmcnamara/rust_xlsxwriter#213: a failure to create a temporary file is
/// `TempFileError(String)`, which drops the `io::Error`, so
/// `convert::temp_file_failure` cannot set `errno` and a refused temporary
/// folder raises `FileError` rather than `FilePermissionError`.
///
/// When upstream keeps the `io::Error`, this stops matching a bare `String`
/// payload: map it with `FileFailure::from_io` in `temp_file_failure`, add a
/// `FilePermissionError` test for a refused folder, and delete this test.
#[test]
fn issue_213_temp_file_error_carries_only_a_string() {
    let missing = std::env::temp_dir().join("xlsxturbo-upstream-213-does-not-exist");
    assert!(!missing.exists());

    // Control: the same failure, made directly, does carry a kind.
    let io = std::fs::File::create(missing.join("probe")).unwrap_err();
    assert_eq!(io.kind(), std::io::ErrorKind::NotFound);

    match Workbook::new().set_tempdir(&missing) {
        Err(XlsxError::TempFileError(message)) => {
            let message: String = message;
            assert!(
                message.contains("xlsxturbo-upstream-213-does-not-exist"),
                "{message}"
            );
        }
        Err(other) => panic!("#213 may be fixed: set_tempdir now returns {other:?}"),
        Ok(_) => panic!("set_tempdir accepted a folder that does not exist"),
    }
}
