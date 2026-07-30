//! Shared workbook-level helpers.

use crate::errors::{errno, FileFailure};
use rust_xlsxwriter::{Workbook, XlsxError};
use std::collections::HashMap;
use std::path::Path;
use tempfile::NamedTempFile;

/// Give the staged file the permissions the finished export should have.
///
/// A `NamedTempFile` is created 0600, so persisting it unchanged would make
/// exports less readable than the 0644-style file `Workbook::save`'s
/// `File::create` used to produce — enough to break a shared-drive workflow.
/// When replacing an existing export we reuse its mode; otherwise we fall back
/// to the usual data-file default. Best effort: a failure here affects only
/// permissions, never contents, so it must not fail the export.
///
/// Note this does not consult the process umask (std exposes no race-free way
/// to read it), so a new file under an unusually strict umask comes out more
/// permissive than `File::create` would have made it. Replacing an existing
/// file — the case this whole function exists for — is exact.
#[cfg(unix)]
fn set_output_permissions(tmp: &NamedTempFile, dest: &Path) {
    use std::fs::Permissions;
    use std::os::unix::fs::PermissionsExt;

    let mode = std::fs::metadata(dest)
        .map(|m| m.permissions().mode() & 0o777)
        .unwrap_or(0o644);
    let _ = tmp.as_file().set_permissions(Permissions::from_mode(mode));
}

/// Windows counterpart: the staged file inherits the destination directory's
/// inheritable ACEs and `persist` carries them across the rename, so there is
/// nothing to fix up.
#[cfg(not(unix))]
fn set_output_permissions(_tmp: &NamedTempFile, _dest: &Path) {}

/// Save a workbook so the destination is only ever replaced by a complete file.
///
/// `Workbook::save` calls `File::create` — which truncates the destination —
/// *before* it serializes and validates anything. Any later failure (a chart
/// range naming a sheet that does not exist, a full disk, a dropped network
/// share) therefore left a 0-byte file where the user's previous export was,
/// destroying it as a side effect of an error that was otherwise reported
/// cleanly. Staging into a temporary file and renaming on success keeps the
/// destination as either the old file or the new one, never a stub.
///
/// The staging file is created in the destination's own directory rather than
/// `$TMPDIR` so the rename stays within one filesystem and is therefore atomic;
/// a cross-filesystem rename would degrade into a copy and reintroduce the
/// partial-write window.
pub(crate) fn save_workbook(workbook: &mut Workbook, output_path: &str) -> Result<(), FileFailure> {
    let dest = Path::new(output_path);
    let dir = match dest.parent() {
        Some(parent) if !parent.as_os_str().is_empty() => parent,
        _ => Path::new("."),
    };

    // Every failure below reports as a save failure: staging is an
    // implementation detail, and a caller who passed a bad path should not have
    // to reason about temporary files to understand what went wrong.
    let context = format!("Failed to save workbook to '{}'", output_path);

    if !dir.is_dir() {
        // `ENOENT` is synthesised rather than observed, because this branch
        // exists to produce a better message than the syscall would. It is the
        // number the syscall one line below *would* have returned, and it is
        // worth setting: a missing output directory is the most common file
        // failure here, so leaving it unnumbered would make `errno` useless
        // exactly where a caller most wants to branch on it.
        return Err(FileFailure::detected(
            format!("{}: directory '{}' does not exist", context, dir.display()),
            errno::ENOENT,
        ));
    }

    let mut tmp =
        NamedTempFile::new_in(dir).map_err(|e| FileFailure::from_io(context.clone(), &e))?;

    workbook
        .save_to_writer(tmp.as_file_mut())
        .map_err(|e| match &e {
            // `save_to_writer` wraps the underlying `io::Error`, so a disk-full
            // or permissions failure keeps its number instead of arriving as an
            // opaque library error.
            XlsxError::IoError(io) => FileFailure::from_io(context.clone(), io),
            other => FileFailure {
                message: format!("{}: {}", context, other),
                errno: None,
            },
        })?;

    set_output_permissions(&tmp, dest);

    tmp.persist(dest)
        .map_err(|e| FileFailure::from_io(context, &e.error))?;

    Ok(())
}

pub(crate) fn apply_defined_names(
    workbook: &mut Workbook,
    defined_names: Option<&HashMap<String, String>>,
) -> Result<(), String> {
    if let Some(names) = defined_names {
        for (name, reference) in names {
            // The local part (after a sheet-qualifying '!') must be non-empty:
            // rust_xlsxwriter's define_name calls `chars().next().unwrap()` and
            // would panic on an empty name (e.g. "" or "Sheet1!").
            let local = name.rsplit('!').next().unwrap_or("");
            if local.is_empty() {
                return Err(format!(
                    "Invalid defined name '{}': name must not be empty",
                    name
                ));
            }
            workbook
                .define_name(name, reference)
                .map_err(|e| format!("Failed to define name '{}': {}", name, e))?;
        }
    }
    Ok(())
}
