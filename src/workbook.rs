//! Shared workbook-level helpers.

use crate::errors::{errno, FileFailure};
use crate::parse::looks_like_cell_reference;
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

/// The part of a defined name Excel validates: everything after a `!` sheet
/// qualifier, or the whole string when there is none.
///
/// `Workbook::define_name` splits on `!` and applies its own rules (non-empty,
/// legal first character, no forbidden characters) to the local part only, so a
/// screen added here has to look at the same half or `"Sheet1!Q1"` walks past
/// it. Reported as an unqualified name by the crate, which is why the caller's
/// key goes back in front of the message.
fn local_part(name: &str) -> &str {
    match name.rfind('!') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

/// Reject a defined name Excel would read as a cell reference.
///
/// Unlike a table name, which is sanitized, this refuses: silently renaming a
/// defined name would leave every formula that references it pointing at a name
/// the workbook no longer defines, so the caller has to choose the new name.
/// rust_xlsxwriter 0.98.2 validates the first character and the character set
/// but not this, so nothing else catches it.
fn reject_reference_shaped_name(name: &str) -> Result<(), String> {
    // rust_xlsxwriter 0.98.2 does not check this rule; reported as
    // https://github.com/jmcnamara/rust_xlsxwriter/issues/189. When the crate
    // validates it, this screen can be revisited the way the empty-name screen
    // was for #186 — kept only if its message names the offending key better.
    if looks_like_cell_reference(local_part(name)) {
        return Err(format!(
            "defined_names['{}']: '{}' is a cell reference and cannot be used as a defined name. \
             Excel reserves names that address a cell (A1-style, the R1C1 forms, and the \
             selection shortcuts 'R' and 'C'); pick another name, e.g. '_{}'",
            name,
            local_part(name),
            local_part(name)
        ));
    }
    Ok(())
}

pub(crate) fn apply_defined_names(
    workbook: &mut Workbook,
    defined_names: Option<&HashMap<String, String>>,
) -> Result<(), String> {
    if let Some(names) = defined_names {
        for (name, reference) in names {
            reject_reference_shaped_name(name)?;
            workbook
                .define_name(name, reference)
                .map_err(|e| format!("Failed to define name '{}': {}", name, e))?;
        }
    }
    Ok(())
}
