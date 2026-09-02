# Compatibility and limitations

xlsxturbo writes `.xlsx`; it never reads or modifies one. What follows is the set of
places where Excel's data model and Python's do not line up, plus what happens to the
file on disk when a write fails.

## Known Limitations

- **Datetime display precision**: Sub-second precision is preserved in the stored Excel datetime serial, but the default display format shows whole seconds.
- **Timezone-aware datetimes**: Written as their local wall-clock value; the UTC offset is **not** preserved (Excel has no timezone concept). A `2024-01-01 12:00 US/Eastern` value is stored as `12:00`, not converted to UTC. Normalize to UTC beforehand (e.g. `df["ts"].dt.tz_convert("UTC").dt.tz_localize(None)`) if you need UTC.
- **Large integers**: Integers exceeding 2^53 (9,007,199,254,740,992) are written as strings to prevent silent precision loss in Excel's floating-point representation.
- **Validation lists**: Limited to 255 total characters (Excel limitation).
- **Append mode**: Existing workbook mutation is not supported because the Rust writer is write-only. Create a new workbook instead.

## Output File Safety

Writes are atomic: the workbook is built in a temporary file alongside the destination and
renamed over it only once it is complete. If a write fails for any reason — an invalid
chart range, a full disk, a dropped network share — the file already at the output path is
left exactly as it was, rather than being truncated. Re-exporting over yesterday's report
can therefore never leave you with neither.

Because the staging file is created in the destination's directory, that directory must
exist and be writable. When an existing file is replaced, its permissions are preserved.

A symlink at the output path is **written through**: the file it points at is replaced and
the link itself survives, so exporting to `latest.xlsx -> archive/2026-09.xlsx` updates the
archive. The staging file is created beside the resolved target, which is what keeps the
rename atomic. A link whose target does not exist is the one exception — the export creates
a regular file at the link's own path instead of failing. Following the link would write
to a path the caller never named, in a directory that need not exist; creating the file
where the caller pointed is the smaller surprise.

This restores the behaviour of every release before 0.18.0, which wrote through a symlink
because the underlying writer used `File::create`. Between 0.18.0 and 1.3.0 the atomic
save's rename replaced the *link*, silently orphaning its target.
