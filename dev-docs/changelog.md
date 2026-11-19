# CHANGELOG headings

Moved verbatim from `AGENTS.md`, step 3 of Documentation Sync, which keeps a one-line
entry pointing here.

## Unreleased work goes under `[Unreleased]`

**Unreleased work goes under `## [Unreleased]`, never `## [X.Y.Z] - Unreleased`, and the difference decides whether a forgotten rename fails loudly or ships wrong.** `release-notes.sh` matches `## [<version>]` as a prefix and ignores whatever follows on the line, so a heading already carrying the version number matches its tag: measured against this repo's own file, `release-notes.sh v1.1.2` on `## [1.1.2] - Unreleased` exits **0** and prints the section, and the release job then publishes a GitHub Release whose notes came from a heading that says Unreleased. The same call on `## [Unreleased]` exits **1** with `no CHANGELOG section found for version '1.1.2'`, which fails the job before anything is published. Step 2 of `BUILD.md`'s release process is where the heading gets its version and date. `tests/test_ci_config.py::TestChangelogHeadings` fails on a bracketed-version heading whose date field is missing or reads "Unreleased", so the unsafe form cannot come back by habit.

The `## [version] - Unreleased` form is a common convention elsewhere, and this rule contradicts it on purpose. That form is only safe where the release tooling does not slice notes by matching the version heading; here it is not. Which form a project can use is decided by its own release script, so a convention carried over from another project does not apply here.
