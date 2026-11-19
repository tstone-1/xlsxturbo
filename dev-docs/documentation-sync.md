# Documentation sync

Moved verbatim from `AGENTS.md` on 2026-11-15, which keeps one index line per section.

## Documentation Sync

Kept in the repository since 2026-07-30. Before that the list lived outside it, and its first step had gone stale (it said usage examples belong in the README, which stopped being true with the 0.19.0 split).

When adding or modifying a feature:

1. **The relevant `docs/` page** — add or update the usage example. `docs/` is the MkDocs Material site published to GitHub Pages; the README is a landing page and should NOT grow new per-feature examples. Match the option to its page from the nav in `mkdocs.yml` (formatting, tables, charts-and-media, cells, ...). A brand-new page must be added to that nav or `tests/test_docs_site.py` fails it as unreachable.

   Every `python` block on a page is executed by `tests/test_docs_examples.py`, a page's blocks in order in one namespace (since 2026-11-19; before that only `docs/migrating.md` was run, and `docs/errors.md` showed a handler that did not catch the error its own example raised). So an example has to run as written. What the test gives it: `xlsxturbo`, `pd`, `pl`, a small `df`, `path`, the frames named in `seeded_namespace()`, and the input files in `CSV_INPUTS` and `IMAGE_INPUTS`. An example that is meant to raise says so on the page, as a trailing `# SomeError: message` comment or as a `>>>` line followed by `xlsxturbo.SomeError: message`; the test then requires exactly that class and the start of that message. A block that cannot run at all (it needs a web framework, or is the body of a function) goes into `NOT_RUN` with its reason. Checked on Python 3.12 with pandas 2 and polars 1 as well as the current versions, because CI runs both.
2. **`docs/capability-matrix.md`** — regenerate with `python scripts/gen_capability_matrix.py --write`. Never hand-edit it. `tests/test_capability_matrix.py` fails when the committed page is stale.
3. **`CHANGELOG.md`** — document all changes under the appropriate version heading. Note the file is NOT uniform: 37 headings are `## [X.Y.Z]` and three older ones are `## X.Y.Z` without brackets. `.github/scripts/release-notes.sh` handles both by fixed-string matching; do not "simplify" it to a regex, which is how a release once got the wrong version's notes.

   *Unreleased work goes under `## [Unreleased]`, never `## [X.Y.Z] - Unreleased`* -- `release-notes.sh` matches the version heading as a prefix, so the second form publishes a Release whose notes say Unreleased with exit 0; `TestChangelogHeadings` refuses it. This repo's rule wins over the cross-repo note. Full text: [dev-docs/changelog.md](dev-docs/changelog.md#unreleased-work-goes-under-unreleased).
4. **Type stubs** (`python/xlsxturbo/xlsxturbo.pyi`) — add new parameters with types and docstrings. This compiled-extension stub is the type source of truth; `python/xlsxturbo/__init__.pyi` is a thin re-export and must not be hand-edited for new options.

Before commit or push, follow the `BUILD.md` checklist.
