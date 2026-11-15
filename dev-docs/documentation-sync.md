# Documentation sync

Moved verbatim from `AGENTS.md` on 2026-11-15, which keeps one index line per section.

## Documentation Sync

Migrated here from the personal cross-repo memory file on 2026-07-30: it was xlsxturbo-specific knowledge living outside the repo, and its first step had gone stale (it said usage examples belong in the README, which stopped being true with the 0.19.0 split).

When adding or modifying a feature:

1. **The relevant `docs/` page** — add or update the usage example. `docs/` is the MkDocs Material site published to GitHub Pages; the README is a landing page and should NOT grow new per-feature examples. Match the option to its page from the nav in `mkdocs.yml` (formatting, tables, charts-and-media, cells, ...). A brand-new page must be added to that nav or `tests/test_docs_site.py` fails it as unreachable.
2. **`docs/capability-matrix.md`** — regenerate with `python scripts/gen_capability_matrix.py --write`. Never hand-edit it. `tests/test_capability_matrix.py` fails when the committed page is stale.
3. **`CHANGELOG.md`** — document all changes under the appropriate version heading. Note the file is NOT uniform: 37 headings are `## [X.Y.Z]` and three older ones are `## X.Y.Z` without brackets. `.github/scripts/release-notes.sh` handles both by fixed-string matching; do not "simplify" it to a regex, which is how a release once got the wrong version's notes.

   *Unreleased work goes under `## [Unreleased]`, never `## [X.Y.Z] - Unreleased`* -- `release-notes.sh` matches the version heading as a prefix, so the second form publishes a Release whose notes say Unreleased with exit 0; `TestChangelogHeadings` refuses it. This repo's rule wins over the cross-repo note. Full text: [dev-docs/changelog.md](dev-docs/changelog.md#unreleased-work-goes-under-unreleased).
4. **Type stubs** (`python/xlsxturbo/xlsxturbo.pyi`) — add new parameters with types and docstrings. This compiled-extension stub is the type source of truth; `python/xlsxturbo/__init__.pyi` is a thin re-export and must not be hand-edited for new options.

Before commit or push, follow the `BUILD.md` checklist.
