# The documentation site

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

### docs/ is published wholesale — mind what is sitting in it

MkDocs publishes every file under `docs/`, and knows nothing about git. Two tracked-but-internal files (`roadmap-1.0.md`) and two untracked ones (`strategic-recommendations-plan.md`, `reviews/`) are therefore listed in `mkdocs.yml`'s `exclude_docs`, and `tests/test_docs_site.py` asserts that list stays in step with `.gitignore`.

Verified by removing one `exclude_docs` entry: the private planning memo appeared in `site/` and `mkdocs build --strict` still **exited 0**. Strict mode does not cover this. The reliable protection is that deployment happens only from `.github/workflows/docs.yml`, which builds from a clean checkout and so cannot see an untracked file at all — never run `mkdocs gh-deploy` by hand.
