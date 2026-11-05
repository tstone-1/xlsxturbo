# The stability page is checked

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

### The page is checked, not maintained by hand

It is **checked, not maintained by hand**. `tests/test_stability_policy.py` compares the
page against the four places that actually decide its contents — the trove classifiers and
`requires-python` in `pyproject.toml`, the interpreter matrix in `ci.yml`, the wheel targets
in `release.yml`, and `xlsxturbo.__all__` — in both directions. Consequences:

- **A new exported name fails the suite until the page names it.** That is deliberate: the
  page is the list of things that cannot be removed before 2.0.0, so adding to it should be a
  decision rather than a side effect. New *options* need nothing here; the touchpoint
  checklist already covers those.
- **Adding a Python version to the CI matrix fails until the classifier and the page agree.**
  A version tested and nowhere else declared reads as a widened promise when it is only a
  widened test.
- **This worked.** Dropping Python 3.9 in 1.1.0 failed both suites exactly as designed — the
  page's version table against the classifiers and the CI matrix, and the Dependabot hold in
  `tests/test_ci_config.py` that existed only because pytest 9 needs 3.10. Neither failure was
  a surprise to be worked around; each named the work that had to accompany the drop.
