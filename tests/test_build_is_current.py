"""Fails when the built extension is older than the Rust sources it came from.

Every other test in this suite measures `python/xlsxturbo/xlsxturbo.abi3.so`
while *reading* `src/`, and nothing connects the two. A stale extension
therefore produces confident, wrong answers rather than an error, and it did
so twice on 2026-09-02: a review measured 1.3.0 behaviour and reported it as
HEAD's, and a gate run reported a `TestSaveReleasesTheGil` timing failure that
was the build rather than the code. Both cost a round of diagnosis aimed at the
source, which was never the problem.

The check is one comparison: the extension's mtime against the newest mtime
among the tracked files a rebuild consumes. It is deliberately strict rather
than tolerant, because the only correct response to a red is `maturin develop
--release`, which costs a minute.

Two consequences worth knowing before treating a failure as a false alarm:

- `git checkout` and `git stash pop` rewrite mtimes, so switching branches
  reddens this until you rebuild. That is the point -- the extension really is
  the other branch's.
- CI is unaffected. Every job builds the extension immediately before running
  pytest, so the `.so` is always the newest artifact there.

The tracked set comes from `git ls-files`, so an editor backup or a scratch
file beside the sources cannot trip it. The same choice means a *new* source
file that has not been `git add`ed yet is invisible here; it becomes visible
the moment it is staged, which is before it can reach anyone else.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

# This module audits repository sources against a build artifact. The release
# smoke-test job copies only `tests/` next to an installed wheel, where there is
# no `src/` to compare against -- see `repo_checkout_available`.
pytestmark = pytest.mark.skipif(
    not repo_checkout_available(),
    reason="build-currency test needs the Rust sources, which a wheel install does not carry",
)

# What a rebuild reads. `Cargo.lock` and `pyproject.toml` are included because a
# dependency bump or a maturin setting changes the binary without any `.rs` file
# moving.
BUILD_INPUT_NAMES = frozenset({"Cargo.toml", "Cargo.lock", "pyproject.toml"})


def _tracked_build_inputs() -> list[Path]:
    """Return the tracked files whose change requires a rebuild.

    Returns:
        Absolute paths to every tracked `src/**/*.rs` plus the three
        manifest files, in `git ls-files` order.

    Raises:
        pytest.skip.Exception: When git cannot enumerate them, which means
            this checkout is not a git working tree (an unpacked sdist, say)
            rather than that the build is current.
    """
    try:
        proc = subprocess.run(  # noqa: S603 - fixed argv, no shell
            ["git", "ls-files", "--", "src", *sorted(BUILD_INPUT_NAMES)],  # noqa: S607
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:  # pragma: no cover - git absent
        pytest.skip(f"git is not runnable, so the tracked set cannot be read: {exc}")
    if proc.returncode != 0:  # pragma: no cover - not a git working tree
        pytest.skip(f"git ls-files failed ({proc.returncode}), so this is not a git checkout")

    names = [line for line in proc.stdout.splitlines() if line.strip()]
    return [
        REPO_ROOT / name
        for name in names
        if name.endswith(".rs") or name in BUILD_INPUT_NAMES
    ]


def _extension_path() -> Path:
    """Return the compiled extension the rest of the suite is testing.

    Taken from the imported module rather than guessed from a filename
    pattern, so this measures the artifact actually in use -- including the
    case where an editable install has been shadowed by a wheel elsewhere,
    which is exactly the staleness worth catching.

    Returns:
        The absolute path to the extension module's shared library.
    """
    import xlsxturbo.xlsxturbo as extension

    path = getattr(extension, "__file__", None)
    assert path is not None, "the extension module reports no __file__"
    return Path(path).resolve()


class TestBuildIsCurrent:
    """The extension in the venv matches the sources in the checkout."""

    def test_the_tracked_build_inputs_are_not_empty(self) -> None:
        """The emptiness control for the comparison below.

        A wrong pathspec, a renamed directory or a git invocation that quietly
        matched nothing would leave `max()` with no candidates and make the
        real assertion pass for every possible build. `src/lib.rs` is named
        explicitly because it is the crate root: it cannot go missing while
        the check remains meaningful.
        """
        inputs = _tracked_build_inputs()

        assert inputs, "git ls-files matched no build inputs; the pathspec is wrong"
        assert REPO_ROOT / "src" / "lib.rs" in inputs, (
            f"src/lib.rs is missing from the {len(inputs)} tracked build inputs"
        )

    def test_the_extension_is_newer_than_every_source_it_was_built_from(self) -> None:
        """The built extension postdates the newest source a rebuild reads.

        The message names the offending file, because "rebuild" is unhelpful
        when what actually changed was `Cargo.lock` rather than anything the
        reader edited.
        """
        extension = _extension_path()
        assert extension.is_file(), f"no extension module at {extension}"
        built_at = extension.stat().st_mtime

        newer = [
            path
            for path in _tracked_build_inputs()
            if path.is_file() and path.stat().st_mtime > built_at
        ]

        assert not newer, (
            f"{extension.name} was built before "
            f"{', '.join(sorted(p.relative_to(REPO_ROOT).as_posix() for p in newer))} "
            f"was last modified, so every test in this run is measuring stale code. "
            f"Rebuild with: maturin develop --release"
        )
