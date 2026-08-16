"""Generate THIRD-PARTY-LICENSES.md from the Rust dependency tree.

The published wheel contains compiled code from 92 Rust crates. Their licenses
-- MIT, Apache-2.0, Zlib, Unicode-3.0 -- all require the copyright notice to be
distributed with the binary, and `LICENSE` covers only xlsxturbo's own code.
maturin already writes a CycloneDX SBOM into the wheel, but an SBOM records
*which* license applies; it does not carry the notice text.

This wraps `cargo about generate` (config in `about.toml`, template in
`scripts/third-party-licenses.hbs`) and removes xlsxturbo's own section, which
cargo-about has no option to exclude and which does not belong in a file about
third parties.

Run with no arguments to print the file, `--write` to update it in place, or
`--check` to verify the committed copy is current (which is what CI would do).
`--check` needs cargo-about installed; `tests/test_third_party_licenses.py`
compares the committed file against `cargo metadata` instead, which needs only a
Rust toolchain.

Install the tool with `cargo install cargo-about --features cli` -- without
`--features cli` the install prints a warning, installs no binary, and exits 0.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DOC_PATH = REPO_ROOT / "THIRD-PARTY-LICENSES.md"
TEMPLATE = REPO_ROOT / "scripts" / "third-party-licenses.hbs"
CONFIG = REPO_ROOT / "about.toml"

#: Our own crate, which cargo-about lists alongside the dependencies.
SELF_CRATE = "xlsxturbo"

#: Heading that starts each per-license section in the rendered template.
SECTION_MARK = "\n### "


def _cargo_about() -> str:
    """Locate the cargo-about binary.

    Returns:
        Path to the executable.

    Raises:
        RuntimeError: If it is not installed.
    """
    for candidate in ("cargo-about", str(Path.home() / ".cargo" / "bin" / "cargo-about")):
        found = shutil.which(candidate) or (candidate if Path(candidate).is_file() else None)
        if found:
            return found
    raise RuntimeError("cargo-about not found; install it with `cargo install cargo-about --features cli`")


def _render() -> str:
    """Run cargo-about and return its output.

    Returns:
        The rendered Markdown, before the self-section is dropped.

    Raises:
        RuntimeError: If cargo-about fails or produces something implausible.
    """
    result = subprocess.run(  # noqa: S603 - fixed argv, no shell, developer tool
        [_cargo_about(), "generate", str(TEMPLATE), "--config", str(CONFIG)],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"cargo-about failed ({result.returncode}):\n{result.stderr}")
    rendered = result.stdout
    # An empty or near-empty render is the failure mode that reads like success:
    # a notice file listing nothing looks tidy and satisfies nobody's license.
    if rendered.count(SECTION_MARK) < 4:
        raise RuntimeError(f"cargo-about produced only {rendered.count(SECTION_MARK)} license sections; refusing")
    return rendered


def _drop_self_section(rendered: str) -> str:
    """Remove the section covering xlsxturbo's own license.

    cargo-about includes the root crate and offers no flag to exclude it. The
    drop is structural rather than a regex over the text, and it asserts both
    what it removed and what remains -- a filter that silently matched nothing,
    or matched a section shared with real dependencies, would leave a file that
    still looks right.

    Args:
        rendered: cargo-about's output.

    Returns:
        The same text without the xlsxturbo section.

    Raises:
        RuntimeError: If the section is missing, duplicated, or not ours alone.
    """
    head, *sections = rendered.split(SECTION_MARK)
    ours = [i for i, section in enumerate(sections) if f"- [{SELF_CRATE} " in section]
    if len(ours) != 1:
        raise RuntimeError(f"expected exactly one section listing {SELF_CRATE}, found {len(ours)}")
    section = sections[ours[0]]
    listed = [line for line in section.splitlines() if line.startswith("- [")]
    if len(listed) != 1:
        raise RuntimeError(
            f"the {SELF_CRATE} section also covers {len(listed) - 1} dependencies "
            f"({listed}); dropping it would drop their notice too"
        )
    del sections[ours[0]]
    out = SECTION_MARK.join([head, *sections])
    if f"- [{SELF_CRATE} " in out:
        raise RuntimeError(f"{SELF_CRATE} still appears after removing its section")
    return out


def build() -> str:
    """Return the finished notice file.

    Returns:
        The Markdown to write.
    """
    return _drop_self_section(_render())


def main(argv: Sequence[str] | None = None) -> int:
    """Entry point.

    Args:
        argv: Command-line arguments, defaulting to `sys.argv[1:]`.

    Returns:
        Process exit status.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true", help="update the notice file in place")
    parser.add_argument("--check", action="store_true", help="fail if the committed file is out of date")
    args = parser.parse_args(argv)

    generated = build()

    if args.check:
        if not DOC_PATH.exists():
            print(f"[FAIL] {DOC_PATH.name} does not exist; run with --write")
            return 1
        if DOC_PATH.read_text(encoding="utf-8") != generated:
            print(f"[FAIL] {DOC_PATH.name} is out of date; run with --write")
            return 1
        print(f"[OK] {DOC_PATH.name} is current")
        return 0

    if args.write:
        # newline="\n" explicitly, so the committed bytes do not depend on which
        # platform ran the generator.
        DOC_PATH.write_text(generated, encoding="utf-8", newline="\n")
        print(f"[OK] wrote {DOC_PATH.name}")
        return 0

    print(generated)
    return 0


if __name__ == "__main__":
    sys.exit(main())
