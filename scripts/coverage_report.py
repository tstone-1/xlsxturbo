"""Measure test coverage of the Rust core and the Python layer.

Run from the repository root::

    python scripts/coverage_report.py            # table to stdout
    python scripts/coverage_report.py --html     # plus browsable HTML
    python scripts/coverage_report.py --lcov out.info

**There is no threshold here and there must not be one.** A number that gates a
merge gets optimised, and what it gets optimised with is tests that execute
lines without asserting anything -- which is worse than no coverage data,
because it looks like progress. What this report is for is the list of *which*
branches are unexercised, error paths in particular; the percentage is a
by-product.

Why it is built the way it is
-----------------------------

Almost every line of this library runs from Python, not from ``cargo test``.
Measuring only the Rust test binaries reports **26%** and shows every
``src/apply/*.rs`` file at zero, which reads as an untested library and is
simply false -- those paths are covered thoroughly, from the other side of the
FFI boundary. Measuring only the Python suite reports **91%** and in turn misses
the parser branches only the Rust unit and property tests reach.

So both halves are measured and their profiles merged: instrumented test
binaries *and* an instrumented extension module, one ``llvm-profdata`` merge,
one ``llvm-cov`` report over every object. Neither half alone is honest.

This deliberately does not use ``cargo-llvm-cov``. It is the better tool for
the Rust half, but its ``report`` subcommand cannot be pointed at an extra
object file, which is exactly what the extension module is. Driving
``llvm-profdata`` and ``llvm-cov`` directly needs only ``llvm-tools-preview``,
which ``rustup`` ships.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PROFILE_DIR = REPO_ROOT / "target" / "coverage"
EXTENSION_DIR = REPO_ROOT / "python" / "xlsxturbo"

# Set for *every* pytest pass this script drives, so the suite can tell it is running
# against an instrumented build. `LLVM_PROFILE_FILE` is not a substitute: it says where
# profiles go, and only one of the two passes below sets it -- which is how a timing
# assertion went on running under instrumentation after being "fixed".
COVERAGE_MARKER = {"XLSXTURBO_COVERAGE": "1"}

# What the report is *not* about, as one regex, because `llvm-cov` filters by
# filename and nothing finer:
#
# - dependencies and the toolchain, or the report is mostly other people's code;
# - `tests/`, `src/parse/proptests.rs` and `src/parse/boundaries.rs`, which are
#   test code. Including them adds several hundred always-executed regions and
#   pushes the total up by about a point, which is a percentage measuring the
#   tests' coverage of themselves.
#
# One thing this cannot exclude: the `#[cfg(test)] mod tests` living inside
# `src/parse/mod.rs`. `llvm-cov` has no sub-file filter, so that module's ~600
# regions still count. The row for `parse/mod.rs` should be read as "the tests
# in this file all ran", not as a statement about the parsers.
IGNORE_PATTERN = (
    r"(/\.cargo/|[\\/]rustc[\\/]|[\\/]registry[\\/]|/rustlib/"
    r"|[\\/]tests[\\/]|[\\/]parse[\\/](proptests|boundaries)\.rs)"
)


def run(
    command: list[str],
    *,
    env: dict[str, str] | None = None,
    capture: bool = False,
) -> subprocess.CompletedProcess[str]:
    """Run a command from the repository root, failing loudly.

    Args:
        command: Argument vector to execute.
        env: Extra environment variables layered over the current environment.
        capture: Whether to capture stdout instead of letting it through.

    Returns:
        The completed process.

    Raises:
        SystemExit: If the command fails. A coverage run that half-happened
            produces a plausible-looking report over incomplete data, which is
            worse than no report.
    """
    merged = {**os.environ, **(env or {})}
    # S603: every command here is built from literals in this file plus paths
    # this script derived itself (the toolchain's LLVM tools, cargo's own
    # reported artifact paths). Nothing reaches it from user input, and no
    # shell is involved.
    result = subprocess.run(  # noqa: S603
        command,
        cwd=REPO_ROOT,
        env=merged,
        text=True,
        capture_output=capture,
        check=False,
    )
    if result.returncode != 0:
        if capture:
            sys.stderr.write(result.stdout or "")
            sys.stderr.write(result.stderr or "")
        sys.exit(f"[FAIL] command failed ({result.returncode}): {' '.join(command)}")
    return result


def llvm_tool(name: str) -> str:
    """Locate an LLVM tool from the active Rust toolchain.

    Args:
        name: Tool name without the platform suffix, e.g. ``llvm-cov``.

    Returns:
        Absolute path to the executable.

    Raises:
        SystemExit: If the tool is missing, with the command that installs it.
    """
    libdir = run(
        ["rustc", "--print", "target-libdir"], capture=True
    ).stdout.strip()
    suffix = ".exe" if sys.platform == "win32" else ""
    tool = Path(libdir).parent / "bin" / f"{name}{suffix}"
    if not tool.exists():
        sys.exit(
            f"[FAIL] {name} not found at {tool}\n"
            "       Install it with: rustup component add llvm-tools-preview"
        )
    return str(tool)


def venv_tool(name: str) -> str:
    """Resolve a tool from the running interpreter's own environment.

    Falls back to ``PATH``. Without this the script works only when the venv
    has been *activated*, and running it as ``.venv/bin/python scripts/...``
    would silently build with whatever ``maturin`` came first on ``PATH`` --
    possibly installing into a different environment than the one about to run
    pytest.

    Args:
        name: Executable name without a platform suffix.

    Returns:
        The tool's path, or ``name`` itself if it is not beside the interpreter.
    """
    suffix = ".exe" if sys.platform == "win32" else ""
    candidate = Path(sys.executable).parent / f"{name}{suffix}"
    return str(candidate) if candidate.exists() else name


def find_extension() -> Path:
    """Locate the compiled extension module ``maturin develop`` produced.

    Returns:
        Path to the ``.pyd`` / ``.so`` / ``.dylib``.

    Raises:
        SystemExit: If no extension is present, which means the build did not
            put one where this expects and every later number would be wrong.
    """
    import sysconfig

    search = [EXTENSION_DIR, Path(sysconfig.get_paths()["purelib"]) / "xlsxturbo"]
    for directory in search:
        for pattern in ("xlsxturbo*.pyd", "xlsxturbo*.so", "xlsxturbo*.dylib"):
            found = sorted(directory.glob(pattern))
            if found:
                return found[0]
    sys.exit(f"[FAIL] no compiled extension found in any of: {search}")


def build_instrumented_binaries(env: dict[str, str]) -> tuple[list[Path], list[Path]]:
    """Compile the Rust test binaries with coverage instrumentation.

    Args:
        env: Environment carrying the instrumentation flags.

    Returns:
        ``(test_binaries, other_binaries)``. The first are executed directly.
        The second is the CLI binary, which ``tests/cli.rs`` spawns as a
        subprocess -- it is instrumented and writes its own profile, but it is
        not in the test list, so without collecting it here as an *object*
        ``src/main.rs`` reports 0% while being fully exercised.

    Raises:
        SystemExit: If cargo reports no test executables at all, which reads
            exactly like a suite with nothing in it.
    """
    result = run(
        ["cargo", "test", "--no-run", "--message-format=json"],
        env=env,
        capture=True,
    )
    tests: list[Path] = []
    others: list[Path] = []
    for line in result.stdout.splitlines():
        try:
            message = json.loads(line)
        except json.JSONDecodeError:
            continue
        if message.get("reason") != "compiler-artifact" or not message.get("executable"):
            continue
        target = tests if message.get("profile", {}).get("test") else others
        target.append(Path(message["executable"]))
    if not tests:
        sys.exit("[FAIL] cargo produced no test executables")
    return tests, others


def report_python_layer() -> None:
    """Report coverage of the pure-Python layer, separately.

    ``python/xlsxturbo`` is the thin layer above the extension: the option
    ``TypedDict``s, :class:`~xlsxturbo.options.ExportOptions`, and the package
    exports. ``llvm-cov`` cannot see it -- it is not compiled -- so it needs
    its own pass, and it is reported separately rather than blended into one
    figure that would mean nothing.

    Raises:
        SystemExit: If ``coverage`` is not installed, naming the fix. A silent
            skip here would leave the whole Python layer unmeasured while the
            report still looked complete.
    """
    try:
        import coverage  # noqa: F401  (probing for availability)
    except ImportError:
        sys.exit(
            '[FAIL] coverage is not installed. Install the dev extras: uv pip install -e ".[dev]"'
        )
    run(
        [
            sys.executable,
            "-m",
            "coverage",
            "run",
            "--source=python/xlsxturbo",
            "-m",
            "pytest",
            "tests/",
            "-q",
        ],
        # This pass runs against whatever the venv holds, which by now is the
        # instrumented extension built in step 3 -- and it adds coverage.py's own
        # tracing on top. Wall-clock measurements are meaningless under both, so
        # the suite is told, and `tests/test_concurrency.py` skips its timing
        # assertions. Without this the pass took 24.2s against the same suite's
        # 12.5s and failed those assertions intermittently.
        env=COVERAGE_MARKER,
    )
    run([sys.executable, "-m", "coverage", "report", "--show-missing"])


def main() -> int:
    """Build both halves instrumented, run both suites, merge, report."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--html", action="store_true", help="also write an HTML report")
    parser.add_argument("--lcov", metavar="PATH", help="also write an lcov trace file")
    parser.add_argument(
        "--summary-only",
        action="store_true",
        help="print only the total line, for a CI summary",
    )
    args = parser.parse_args()

    profdata_tool = llvm_tool("llvm-profdata")
    cov_tool = llvm_tool("llvm-cov")

    # A stale profile silently inflates the result, so start from nothing.
    if PROFILE_DIR.exists():
        shutil.rmtree(PROFILE_DIR)
    PROFILE_DIR.mkdir(parents=True)

    instrument = {"RUSTFLAGS": "-Cinstrument-coverage"}
    profile_env = {
        **instrument,
        **COVERAGE_MARKER,
        "LLVM_PROFILE_FILE": str(PROFILE_DIR / "xlsxturbo-%p-%m.profraw"),
    }

    print("[1/5] building instrumented Rust binaries", flush=True)
    test_binaries, other_binaries = build_instrumented_binaries(instrument)

    print(f"[2/5] running {len(test_binaries)} Rust test binaries", flush=True)
    for binary in test_binaries:
        run([str(binary)], env=profile_env)

    print(
        "[3/5] building the instrumented extension and running the Python suite",
        flush=True,
    )
    run([venv_tool("maturin"), "develop", "--release"], env=instrument)
    extension = find_extension()
    run([sys.executable, "-m", "pytest", "tests/", "-q"], env=profile_env)

    raw = sorted(PROFILE_DIR.glob("*.profraw"))
    if not raw:
        # No profile at all reads as "nothing ran", and a report over an empty
        # profile prints a tidy 0% rather than an error.
        sys.exit("[FAIL] no .profraw files were produced; nothing was measured")
    print(f"[4/5] merging {len(raw)} profiles and reporting", flush=True)

    profdata = PROFILE_DIR / "xlsxturbo.profdata"
    run([profdata_tool, "merge", "-sparse", *[str(p) for p in raw], "-o", str(profdata)])

    objects: list[str] = [str(extension)]
    for binary in [*test_binaries, *other_binaries]:
        objects += ["-object", str(binary)]
    common = [
        f"--instr-profile={profdata}",
        *objects,
        f"--ignore-filename-regex={IGNORE_PATTERN}",
    ]

    report = [cov_tool, "report", *common]
    if args.summary_only:
        report.append("--summary-only")
    print("\n=== Rust core (src/), from both suites ===\n", flush=True)
    run(report)

    print("\n[5/5] measuring the pure-Python layer", flush=True)
    print("\n=== Python layer (python/xlsxturbo/) ===\n", flush=True)
    report_python_layer()

    if args.lcov:
        with Path(args.lcov).open("w", encoding="utf-8") as handle:
            handle.write(
                run(
                    [cov_tool, "export", "--format=lcov", *common], capture=True
                ).stdout
            )
        print(f"\n[OK] lcov written to {args.lcov}")

    if args.html:
        out = REPO_ROOT / "target" / "coverage-html"
        run([cov_tool, "show", "--format=html", f"-output-dir={out}", *common])
        print(f"\n[OK] HTML report at {out / 'index.html'}")

    print(
        "\nNo threshold is applied. Read the per-file rows for unexercised "
        "error paths; the total is context, not a target."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
