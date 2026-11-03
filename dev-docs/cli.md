# The two CLIs

Moved verbatim from `AGENTS.md`, "Two CLIs, one interface", which keeps a one-line entry pointing here.

## Two CLIs, one interface

The wheel's `xlsxturbo` command is `python/xlsxturbo/cli.py`, registered under `[project.scripts]`. It is **not** the Rust binary: maturin builds only the extension module, so the `[[bin]] xlsxturbo` target (`src/main.rs`, `default = ["cli"]`) exists only in `cargo build --release` output and has never been in a wheel (checked against the PyPI artifact for 0.18.0; the Python command was added after that and checked in a locally built wheel's `entry_points.txt`). Both must accept the same flags and use the same exit codes (0 / 1 failed conversion / 2 usage), so a new option goes into **both**, and `tests/test_cli.py::TestParityWithRustBinary` fails when their option lists differ. `docs/stability.md` covers the Python command's options, positionals, `OK <rows> <cols>` line and exit codes; the Rust binary is not covered.
