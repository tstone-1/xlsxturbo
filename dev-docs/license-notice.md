# The bundled license notice

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

## The bundled license notice

`THIRD-PARTY-LICENSES.md` is generated — `python scripts/gen_third_party_licenses.py --write`
(cargo-about, config `about.toml`, template `scripts/third-party-licenses.hbs`). Never edit
it by hand; `tests/test_third_party_licenses.py` compares it against `cargo metadata` in
both directions and will say so.

Why it exists at all: the wheel is a binary containing compiled code from every crate in
the dependency tree (the notice itself is the count; a hand-written number here disagreed
with three others within two releases), and
MIT, Apache-2.0, Zlib and Unicode-3.0 all require the copyright notice to be distributed
with a binary. `LICENSE` covers xlsxturbo's own code and nothing else. **maturin's
CycloneDX SBOM is not a substitute** — it records which license applies to each crate,
which is not the notice the license asks for.

Five things that cost time to find:

- **cargo-about 0.9 on Windows refuses to write to a redirected stdout when PowerShell is a
  parent process** (`ERROR ... please use the -o, --output-file option`), and an agent shell
  always has one. The generator therefore passes `--output-file`; do not switch it back to
  capturing stdout. Found during the 1.5.1 release, when Dependabot's rust_xlsxwriter bump
  (#42) needed the notice regenerated and the script could not run on the desktop.
- **`cargo install cargo-about` installs nothing and exits 0.** Its binary is behind a
  feature: `cargo install cargo-about --features cli`. Without it you get a warning, a
  clean exit, and no `cargo-about` on PATH.
- **PEP 639 `license-files` is what puts the notice in the wheel**, at
  `.dist-info/licenses/`. maturin gained it in **1.9.0**, so `[build-system] requires` says
  `maturin>=1.9`. An older backend ignores the key and builds a wheel with neither notice
  and no error. The two are asserted together in the test for that reason. `[tool.maturin]
  include` also works but drops the file at the *site-packages root*, which is worse.
- **cargo-about lists the root crate and has no flag to exclude it**, so the generator drops
  that section structurally and refuses if it is missing, duplicated, or shared with a real
  dependency.
- **`serde_core → serde_derive` is declared under `target = "cfg(any())"`** — false for
  every target, the idiom for "declared but never compiled". cargo-about omits it correctly;
  the test has to skip that edge or it reports a missing notice for a crate we do not ship.
  Every other `cfg(...)` edge is kept, so a Windows-only crate is still covered in a
  macOS-built wheel's notice.
