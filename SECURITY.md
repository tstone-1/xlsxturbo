# Security Policy

## Supported versions

xlsxturbo is pre-1.0. Security fixes go into the next release from `main`; there are no
maintained backport branches.

| Version | Supported |
|---------|-----------|
| 0.18.x  | Yes       |
| < 0.18  | No -- upgrade |

Once 1.0 ships, this table will state a real support window. Until then, "supported"
means the latest release.

## Reporting a vulnerability

Please report privately, not as a public issue.

Use GitHub's [private vulnerability reporting](https://github.com/tstone-1/xlsxturbo/security/advisories/new)
on this repository. That opens a draft advisory visible only to the maintainer.

Useful things to include:

- what the vulnerability lets an attacker do
- the version, platform and Python version
- a minimal reproducer -- an input file plus the call that processes it
- whether the input needs to be attacker-controlled, and which parameter carries it

This is a single-maintainer project, so expect an acknowledgement within about a week
rather than within hours. If a fix is warranted it will ship in the next release, with
the advisory published at the same time. Credit is given unless you'd rather not have it.

## Scope

xlsxturbo writes `.xlsx` files. It does not read or parse them, so the classes of
vulnerability that affect spreadsheet *readers* -- formula evaluation, external entity
resolution, macro execution -- mostly do not apply.

What is in scope:

- memory unsafety reachable from Python-level input (a panic is a bug; a segfault or
  memory corruption is a security bug)
- path traversal or unintended file access through an output path, an image path, or
  any other filesystem-touching parameter
- generating a workbook that reliably exploits a consumer -- for example an injected
  formula or hyperlink that a spreadsheet application acts on without user consent
- a supply-chain problem in the release pipeline or in the published wheels

What is not in scope:

- untrusted data landing in cells as data. Values are written as values; if your
  threat model includes a downstream application that evaluates cell contents, treat
  formula-shaped strings the way you would with any writer.
- denial of service through deliberately enormous input. Constant-memory mode exists
  for size; there is no input-size limit and none is planned.
- vulnerabilities in `pandas`, `polars` or `openpyxl` -- report those upstream.

## Dependencies

`cargo audit` and `pip-audit` run on every push and pull request, so a new advisory
against a dependency surfaces on the next CI run rather than at release time. The
shipped wheel declares no runtime Python dependencies; the Rust crates linked into the
extension are recorded in `Cargo.lock` and published as a CycloneDX SBOM with each
release. Wheels carry build provenance attestations from the release workflow.
