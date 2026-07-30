#!/usr/bin/env bash
# Print the CHANGELOG.md section for one version to stdout.
#
# Usage: release-notes.sh <version> [changelog-path]
#   version        e.g. 0.18.0 (with or without a leading "v")
#   changelog-path defaults to CHANGELOG.md relative to the repo root
#
# Exits non-zero with a message on stderr if the version has no section, so a
# caller can fail loudly rather than publish an empty release body.
#
# Two heading formats exist in this file and both must be handled: the current
# "## [0.18.0] - 2026-07-25" and an older unbracketed "## 0.10.0 - 2026-01-16"
# used by 0.2.0, 0.5.0 and 0.10.0. Matching is done with index() on fixed
# strings rather than a regex: a version contains dots, and building a dynamic
# regex from it turned "[0.18.0]" into a character class that matched the wrong
# heading entirely -- silently, and with plausible-looking output.
set -euo pipefail

version="${1:?usage: release-notes.sh <version> [changelog-path]}"
version="${version#v}"
changelog="${2:-CHANGELOG.md}"

if [ ! -f "$changelog" ]; then
    echo "release-notes.sh: no such file: $changelog" >&2
    exit 2
fi

notes="$(
    awk -v bracketed="## [${version}]" -v bare="## ${version} " '
        # Start at this version'"'"'s heading, in either format.
        index($0, bracketed) == 1 || index($0, bare) == 1 { found = 1; next }
        # Stop at the next version heading, in either format.
        found && index($0, "## [") == 1 { exit }
        found && $0 ~ /^## [0-9]/       { exit }
        found                           { print }
    ' "$changelog"
)"

# Strip leading and trailing blank lines.
notes="$(printf '%s\n' "$notes" | sed -e '/./,$!d' | sed -e :a -e '/^\n*$/{$d;N;};/\n$/ba')"

if [ -z "$notes" ]; then
    echo "release-notes.sh: no CHANGELOG section found for version '$version' in $changelog" >&2
    exit 1
fi

printf '%s\n' "$notes"
