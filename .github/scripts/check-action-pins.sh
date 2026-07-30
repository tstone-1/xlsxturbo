#!/usr/bin/env bash
# Verify every SHA-pinned GitHub Action's trailing version comment actually
# matches the SHA it labels.
#
# Why this exists: Dependabot bumps the SHA but rewrites the trailing comment
# only when that comment matches the tag being replaced. A `# v4` shorthand was
# silently left in place on a bump to v5.0.0, leaving a correct SHA with a wrong
# human label -- and because the workflow still ran fine, nothing failed. The
# comment is the only human-readable record of what is pinned, so a wrong one is
# worse than none.
#
# Resolution is done by asking for the commit at the labelled version and
# comparing SHAs, rather than by listing tags at the SHA. One API call per pin,
# and it dereferences annotated tags correctly -- `git/refs/tags` returns the tag
# object's SHA for those, not the commit's, which makes a correct pin look wrong.
#
# Requires `gh` (preinstalled on GitHub runners) authenticated via GH_TOKEN.
set -uo pipefail

cd "$(dirname "$0")/../.." || exit 2

checked=0
failed=0
skipped=0

while IFS= read -r line; do
    spec="${line#*uses: }"
    spec="${spec%% *}"
    repo_with_path="${spec%@*}"
    sha="${spec#*@}"
    label="${line#*# }"

    # Subdirectory actions (github/codeql-action/init) live in the two-segment
    # repository above them.
    repo="$(printf '%s' "$repo_with_path" | cut -d/ -f1,2)"

    # A pin may deliberately track a branch, which has no version to check.
    case "$label" in
        *"no tag"* | *branch*)
            printf '[SKIP] %-46s %s (%s)\n' "$repo_with_path" "${sha:0:10}" "$label"
            skipped=$((skipped + 1))
            continue
            ;;
    esac

    checked=$((checked + 1))
    actual="$(gh api "repos/${repo}/commits/${label}" --jq '.sha' 2>/dev/null)"

    if [ -z "$actual" ]; then
        printf '[FAIL] %-46s comment says %s, which does not resolve in %s\n' \
            "$repo_with_path" "$label" "$repo"
        failed=$((failed + 1))
    elif [ "$actual" != "$sha" ]; then
        printf '[FAIL] %-46s pinned %s but %s is %s\n' \
            "$repo_with_path" "${sha:0:10}" "$label" "${actual:0:10}"
        failed=$((failed + 1))
    else
        printf '[OK]   %-46s %s == %s\n' "$repo_with_path" "${sha:0:10}" "$label"
    fi
done < <(grep -hoE "uses: [^@[:space:]]+@[0-9a-f]{40} # .*" .github/workflows/*.yml | sort -u)

echo "---"
echo "checked=$checked failed=$failed skipped=$skipped"

# A run that checked nothing is a broken run, not a clean one: the grep pattern
# could stop matching after a formatting change and this would report success
# over an empty set.
if [ "$checked" -lt 1 ]; then
    echo "no SHA-pinned actions were found -- the extraction pattern is broken" >&2
    exit 1
fi

[ "$failed" -eq 0 ] || exit 1
