#!/usr/bin/env bash
# Auto-generates a release manifest: what PRs/commits a release actually contains, what
# version each versioned crate/package is at, and (optionally) what was deliberately
# excluded relative to another ref (e.g. master, when the release was cut from an older
# tag specifically to leave still-unverified work out -- see ROADMAP.md's "Current
# state"). Answers ROADMAP.md's 0.13.0 "auto-generate a release manifest" packaging-
# integrity item -- previously this was hand-written prose in commit messages, PR
# descriptions, and tasks/todo.md round-summaries, reconstructed by memory each time.
#
# Purely read-only: never checks out a ref, never mutates the working tree or any
# branch -- every fact is read via `git log`/`git show` against the refs given, so it's
# safe to run against the current branch mid-work, or in CI.
#
# Usage: scripts/generate-release-manifest.sh <base-ref> <head-ref> [exclude-vs-ref]
#
#   base-ref       the release's starting point (e.g. a prior release tag, v0.12.0)
#   head-ref       the release branch/commit/tag whose contents to describe
#   exclude-vs-ref optional -- also lists commits reachable from this ref but not from
#                  head-ref (relative to base-ref), labeled "excluded from this
#                  release" (e.g. pass `master` to see what's on master but deliberately
#                  left out of a release branch cut from an older tag)
#
# Requires `gh` (GitHub CLI, authenticated) to fetch PR titles/URLs, and `jq`.
set -euo pipefail

cd "$(dirname "$0")/.."

if [[ $# -lt 2 || $# -gt 3 ]]; then
  echo "usage: $0 <base-ref> <head-ref> [exclude-vs-ref]" >&2
  exit 1
fi

base_ref=$1
head_ref=$2
exclude_ref=${3:-}

for ref in "$base_ref" "$head_ref" ${exclude_ref:+"$exclude_ref"}; do
  if ! git rev-parse --verify --quiet "${ref}^{commit}" >/dev/null; then
    echo "generate-release-manifest: '$ref' does not resolve to a commit" >&2
    exit 1
  fi
done

version_at() {
  # $1: ref, $2: path, $3: (optional) line-matching prefix -- default 'version = "'
  git show "$1:$2" 2>/dev/null | grep -m1 "${3:-version = \"}" | sed -E 's/^[^"]*"([^"]*)".*/\1/' || echo "(absent)"
}

echo "# Release manifest: $base_ref..$head_ref"
echo
echo "Generated $(date -u +%Y-%m-%dT%H:%M:%SZ 2>/dev/null || echo unknown) — read-only, no branch/ref was touched."
echo
echo "## Versions at $head_ref"
echo
echo "| Component | Version |"
echo "|---|---|"
echo "| \`elixcee\` (Cargo.toml) | $(version_at "$head_ref" Cargo.toml) |"
echo "| \`elixcee\` (pyproject.toml) | $(version_at "$head_ref" pyproject.toml) |"
echo "| \`elixcee-types\` | $(version_at "$head_ref" crates/elixcee-types/Cargo.toml) |"
echo "| \`elixcee-wasm\` | $(version_at "$head_ref" crates/elixcee-wasm/Cargo.toml) |"
echo

echo "## Included: $base_ref..$head_ref"
echo

all_commits=$(git log --format="%H %s" "$base_ref..$head_ref")

# Classified by commit MESSAGE ("Merge pull request #N from ..."), not parent count --
# a release branch built via `cherry-pick -m 1` (see build-release-branch.sh) turns a
# real merge commit into a single-parent commit that just keeps the original message,
# so parent count can't tell PRs apart from standalone commits there.
merge_commits=$(echo "$all_commits" | grep -E ' Merge pull request #[0-9]+' || true)
standalone_commits=$(echo "$all_commits" | grep -vE ' Merge pull request #[0-9]+' || true)

if [[ -n "$merge_commits" ]]; then
  echo "### Merged PRs"
  echo
  while IFS= read -r line; do
    [[ -z "$line" ]] && continue
    sha=${line%% *}
    subject=${line#* }
    pr_num=$(echo "$subject" | grep -oE '#[0-9]+' | head -1 | tr -d '#' || true)
    if [[ -n "$pr_num" ]] && gh pr view "$pr_num" --json title,url,mergedAt >/tmp/pr_manifest_$$.json 2>/dev/null; then
      title=$(jq -r '.title' /tmp/pr_manifest_$$.json)
      url=$(jq -r '.url' /tmp/pr_manifest_$$.json)
      echo "- [#$pr_num]($url): $title (\`${sha:0:9}\`)"
      rm -f /tmp/pr_manifest_$$.json
    else
      echo "- $subject (\`${sha:0:9}\`) -- PR metadata unavailable"
    fi
  done <<< "$merge_commits"
  echo
fi

if [[ -n "$standalone_commits" ]]; then
  echo "### Standalone commits (not part of a merged PR)"
  echo
  while IFS= read -r line; do
    [[ -z "$line" ]] && continue
    sha=${line%% *}
    subject=${line#* }
    echo "- $subject (\`${sha:0:9}\`)"
  done <<< "$standalone_commits"
  echo
fi

if [[ -z "$merge_commits" && -z "$standalone_commits" ]]; then
  echo "(no commits between $base_ref and $head_ref)"
  echo
fi

if [[ -n "$exclude_ref" ]]; then
  echo "## Excluded from this release (on $exclude_ref, not on $head_ref)"
  echo
  excluded_all=$(git log --format="%H %s" "$head_ref..$exclude_ref" 2>/dev/null || true)
  if [[ -n "$excluded_all" ]]; then
    while IFS= read -r line; do
      [[ -z "$line" ]] && continue
      sha=${line%% *}
      subject=${line#* }
      echo "- $subject (\`${sha:0:9}\`)"
    done <<< "$excluded_all"
  else
    echo "(nothing -- $head_ref already contains everything on $exclude_ref since $base_ref)"
  fi
  echo
fi
