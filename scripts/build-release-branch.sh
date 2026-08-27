#!/usr/bin/env bash
# Mechanizes building a release branch by cherry-picking a chosen commit list onto a
# base tag/commit -- the exact process elixcee's 0.11.0 and 0.12.0 releases both needed
# (cutting from a release branch instead of master's tip, to keep still-unverified
# master-only work like 0.10.0-D/t="e" out of the release), previously done by hand or
# via a long natural-language prompt to an agent. Shrinks ROADMAP.md's disclosed
# "manual cherry-pick dependency" gap: the caller still has to DECIDE which commits
# belong in the release (that's a judgment call this script deliberately doesn't make),
# but no longer has to hand-run and hand-order each `git cherry-pick` themselves.
#
# Usage: scripts/build-release-branch.sh <base-ref> <new-branch> <commit>...
#
#   base-ref    a tag or commit the new branch starts from (e.g. v0.12.0)
#   new-branch  name for the new branch (must not already exist)
#   commit...   one or more commits to cherry-pick, in the order given -- a merge
#               commit (detected automatically) is cherry-picked with `-m 1`
#               (its first-parent diff), matching how this project's own PRs are
#               merged; anything else is cherry-picked as a plain commit
#
# On the first conflict, stops and leaves the repository mid-cherry-pick for manual
# resolution -- this script automates the mechanical parts, not conflict resolution
# itself, which needs human/agent judgment about which side's change is correct (see
# `git status`'s own instructions: `git cherry-pick --continue` or `--abort`).
#
# After every commit applies cleanly, runs a fast local verification subset (fmt,
# clippy, check-versions.sh) and reports it -- NOT the full sweep this project also
# runs before a real release (corpus/semantics/differential suites), which takes much
# longer and stays the caller's job.
set -euo pipefail

cd "$(dirname "$0")/.."

if [[ $# -lt 3 ]]; then
  echo "usage: $0 <base-ref> <new-branch> <commit>..." >&2
  exit 1
fi

base_ref=$1
new_branch=$2
shift 2
commits=("$@")

if [[ -n "$(git status --porcelain)" ]]; then
  echo "build-release-branch: working tree is not clean -- commit, stash, or discard changes first" >&2
  exit 1
fi

if git rev-parse --verify --quiet "refs/heads/$new_branch" >/dev/null; then
  echo "build-release-branch: branch '$new_branch' already exists -- pick a new name or delete it first" >&2
  exit 1
fi

if ! git rev-parse --verify --quiet "$base_ref^{commit}" >/dev/null; then
  echo "build-release-branch: '$base_ref' does not resolve to a commit" >&2
  exit 1
fi

echo "build-release-branch: creating '$new_branch' from '$base_ref'"
git branch "$new_branch" "$base_ref"
git checkout "$new_branch"

for commit in "${commits[@]}"; do
  resolved=$(git rev-parse --verify --quiet "${commit}^{commit}") || {
    echo "build-release-branch: '$commit' does not resolve to a commit -- branch '$new_branch' left as-is for inspection" >&2
    exit 1
  }
  parent_count=$(git cat-file -p "$resolved" | grep -c '^parent ' || true)

  if [[ "$parent_count" -ge 2 ]]; then
    echo "build-release-branch: cherry-picking ${commit} (merge commit, -m 1)"
    if ! git cherry-pick -m 1 "$commit"; then
      echo "build-release-branch: conflict on ${commit} -- resolve it, then 'git cherry-pick --continue', or 'git cherry-pick --abort' to stop" >&2
      exit 1
    fi
  else
    echo "build-release-branch: cherry-picking ${commit}"
    if ! git cherry-pick "$commit"; then
      echo "build-release-branch: conflict on ${commit} -- resolve it, then 'git cherry-pick --continue', or 'git cherry-pick --abort' to stop" >&2
      exit 1
    fi
  fi
done

echo "build-release-branch: all ${#commits[@]} commit(s) applied cleanly onto '$new_branch'"
echo "build-release-branch: running fast local verification (fmt, clippy, check-versions.sh)..."

verification_failed=0

if ! cargo fmt --all --check; then
  echo "build-release-branch: cargo fmt --check FAILED" >&2
  verification_failed=1
fi

if ! cargo clippy --workspace --all-targets -- -D warnings; then
  echo "build-release-branch: cargo clippy FAILED" >&2
  verification_failed=1
fi

if ! ./scripts/check-versions.sh; then
  echo "build-release-branch: check-versions.sh FAILED" >&2
  verification_failed=1
fi

if [[ "$verification_failed" -ne 0 ]]; then
  echo "build-release-branch: branch '$new_branch' built, but fast verification found problems above -- fix them before running the full test/differential sweep" >&2
  exit 1
fi

echo "build-release-branch: fast verification passed. Still needed before this is release-ready:"
echo "  - the full test suite (cargo test --workspace) and RUSTDOCFLAGS doc build"
echo "  - compat/corpus, compat/vba-semantics, the 5 JS differential suites, both differential-python scripts"
echo "  - a version bump + CHANGELOG.md reorganization (this script does neither)"
echo "  - explicit approval before any push, tag, or publish -- this script does none of those either"
