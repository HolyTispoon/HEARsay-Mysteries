#!/bin/bash
# Brings the Drive clone up to date with GitHub's main. Run it by hand, from anywhere:
#
#   bash tools/update-drive.sh            # lists what would change, asks, then updates
#   bash tools/update-drive.sh --check    # only lists what would change
#
# What it changes: only files inside this repository folder (HEARsay-Mysteries/) that git tracks, and only to
# make them match main on GitHub. It lists every file it would add, change or delete, and the commits behind
# them, before it touches anything. Files outside the repository folder (the Google Doc, the Sheet, Archive/,
# the old Drive folders) and files git doesn't track are never touched. It never pushes anything to GitHub.
#
# It refuses, and changes nothing, when the clone is not on main, when a tracked file has been edited here and
# not committed, or when the clone has commits that GitHub's main doesn't have.
set -euo pipefail

repo="$(cd "$(dirname "$0")/.." && pwd)"
check_only=false
[ "${1:-}" = "--check" ] && check_only=true
g() { git -C "$repo" -c core.quotepath=false "$@"; }  # quotepath: show "–" in file names, not \342\200\223

branch="$(g symbolic-ref --quiet --short HEAD 2>/dev/null || true)"
if [ "$branch" != "main" ]; then
  echo "Not updated: the clone is on '${branch:-a detached commit}', not main. Switch with: git -C \"$repo\" switch main"
  exit 1
fi
edited="$(g status --porcelain --untracked-files=no)"
if [ -n "$edited" ]; then
  echo "Not updated: these files were edited here and not committed. Commit them, or undo the edits, first:"
  echo "$edited"
  exit 1
fi

echo "Fetching main from GitHub …"
g fetch --quiet origin main
if [ -n "$(g rev-list origin/main..HEAD)" ]; then
  echo "Not updated: the clone has commits GitHub's main doesn't have:"
  g log --oneline origin/main..HEAD
  exit 1
fi
if [ -z "$(g rev-list HEAD..origin/main)" ]; then
  echo "Already up to date with GitHub's main ($(g log -1 --format='%h %s' HEAD))."
  exit 0
fi

echo
echo "Changes on GitHub since this copy was updated:"
g log --reverse --format='  %h %ad  %s' --date=short HEAD..origin/main
echo
echo "Files that will change (A added, M changed, D deleted, R renamed):"
g diff --name-status HEAD origin/main | sed 's/^/  /'
echo

if $check_only; then
  echo "Nothing changed (--check)."
  exit 0
fi
read -r -p "Update these files in $repo? [y/N] " answer
case "$answer" in
  y|Y|yes|Yes) ;;
  *) echo "Nothing changed."; exit 0 ;;
esac
g merge --ff-only --quiet origin/main
echo "Updated to $(g log -1 --format='%h %s' HEAD)."
