#!/bin/bash
# Brings the Drive clone up to date with GitHub's main. Run it by hand, from anywhere:
#
#   bash tools/update-drive.sh            # lists what would change, asks, then updates
#   bash tools/update-drive.sh --check    # only lists what would change
#   bash tools/update-drive.sh --yes      # lists what changes and resets the clone to main without asking
#                                         # (publish.sh runs this)
#   bash tools/update-drive.sh --repo DIR …   # updates the clone in DIR rather than the one holding this script
#
# What it changes: only files inside this repository folder (HEARsay-Mysteries/) that git tracks, and only to
# make them match main on GitHub. It lists every file it would add, change or delete, and the commits behind
# them, before it touches anything. Files outside the repository folder (the Google Doc, the Sheet, Archive/,
# the old Drive folders) and files git doesn't track are never touched. It never pushes anything to GitHub.
#
# Without --yes it refuses, and changes nothing, when the clone is not on main, when a tracked file has been
# edited here and not committed, or when the clone has commits that GitHub's main doesn't have. With --yes it
# resets the clone to main instead: it switches to main, and drops those edits and commits, after listing them.
set -euo pipefail

repo="$(cd "$(dirname "$0")/.." && pwd)"
check_only=false
yes=false
while [ $# -gt 0 ]; do
  case "$1" in
    --check) check_only=true; shift ;;
    --yes) yes=true; shift ;;
    --repo) repo="$(cd "$2" && pwd)"; shift 2 ;;
    *) echo "Unknown option: $1 (see the top of $0)"; exit 1 ;;
  esac
done
g() { git --no-pager -C "$repo" -c core.quotepath=false "$@"; }  # no-pager: never stop for a keypress; quotepath: show "–" in file names, not \342\200\223

if $yes; then
  echo "Fetching main from GitHub …"
  g fetch --quiet origin main
  branch="$(g symbolic-ref --quiet --short HEAD 2>/dev/null || true)"
  [ "$branch" != "main" ] && echo "The clone was on '${branch:-a detached commit}'; switching it to main."
  edited="$(g status --porcelain --untracked-files=no)"
  [ -n "$edited" ] && { echo "Dropping these edits to tracked files, to match main:"; echo "$edited" | sed 's/^/  /'; }
  extra="$(g rev-list origin/main..HEAD 2>/dev/null || true)"
  [ -n "$extra" ] && { echo "Dropping these commits GitHub's main doesn't have:"; g log --oneline origin/main..HEAD | sed 's/^/  /'; }
  before="$(g rev-parse HEAD)"
  if [ "$before" = "$(g rev-parse origin/main)" ] && [ -z "$edited" ] && [ "$branch" = "main" ]; then
    echo "Already up to date with GitHub's main ($(g log -1 --format='%h %s' HEAD))."
    exit 0
  fi
  if [ -z "$extra" ] && [ -n "$(g rev-list "$before..origin/main")" ]; then
    echo
    echo "Changes on GitHub since this copy was updated:"
    g log --reverse --format='  %h %ad  %s' --date=short "$before..origin/main"
  fi
  echo
  echo "Files changed (A added, M changed, D deleted, R renamed):"
  g diff --name-status HEAD origin/main | sed 's/^/  /'
  g switch --quiet --force -C main origin/main
  g reset --quiet --hard origin/main
  echo "Updated to $(g log -1 --format='%h %s' HEAD)."
  exit 0
fi

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
