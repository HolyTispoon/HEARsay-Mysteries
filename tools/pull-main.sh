#!/bin/bash
# Brings the clone up to date with GitHub's main, if it is safe to: only when the clone is on main, and only
# as a fast-forward (git refuses if your own commits or uncommitted edits would conflict). Never pushes.
# Run every 10 minutes on the Mac by the launchd job that tools/setup-drive-clone.sh installs.
set -u
repo="$(cd "$(dirname "$0")/.." && pwd)"
branch="$(git -C "$repo" symbolic-ref --quiet --short HEAD 2>/dev/null || true)"
if [ "$branch" != "main" ]; then
  echo "$(date '+%F %T') skipped: the clone is on '${branch:-a detached commit}', not main"
  exit 0
fi
if out="$(git -C "$repo" pull --ff-only --quiet origin main 2>&1)"; then
  [ -n "$out" ] && echo "$(date '+%F %T') $out"
  exit 0
fi
echo "$(date '+%F %T') not updated: $out"
exit 1
