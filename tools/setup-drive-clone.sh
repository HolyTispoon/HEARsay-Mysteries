#!/bin/bash
# One-time setup on the Mac: clones the repository into the Google Drive folder "HEARsay mystery game", as
# HEARsay-Mysteries/. Afterwards, bring it up to date by hand with tools/update-drive.sh.
#
#   bash setup-drive-clone.sh ["/path/to/HEARsay mystery game"]
#
# Before running it: in Google Drive for desktop, right-click "HEARsay mystery game" → Available offline,
# so git works on local files rather than on files Drive downloads on demand.
#
# An earlier version installed a launchd job that pulled every 10 minutes. This version removes that job if
# it is still installed.
set -euo pipefail

LABEL="io.github.holytispoon.hearsay-pull"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
URL="https://github.com/HolyTispoon/HEARsay-Mysteries.git"
NAME="HEARsay-Mysteries"

if [ -f "$PLIST" ]; then
  launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "Removed the old scheduled pull ($LABEL)."
fi

command -v git >/dev/null || { echo "git is missing. Run: xcode-select --install"; exit 1; }

drive="${1:-}"
if [ -z "$drive" ]; then
  # Google Drive for desktop mounts each account at ~/Library/CloudStorage/GoogleDrive-<email>/
  drive="$(find "$HOME/Library/CloudStorage" -maxdepth 4 -type d -name "HEARsay mystery game" 2>/dev/null | head -1)"
fi
if [ -z "$drive" ] || [ ! -d "$drive" ]; then
  echo "Couldn't find the Drive folder \"HEARsay mystery game\". Pass its path: bash $0 \"/path/to/HEARsay mystery game\""
  exit 1
fi
repo="$drive/$NAME"

if [ -d "$repo/.git" ]; then
  echo "Already cloned: $repo"
elif [ -e "$repo" ]; then
  echo "$repo exists but is not a git clone. Move it away, then run this again."
  exit 1
else
  git clone "$URL" "$repo"
  echo "Cloned: $repo"
fi
echo "To bring it up to date with GitHub later: bash \"$repo/tools/update-drive.sh\""
