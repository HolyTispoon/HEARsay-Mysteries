#!/bin/bash
# One-time setup on the Mac: clones the repository into the Google Drive folder "HEARsay mystery game" and
# installs a launchd job that runs tools/pull-main.sh every 10 minutes, so changes merged on GitHub reach
# Drive without anyone copying files.
#
#   bash setup-drive-clone.sh ["/path/to/HEARsay mystery game"]
#   bash setup-drive-clone.sh --uninstall      # removes the launchd job; leaves the clone alone
#
# Before running it: in Google Drive for desktop, right-click "HEARsay mystery game" → Available offline,
# so git works on local files rather than on files Drive downloads on demand.
set -euo pipefail

LABEL="io.github.holytispoon.hearsay-pull"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOG="$HOME/Library/Logs/hearsay-pull.log"
URL="https://github.com/HolyTispoon/HEARsay-Mysteries.git"
NAME="HEARsay-Mysteries"

if [ "${1:-}" = "--uninstall" ]; then
  launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "Removed the scheduled pull ($LABEL). The clone is unchanged."
  exit 0
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
fi

mkdir -p "$(dirname "$PLIST")" "$(dirname "$LOG")"
cat > "$PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>Label</key><string>$LABEL</string>
	<key>ProgramArguments</key>
	<array>
		<string>/bin/bash</string>
		<string>$repo/tools/pull-main.sh</string>
	</array>
	<key>StartInterval</key><integer>600</integer>
	<key>RunAtLoad</key><true/>
	<key>StandardOutPath</key><string>$LOG</string>
	<key>StandardErrorPath</key><string>$LOG</string>
</dict>
</plist>
PLIST
plutil -lint "$PLIST" >/dev/null
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"

echo
echo "Clone:          $repo"
echo "Scheduled pull: every 10 minutes ($PLIST)"
echo "Log:            $LOG"
echo
echo "Check that the scheduled pull can read the Drive folder:"
echo "  launchctl kickstart -k gui/$(id -u)/$LABEL; sleep 5; tail -5 \"$LOG\""
echo "An empty log, or lines without 'not updated', means it works. 'Operation not permitted' means macOS is"
echo "blocking background access to Google Drive: System Settings → Privacy & Security → Full Disk Access → add /bin/bash."
