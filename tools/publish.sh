#!/bin/bash
# Rebuilds every game from the Google files and publishes it, in one run. Run it by hand on the Mac, from anywhere:
#
#   bash tools/publish.sh                        # downloads the Google files through your browser, builds, publishes
#   bash tools/publish.sh --downloads FOLDER     # uses Google files you already downloaded into FOLDER
#   bash tools/publish.sh --check                # downloads and builds, lists what would change, publishes nothing
#
# What it does, in order (it stops at the first problem and publishes nothing):
#   1. The first time only: installs Tweego 2.1.1, SugarCube 2.37.3 and the Python packages the build needs into
#      ~/.cache/hearsay-build-tools (checked against their published checksums). Nothing else on the Mac changes.
#   2. Opens each Google file's download link in your browser (the framework Doc, the game text Sheet, and each
#      scenario's matrix and packet), waits for the downloads to land in ~/Downloads, and moves them into a new
#      folder in ~/.cache/hearsay-publish/downloads.
#   3. Builds in its own copy of the repository (~/.cache/hearsay-publish/HEARsay-Mysteries, reset to GitHub's main
#      each run), never in the Drive clone. The build runs every check and stops if one fails.
#   4. Lists every file the rebuild changed. If nothing changed, it stops there.
#   5. Publishes: commits to a new branch, opens a pull request, and merges it into main. GitHub Pages serves the
#      new games a minute or two later.
#   6. Runs tools/update-drive.sh, which lists what it would change in the Drive clone and asks first.
#
# Needs the GitHub CLI signed in (gh auth status), and your browser signed in to the Google account that owns the files.
set -euo pipefail

downloads=""
check_only=false
while [ $# -gt 0 ]; do
  case "$1" in
    --downloads) downloads="$(cd "$2" && pwd)"; shift 2 ;;
    --check) check_only=true; shift ;;
    *) echo "Unknown option: $1 (see the top of $0)"; exit 1 ;;
  esac
done

here="$(cd "$(dirname "$0")" && pwd)"
tools="$HOME/.cache/hearsay-build-tools"
work="$HOME/.cache/hearsay-publish"
repo="$work/HEARsay-Mysteries"
remote="https://github.com/HolyTispoon/HEARsay-Mysteries.git"

TWEEGO_ZIP="https://github.com/tmedwards/tweego/releases/download/v2.1.1/tweego-2.1.1-macos-x64.zip"
TWEEGO_SHA=93d8da9df25e6b08d9011175ecebe67bef76a639f3aa3b20b5deefb691316ef1
SUGARCUBE_ZIP="https://github.com/tmedwards/sugarcube-2/releases/download/v2.37.3/sugarcube-2.37.3-for-twine-2.1-local.zip"
SUGARCUBE_SHA=da00a8c15ec4e88a9e231a3ff6c516c57055f84231bb999f869ed34ade353dab

command -v gh >/dev/null || { echo "Needs the GitHub CLI: brew install gh, then gh auth login"; exit 1; }
gh api user --jq .login >/dev/null 2>&1 || { echo "The GitHub CLI isn't signed in. Run: gh auth login"; exit 1; }

# 1. build tools
fetch() {  # url sha256 destination-folder
  local zip
  zip="$(mktemp -t hearsay)"
  curl -fsSL -o "$zip" "$1"
  echo "$2  $zip" | shasum -a 256 -c --quiet -
  mkdir -p "$3"
  unzip -q -o "$zip" -d "$3"
  rm -f "$zip"
}
if [ ! -x "$tools/tweego/tweego" ]; then
  echo "Installing Tweego 2.1.1 into $tools …"
  fetch "$TWEEGO_ZIP" "$TWEEGO_SHA" "$tools/tweego"
  chmod +x "$tools/tweego/tweego"
fi
# Tweego 2.1.1 bundles an older SugarCube; TWEEGO_PATH points it at 2.37.3, which the games are built with
if [ ! -f "$tools/storyformats/sugarcube-2/format.js" ]; then
  echo "Installing SugarCube 2.37.3 into $tools …"
  fetch "$SUGARCUBE_ZIP" "$SUGARCUBE_SHA" "$tools/storyformats"
fi
if ! "$tools/venv/bin/python" -c "import openpyxl, docx, pypdf" 2>/dev/null; then
  echo "Installing the build's Python packages into $tools/venv …"
  python3 -m venv "$tools/venv"
  "$tools/venv/bin/pip" install --quiet openpyxl python-docx pypdf
fi
py="$tools/venv/bin/python"
export TWEEGO="$tools/tweego/tweego" TWEEGO_PATH="$tools/storyformats"

# 3 (first half). a fresh copy of main, so the download list below matches what will be built
mkdir -p "$work"
if [ ! -d "$repo/.git" ]; then
  git clone --quiet "$remote" "$repo"
fi
g() { git -C "$repo" -c core.quotepath=false "$@"; }
echo "Fetching main from GitHub …"
g fetch --quiet origin main
g switch --quiet --force -C main origin/main
g clean --quiet -fd

# 2. downloads
if [ -z "$downloads" ]; then
  downloads="$work/downloads/$(date +%Y-%m-%d-%H%M%S)"
  mkdir -p "$downloads"
  marker="$downloads/.started"
  touch "$marker"
  # one line per Google file: the name it downloads as (no extension), the extension, the download link
  list="$("$py" - "$repo" <<'EOF'
import glob, json, os, re, sys
root = sys.argv[1]
sys.path.insert(0, os.path.join(root, "engine"))
from sources import copies
ids = sorted(n for n in os.listdir(os.path.join(root, "scenarios"))
             if os.path.isdir(os.path.join(root, "scenarios", n)) and not n.startswith("."))
def link(url, ext):
    m = re.search(r"/(document|spreadsheets)/d/([^/]+)", url)
    return f"https://docs.google.com/{m.group(1)}/d/{m.group(2)}/export?format={ext[1:]}" if m else None
rows = []
for label, path, url in copies(ids):
    stem, ext = os.path.splitext(os.path.basename(path))
    rows.append((stem, ext, link(url, ext)))
for stem, ext, url in rows:
    if url:
        print(f"{stem}\t{ext}\t{url}")
EOF
)"
  echo
  echo "Downloading the Google files through your browser (into ~/Downloads, then moved to $downloads):"
  while IFS=$'\t' read -r stem ext url; do
    echo "  $stem$ext"
    open "$url"
  done <<< "$list"
  echo "Waiting for the downloads (up to 3 minutes) …"
  for _ in $(seq 90); do
    missing=0
    while IFS=$'\t' read -r stem ext url; do
      found="$(find "$HOME/Downloads" -maxdepth 1 -newer "$marker" -name "$stem*$ext" ! -name "*.crdownload" 2>/dev/null | head -1)"
      [ -n "$found" ] || missing=1
    done <<< "$list"
    [ $missing = 0 ] && break
    sleep 2
  done
  while IFS=$'\t' read -r stem ext url; do
    found="$(find "$HOME/Downloads" -maxdepth 1 -newer "$marker" -name "$stem*$ext" ! -name "*.crdownload" 2>/dev/null | head -1)"
    if [ -z "$found" ]; then
      echo "Not published: $stem$ext didn't download. Check the browser is signed in to Google, then run this again."
      exit 1
    fi
    mv "$found" "$downloads/$stem$ext"
  done <<< "$list"
  rm -f "$marker"
  echo "All downloads are in $downloads"
fi

# 3. build
echo
echo "Building …"
if ! "$py" "$repo/engine/build.py" --downloads "$downloads"; then
  echo
  echo "Not published: the build stopped (see above). Fix the Google files, then run this again."
  exit 1
fi

# 4. what changed
changes="$(g status --porcelain)"
if [ -z "$changes" ]; then
  echo
  echo "Nothing to publish: the published games already match the Google files."
  exit 0
fi
echo
echo "Files the rebuild changed (M changed, ?? new, D deleted):"
g status --short | sed 's/^/  /'
echo
if $check_only; then
  echo "Nothing published (--check)."
  exit 0
fi

# 5. publish
branch="publish/$(date +%Y-%m-%d-%H%M%S)"
title="Rebuild from the Google files ($(date +%Y-%m-%d))"
g switch --quiet -c "$branch"
g add -A
g commit --quiet -m "$title"
g push --quiet -u origin "$branch"
body="Published with \`tools/publish.sh\`: rebuilt every game from fresh downloads of the Google files.

Files changed:
\`\`\`
$(g diff --name-status HEAD~1 HEAD)
\`\`\`"
url="$(cd "$repo" && gh pr create --base main --head "$branch" --title "$title" --body "$body")"
(cd "$repo" && gh pr merge "$url" --merge --delete-branch >/dev/null)
g switch --quiet main
echo "Published: $url"
echo "GitHub Pages serves the new games in a minute or two: https://holytispoon.github.io/HEARsay-Mysteries/"

# 6. the Drive clone
if [ -f "$here/update-drive.sh" ] && [ "$(cd "$here/.." && pwd)" != "$repo" ]; then
  echo
  bash "$here/update-drive.sh"
fi
