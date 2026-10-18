#!/bin/bash
# Installs what engine/build.py needs in a Claude Code cloud session: Tweego 2.1.1, SugarCube 2.37.3,
# openpyxl, python-docx and pypdf. Sets TWEEGO and TWEEGO_PATH for the session. Safe to run again.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

TOOLS="$HOME/.cache/hearsay-build-tools"
TWEEGO_ZIP="https://github.com/tmedwards/tweego/releases/download/v2.1.1/tweego-2.1.1-linux-x64.zip"
TWEEGO_SHA=2ed11c3aa66c88ec5382b51061c7782152632b88916777428ab57a66feb7e5aa
SUGARCUBE_ZIP="https://github.com/tmedwards/sugarcube-2/releases/download/v2.37.3/sugarcube-2.37.3-for-twine-2.1-local.zip"
SUGARCUBE_SHA=da00a8c15ec4e88a9e231a3ff6c516c57055f84231bb999f869ed34ade353dab

fetch() {  # url sha256 destination-folder
  local zip
  zip="$(mktemp --suffix=.zip)"
  curl -fsSL -o "$zip" "$1"
  echo "$2  $zip" | sha256sum -c --quiet -
  mkdir -p "$3"
  unzip -q -o "$zip" -d "$3"
  rm -f "$zip"
}

if [ ! -x "$TOOLS/tweego/tweego" ]; then
  fetch "$TWEEGO_ZIP" "$TWEEGO_SHA" "$TOOLS/tweego"
  chmod +x "$TOOLS/tweego/tweego"
fi
# Tweego 2.1.1 bundles an older SugarCube; TWEEGO_PATH points it at 2.37.3, which the games are built with
if [ ! -f "$TOOLS/storyformats/sugarcube-2/format.js" ]; then
  fetch "$SUGARCUBE_ZIP" "$SUGARCUBE_SHA" "$TOOLS/storyformats"
fi

if ! python3 -c "import openpyxl, docx, pypdf" 2>/dev/null; then
  pip3 install --quiet --break-system-packages openpyxl python-docx pypdf
fi

if [ -n "${CLAUDE_ENV_FILE:-}" ]; then
  {
    echo "export TWEEGO=\"$TOOLS/tweego/tweego\""
    echo "export TWEEGO_PATH=\"$TOOLS/storyformats\""
  } >> "$CLAUDE_ENV_FILE"
fi
