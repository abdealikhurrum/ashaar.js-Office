#!/usr/bin/env bash
#
# Ashaar Poetry — one-click installer for Word on macOS.
# Double-click this file (if macOS blocks it: right-click ▸ Open ▸ Open).
#
# Downloads the current add-in manifest, sideloads it into Word, and clears the
# Office WebView cache so you get the latest pane — superseding any old copy.

set -euo pipefail

ID="a03d69e0-06d9-47c9-8995-06d3bdb5b197"
URL="https://abdealikhurrum.github.io/ashaar.js-Office/manifest.prod.xml"
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
CACHE="$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches"

echo "── Ashaar Poetry — installing add-in for Word ──"
echo ""

echo "→ Closing Word…"
osascript -e 'quit app "Microsoft Word"' 2>/dev/null || true
sleep 2

echo "→ Downloading manifest…"
mkdir -p "$WEF"
if ! curl -fsSL "$URL" -o "$WEF/$ID.manifest.xml"; then
  echo "✗ Download failed. Check your internet connection and try again."
  echo "  (Press any key to close.)"; read -n 1 -s; exit 1
fi

echo "→ Clearing Word's add-in cache…"
rm -rf "$CACHE"/* 2>/dev/null || true

echo "→ Reopening Word…"
open -a "Microsoft Word" 2>/dev/null || true

echo ""
echo "✓ Installed."
echo "  In Word, open the pane from:  Home ▸ Add-ins ▸ Ashaar Poetry"
echo "  (or Insert ▸ My Add-ins ▸ Ashaar Poetry)"
echo ""
echo "You can close this window."
