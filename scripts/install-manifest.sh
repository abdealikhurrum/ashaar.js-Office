#!/usr/bin/env bash
#
# Quick (re)install of the Ashaar add-in manifest for Word on macOS.
#
# Sideloads a manifest into Word's wef folder and clears the Office WebView
# cache so the freshly-served taskpane.html is fetched instead of a stale copy.
# The manifests no longer carry a ?v= cache-buster on the taskpane URL, so
# clearing the WebView cache is the reload mechanism — this script does it.
#
# Usage:
#   scripts/install-manifest.sh              # sideload dev manifest (localhost:3000)
#   scripts/install-manifest.sh prod         # sideload prod manifest (GitHub Pages)
#   scripts/install-manifest.sh path/to.xml  # sideload an explicit manifest
#   scripts/install-manifest.sh dev --restart # also quit + reopen Word
#
# Flags:
#   --restart          quit Word before, reopen after (needed for cache clear to take)
#   --no-cache-clear   skip clearing the Office WebView cache

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

# ── Parse args ────────────────────────────────────────────────────────────────
MANIFEST="manifest.xml"
RESTART=0
CLEAR_CACHE=1

for arg in "$@"; do
  case "$arg" in
    dev)              MANIFEST="manifest.xml" ;;
    prod)             MANIFEST="manifest.prod.xml" ;;
    --restart)        RESTART=1 ;;
    --no-cache-clear) CLEAR_CACHE=0 ;;
    *.xml)            MANIFEST="$arg" ;;
    *) echo "Unknown argument: $arg" >&2; exit 1 ;;
  esac
done

if [[ ! -f "$MANIFEST" ]]; then
  echo "✗ Manifest not found: $MANIFEST" >&2
  exit 1
fi

# ── Extract the add-in <Id> — wef sideload files are named <Id>.manifest.xml ──
ADDIN_ID="$(grep -m1 -oE '<Id>[^<]+</Id>' "$MANIFEST" | sed -E 's/<\/?Id>//g')"
if [[ -z "$ADDIN_ID" ]]; then
  echo "✗ Could not read <Id> from $MANIFEST" >&2
  exit 1
fi

WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
CACHE_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches"

echo "→ Manifest : $MANIFEST"
echo "→ Add-in Id: $ADDIN_ID"
echo "→ wef dir  : $WEF_DIR"

# ── Quit Word if restarting (cache clear only sticks when Word is closed) ─────
if [[ "$RESTART" -eq 1 ]]; then
  echo "→ Quitting Word…"
  osascript -e 'quit app "Microsoft Word"' 2>/dev/null || true
  sleep 2
fi

# ── Clear the Office WebView cache ────────────────────────────────────────────
if [[ "$CLEAR_CACHE" -eq 1 ]]; then
  if pgrep -x "Microsoft Word" >/dev/null 2>&1; then
    echo "⚠ Word is running — cache clear may not take effect. Re-run with --restart."
  fi
  echo "→ Clearing Office WebView cache…"
  rm -rf "$CACHE_DIR"/* 2>/dev/null || true
fi

# ── Sideload: copy the manifest into the wef folder ──────────────────────────
mkdir -p "$WEF_DIR"
cp "$MANIFEST" "$WEF_DIR/$ADDIN_ID.manifest.xml"
echo "✓ Sideloaded to $WEF_DIR/$ADDIN_ID.manifest.xml"

# ── Reopen Word if requested ─────────────────────────────────────────────────
if [[ "$RESTART" -eq 1 ]]; then
  echo "→ Reopening Word…"
  open -a "Microsoft Word"
fi

echo ""
echo "Done. In Word: Home ▸ Add-ins (or Insert ▸ My Add-ins) ▸ open the Ashaar pane."
if [[ "$RESTART" -eq 0 ]]; then
  echo "If you still see old code, quit Word and re-run with --restart."
fi
