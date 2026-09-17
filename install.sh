#!/usr/bin/env bash
# Sideloads AI Paste into Word on Mac by copying manifest.xml into Word's
# "wef" folder — the standard sideload location Word watches for add-in
# manifests on macOS. Does not touch certificates, keychain, or any other
# system/security settings.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MANIFEST_SRC="$SCRIPT_DIR/manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

if [ ! -f "$MANIFEST_SRC" ]; then
  echo "error: manifest.xml not found next to install.sh ($MANIFEST_SRC)" >&2
  exit 1
fi

mkdir -p "$WEF_DIR"
cp "$MANIFEST_SRC" "$WEF_DIR/word-ai-paste-manifest.xml"

echo "Installed manifest to: $WEF_DIR/word-ai-paste-manifest.xml"
echo
echo "Next steps:"
echo "  1. Quit Word completely (Cmd+Q) and reopen it."
echo "  2. Open (or create) a document."
echo "  3. Home tab -> \"AI Paste\" button (or Insert -> Add-ins -> My Add-ins -> Developer Add-ins)."
echo
echo "To uninstall: rm \"$WEF_DIR/word-ai-paste-manifest.xml\" and restart Word."
