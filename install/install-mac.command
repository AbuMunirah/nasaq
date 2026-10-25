#!/bin/bash
# Installs the Nasaq PowerPoint add-in for the current Mac user.
# Usage: double-click this file, or run it in Terminal. Add --uninstall to remove Nasaq.
set -euo pipefail

MANIFEST_URL="https://abumunirah.github.io/nasaq/manifest.xml"
CONTAINER="$HOME/Library/Containers/com.microsoft.Powerpoint/Data"
WEF="$CONTAINER/Documents/wef"
TARGET="$WEF/nasaq.manifest.xml"

finish() {
  echo
  echo "$1"
  # Keep the window open when the script was double-clicked.
  if [[ -t 0 ]]; then read -r -p "Press Return to close / اضغط Return للإغلاق " _ || true; fi
  exit "${2:-0}"
}

if [[ "${1:-}" == "--uninstall" ]]; then
  rm -f "$TARGET"
  finish "Nasaq was removed. Restart PowerPoint.
تمت إزالة نَسَق. أعد تشغيل PowerPoint."
fi

echo "Installing Nasaq | جارٍ تثبيت نَسَق ..."

if [[ ! -d "/Applications/Microsoft PowerPoint.app" ]]; then
  finish "Microsoft PowerPoint was not found in Applications.
لم يتم العثور على Microsoft PowerPoint في مجلد التطبيقات." 1
fi

if [[ ! -d "$CONTAINER" ]]; then
  finish "Open PowerPoint once, quit it, then run this installer again.
افتح PowerPoint مرة واحدة ثم أغلقه، وبعدها شغّل هذا المثبّت مرة أخرى." 1
fi

mkdir -p "$WEF"
TMP="$(mktemp)"
if ! curl -fsSL "$MANIFEST_URL" -o "$TMP" || ! grep -q "<OfficeApp" "$TMP"; then
  rm -f "$TMP"
  finish "Could not download Nasaq. Check your internet connection and try again.
تعذّر تنزيل نَسَق. تحقق من اتصالك بالإنترنت وحاول مرة أخرى." 1
fi
mv "$TMP" "$TARGET"
chmod 644 "$TARGET"

finish "Nasaq is installed.
1. Quit PowerPoint completely (Cmd+Q) and open it again.
2. Choose Home > Add-ins > My Add-ins > Developer Add-ins > Nasaq.

تم تثبيت نَسَق.
١. أغلق PowerPoint تمامًا (Cmd+Q) ثم افتحه من جديد.
٢. اختر الصفحة الرئيسية ← الوظائف الإضافية ← وظائفي الإضافية ← وظائف المطوّر الإضافية ← نَسَق."
