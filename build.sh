#!/bin/zsh
# Rebuild all AAS PDFs from the editable HTML sources.
# Usage: ./build.sh            (build everything)
#        ./build.sh 04         (build only files starting with 04)
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
SRC="$(cd "$(dirname "$0")" && pwd)"
OUT="$(dirname "$SRC")"

typeset -A DOCS
DOCS=(
  00-start-here                "00 - START HERE - Package Guide"
  01-business-plan             "01 - AAS Business Plan"
  02-executive-summary         "02 - Executive Summary"
  03-sdvosb-launch-guide       "03 - SDVOSB Launch Guide - Step by Step"
  04-capability-statement      "04 - AAS Capability Statement"
  06-90-day-action-plan        "06 - 90-Day Action Plan"
  07-startup-setup-checklist   "07 - AAS Startup Setup Checklist"
)

for slug name in "${(@kv)DOCS}"; do
  if [[ -n "$1" && "$slug" != "$1"* ]]; then continue; fi
  if [[ -f "$SRC/$slug.html" ]]; then
    echo "Rendering $name.pdf"
    "$CHROME" --headless=new --disable-gpu --hide-scrollbars \
      --no-pdf-header-footer --virtual-time-budget=20000 \
      --print-to-pdf="$OUT/$name.pdf" \
      "file://$SRC/$slug.html" 2>/dev/null | tail -1
  fi
done
echo "Done."
