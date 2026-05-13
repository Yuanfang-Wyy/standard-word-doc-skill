#!/usr/bin/env bash

set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TEMPLATE="${SKILL_DIR}/assets/standard-word-template.docx"

find_soffice() {
  if command -v soffice >/dev/null 2>&1; then
    command -v soffice
    return 0
  fi
  if command -v libreoffice >/dev/null 2>&1; then
    command -v libreoffice
    return 0
  fi
  if [[ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ]]; then
    printf '%s\n' "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    return 0
  fi
  return 1
}

command -v python3 >/dev/null 2>&1 || {
  echo "missing dependency: python3" >&2
  exit 1
}

[[ -f "${TEMPLATE}" ]] || {
  echo "missing template: ${TEMPLATE}" >&2
  exit 1
}

echo "python3: $(command -v python3)"
echo "template: ${TEMPLATE}"
if SOFFICE="$(find_soffice)"; then
  echo "libreoffice: ${SOFFICE}"
else
  echo "libreoffice: optional dependency not found"
  echo "optional install: brew install --cask libreoffice"
fi
