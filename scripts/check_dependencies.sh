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

python3 - <<'PY' >/dev/null 2>&1 || {
import docx
PY
  echo "missing dependency: python-docx" >&2
  echo "install: pip install python-docx" >&2
  exit 1
}

echo "python3: $(command -v python3)"
python3 - <<'PY'
import docx
print(f"python-docx: {docx.__version__}")
PY
if [[ -f "${TEMPLATE}" ]]; then
  echo "template: ${TEMPLATE}"
else
  echo "template: optional dependency not found"
  echo "fallback: scripts will create a DOCX from the default style guide"
fi
if SOFFICE="$(find_soffice)"; then
  echo "libreoffice: ${SOFFICE}"
else
  echo "libreoffice: optional dependency not found"
  echo "optional install: brew install --cask libreoffice"
fi
