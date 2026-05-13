#!/usr/bin/env bash

set -euo pipefail

if [[ $# -ne 2 ]]; then
  echo "Usage: finalize_docx.sh <working.docx> <final.docx>" >&2
  exit 1
fi

INPUT="$1"
OUTPUT="$2"

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

[[ -f "${INPUT}" ]] || {
  echo "input not found: ${INPUT}" >&2
  exit 1
}

if ! SOFFICE="$(find_soffice)"; then
  cat >&2 <<'EOF'
missing dependency: LibreOffice

Install on macOS:
  brew install --cask libreoffice

No final .docx was produced.
EOF
  exit 1
fi

WORKDIR="$(mktemp -d)"
trap 'rm -rf "${WORKDIR}"' EXIT

mkdir -p "$(dirname "${OUTPUT}")"
cp "${INPUT}" "${WORKDIR}/input.docx"

"${SOFFICE}" \
  --headless \
  --invisible \
  --nodefault \
  --nofirststartwizard \
  --nolockcheck \
  --nologo \
  --convert-to docx \
  --outdir "${WORKDIR}" \
  "${WORKDIR}/input.docx" >/dev/null

if [[ ! -f "${WORKDIR}/input.docx" ]]; then
  echo "LibreOffice did not produce a finalized docx" >&2
  exit 1
fi

cp "${WORKDIR}/input.docx" "${OUTPUT}"
echo "finalized ${OUTPUT}"

