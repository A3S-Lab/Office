#!/usr/bin/env bash
# Phase 2.2 — reopen real-corpus OOXML exports with LibreOffice headless.
# Skips cleanly when soffice is not on PATH (local macOS without LO).
# CI images that ship LibreOffice must set REQUIRE_LIBREOFFICE=1 so a miss fails.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
FIXTURES="$ROOT/tests/corpus/real/fixtures"
OUT="${TMPDIR:-/tmp}/a3s-office-real-corpus-lo"
REQUIRE="${REQUIRE_LIBREOFFICE:-0}"

find_soffice() {
  if command -v soffice >/dev/null 2>&1; then
    command -v soffice
    return
  fi
  if command -v libreoffice >/dev/null 2>&1; then
    command -v libreoffice
    return
  fi
  local mac="/Applications/LibreOffice.app/Contents/MacOS/soffice"
  if [[ -x "$mac" ]]; then
    echo "$mac"
    return
  fi
  return 1
}

if ! SOFFICE="$(find_soffice)"; then
  if [[ "$REQUIRE" == "1" ]]; then
    echo "LibreOffice (soffice) is required but was not found on PATH." >&2
    exit 1
  fi
  echo "LibreOffice not found; skipping Phase 2.2 independent reopen (set REQUIRE_LIBREOFFICE=1 to fail)."
  exit 0
fi

rm -rf "$OUT"
mkdir -p "$OUT"

shopt -s nullglob
files=("$FIXTURES"/*.{docx,xlsx,pptx})
if ((${#files[@]} == 0)); then
  echo "No OOXML fixtures under $FIXTURES" >&2
  exit 1
fi

failed=0
for src in "${files[@]}"; do
  name="$(basename "$src")"
  echo "LibreOffice reopen: $name"
  if ! "$SOFFICE" --headless --norestore --nolockcheck --nodefault \
    --convert-to pdf --outdir "$OUT" "$src" >/dev/null 2>"$OUT/$name.lo.log"; then
    echo "FAILED: $name (see $OUT/$name.lo.log)" >&2
    failed=1
    continue
  fi
  pdf="$OUT/${name%.*}.pdf"
  if [[ ! -s "$pdf" ]]; then
    echo "FAILED: $name produced no PDF output" >&2
    failed=1
  fi
done

if ((failed)); then
  exit 1
fi
echo "LibreOffice reopen OK for ${#files[@]} OOXML fixture(s)."
