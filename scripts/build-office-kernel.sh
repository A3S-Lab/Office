#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TARGET="wasm32-unknown-unknown"
OUTPUT_DIR="${ROOT_DIR}/generated"
SOURCE="${ROOT_DIR}/target/${TARGET}/release/a3s_office_web_kernel.wasm"

# Git Bash on Windows can expose a separate Linux rustup first on PATH. Prefer
# the native Windows toolchain when its executable is available so the package
# build sees the same installed targets as PowerShell and CI.
if command -v rustup.exe >/dev/null 2>&1; then
  RUSTUP=rustup.exe
  CARGO=cargo.exe
  if command -v cygpath >/dev/null 2>&1; then
    MANIFEST_PATH="$(cygpath -w "${ROOT_DIR}/Cargo.toml")"
  else
    MANIFEST_PATH="$(wslpath -w "${ROOT_DIR}/Cargo.toml")"
  fi
else
  RUSTUP=rustup
  CARGO=cargo
  MANIFEST_PATH="${ROOT_DIR}/Cargo.toml"
fi

if ! "${RUSTUP}" target list --installed | grep -qx "${TARGET}"; then
  echo "Missing Rust target ${TARGET}; run: rustup target add ${TARGET}" >&2
  exit 1
fi

"${CARGO}" build \
  --manifest-path "${MANIFEST_PATH}" \
  --package a3s-office-web-kernel \
  --target "${TARGET}" \
  --release

mkdir -p "${OUTPUT_DIR}"
cp "${SOURCE}" "${OUTPUT_DIR}/office-kernel.wasm"
echo "Built ${OUTPUT_DIR}/office-kernel.wasm"
