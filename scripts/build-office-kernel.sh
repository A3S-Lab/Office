#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TARGET="wasm32-unknown-unknown"
OUTPUT_DIR="${ROOT_DIR}/generated"
SOURCE="${ROOT_DIR}/target/${TARGET}/release/a3s_office_web_kernel.wasm"

if ! rustup target list --installed | grep -qx "${TARGET}"; then
  echo "Missing Rust target ${TARGET}; run: rustup target add ${TARGET}" >&2
  exit 1
fi

cargo build \
  --manifest-path "${ROOT_DIR}/Cargo.toml" \
  --package a3s-office-web-kernel \
  --target "${TARGET}" \
  --release

mkdir -p "${OUTPUT_DIR}"
cp "${SOURCE}" "${OUTPUT_DIR}/office-kernel.wasm"
echo "Built ${OUTPUT_DIR}/office-kernel.wasm"
