#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "usage: $0 <output-directory>" >&2
  exit 2
fi

repository_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output_root="$1"

if [[ "$output_root" != /* ]]; then
  output_root="$(pwd)/$output_root"
fi

cargo build \
  --manifest-path "$repository_root/Cargo.toml" \
  --release \
  --package a3s-office-cli

rm -rf "$output_root"
mkdir -p "$output_root/bin" "$output_root/skills"
install -m 0755 \
  "$repository_root/target/release/a3s-office" \
  "$output_root/bin/a3s-office"
cp -R \
  "$repository_root/crates/cli/skills/a3s-office" \
  "$output_root/skills/a3s-office"
cp \
  "$repository_root/integrations/a3s-use/a3s-use-extension.acl" \
  "$output_root/a3s-use-extension.acl"

echo "$output_root"
