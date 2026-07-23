#!/usr/bin/env bash
set -euo pipefail

repository_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
skill_root="$repository_root/crates/cli/skills"
output_root="$repository_root/playground/generated"
archive="$output_root/a3s-office-skill.tar.gz"
temporary_archive="$archive.tmp"

mkdir -p "$output_root"
rm -f "$temporary_archive"
tar \
  --exclude=".DS_Store" \
  -czf "$temporary_archive" \
  -C "$skill_root" \
  a3s-office
mv "$temporary_archive" "$archive"

echo "Packaged $archive"
