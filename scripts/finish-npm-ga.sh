#!/usr/bin/env bash
# Finish Office GA publish gate once npm Trusted Publisher (or a valid token) works.
# Usage: bash scripts/finish-npm-ga.sh [version]
set -euo pipefail

VERSION="${1:-0.311.2}"
TAG="v${VERSION}"
REG="https://registry.npmjs.org/"

echo "==> Dispatch Publish for ${TAG}"
gh workflow run "Publish npm package" --repo A3S-Lab/Office --ref main -f "ref=${TAG}"

echo "==> Waiting for @a3s-lab/office@${VERSION} on npmjs..."
for i in $(seq 1 60); do
  got="$(npm view "@a3s-lab/office@version" --registry "$REG" 2>/dev/null || true)"
  # npm view pkg@version returns the version field when it exists
  live="$(npm view "@a3s-lab/office" version --registry "$REG" 2>/dev/null || true)"
  echo "  poll ${i}: live=${live}"
  if [ "${live}" = "${VERSION}" ]; then
    echo "==> Registry has ${VERSION}"
    echo "Next: bump packages/ui/site pin to ${VERSION} in the a3s monorepo."
    exit 0
  fi
  sleep 30
done

echo "Timed out waiting for ${VERSION}. Check Trusted Publisher / Office#162." >&2
exit 1
