#!/usr/bin/env bash
set -euo pipefail

a3s_test_version="${A3S_TEST_VERSION:-v0.16.2}"
agent_browser_version="${AGENT_BROWSER_VERSION:-0.26.0}"
runner_temp="${RUNNER_TEMP:-}"

if [[ -z "$runner_temp" ]]; then
  echo "RUNNER_TEMP must be set for the CI-only A3S Test installation." >&2
  exit 1
fi

install_dir="${A3S_TEST_INSTALL_DIR:-$runner_temp/a3s-test/bin}"
installer_url="https://github.com/A3S-Lab/Test/releases/download/$a3s_test_version/install.sh"

curl --fail --silent --show-error --location --retry 3 "$installer_url" |
  sh -s -- \
    --version "$a3s_test_version" \
    --cli-only \
    --install-dir "$install_dir"

npm install --global "agent-browser@$agent_browser_version"

"$install_dir/a3s-test" --version
agent-browser --version

if [[ -n "${GITHUB_PATH:-}" ]]; then
  printf '%s\n' "$install_dir" >>"$GITHUB_PATH"
fi
