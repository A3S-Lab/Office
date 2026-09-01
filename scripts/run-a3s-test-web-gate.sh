#!/usr/bin/env bash
set -euo pipefail

repository_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_root="${A3S_TEST_ARTIFACT_ROOT:-$repository_root/.a3s-test/web-gate}"
preview_log="$artifact_root/preview.log"
fixture_log="$artifact_root/fixtures.log"
preview_pid=""

cd "$repository_root"
mkdir -p "$artifact_root"

require_executable() {
  local executable_name="$1"
  local executable_path
  executable_path="$(command -v "$executable_name" || true)"
  if [[ -z "$executable_path" ]]; then
    echo "Required executable not found: $executable_name" >&2
    exit 1
  fi
  printf '%s' "$executable_path"
}

a3s_test="${A3S_TEST_BIN:-$(require_executable a3s-test)}"
browser_driver="${A3S_TEST_BROWSER_DRIVER:-standalone}"
browser_arguments=(--browser-driver "$browser_driver")
required_a3s_test_version="a3s-test 1.0.0"
required_agent_browser_version="agent-browser 0.26.0"
required_protocol_revision=15
require_executable bun >/dev/null
require_executable curl >/dev/null
require_executable jq >/dev/null

if [[ ! -x "$a3s_test" ]]; then
  echo "A3S Test executable is missing: $a3s_test" >&2
  exit 1
fi

actual_a3s_test_version="$("$a3s_test" --version)"
if [[ "$actual_a3s_test_version" != "$required_a3s_test_version" ]]; then
  echo "Unsupported A3S Test version: $actual_a3s_test_version" >&2
  echo "Expected the local release gate to use: $required_a3s_test_version" >&2
  exit 1
fi

case "$browser_driver" in
  a3s)
    ;;
  standalone)
    agent_browser="${A3S_TEST_AGENT_BROWSER:-$(require_executable agent-browser)}"
    require_executable node >/dev/null

    if [[ ! -x "$agent_browser" ]]; then
      echo "Standalone browser executable is missing: $agent_browser" >&2
      exit 1
    fi

    actual_agent_browser_version="$("$agent_browser" --version)"
    if [[ "$actual_agent_browser_version" != "$required_agent_browser_version" ]]; then
      echo "Unsupported agent-browser version: $actual_agent_browser_version" >&2
      echo "Expected the local release gate to use: $required_agent_browser_version" >&2
      exit 1
    fi

    if [[ -z "${AGENT_BROWSER_EXECUTABLE_PATH:-}" ]]; then
      playwright_browser_path="$(
        node -e \
          "const { chromium } = require('playwright'); process.stdout.write(chromium.executablePath())"
      )"
      if [[ -x "$playwright_browser_path" ]]; then
        AGENT_BROWSER_EXECUTABLE_PATH="$playwright_browser_path"
      else
        AGENT_BROWSER_EXECUTABLE_PATH="$(
          "$agent_browser" doctor --json |
            jq -er '.checks[] | select(.id == "chrome.installed" and .status == "pass") | (.message | capture(" at (?<path>.*)$") | .path)'
        )"
      fi
      export AGENT_BROWSER_EXECUTABLE_PATH
    fi

    if [[ ! -x "$AGENT_BROWSER_EXECUTABLE_PATH" ]]; then
      echo "Chromium executable is missing from Playwright and agent-browser: ${AGENT_BROWSER_EXECUTABLE_PATH:-<unset>}" >&2
      exit 1
    fi

    browser_arguments+=(--browser-executable "$agent_browser")
    ;;
  *)
    echo "Unsupported A3S Test browser driver: $browser_driver" >&2
    exit 1
    ;;
esac

capabilities_result="$artifact_root/capabilities.json"
"$a3s_test" capabilities "${browser_arguments[@]}" --json \
  >"$capabilities_result"
if ! jq -e \
  --arg driver "$browser_driver" \
  --arg standalone_version "${required_agent_browser_version#agent-browser }" \
  --argjson protocol_revision "$required_protocol_revision" \
  '
    .protocol_revision == $protocol_revision and
    if $driver == "standalone" then
      .integration == "standalone" and .version == $standalone_version
    else
      .integration == "a3s"
    end
  ' "$capabilities_result" >/dev/null; then
  echo "A3S Test Web capability contract is incompatible:" >&2
  jq '{integration, version, protocol_revision}' "$capabilities_result" >&2
  exit 1
fi
jq '{integration, version, protocol_revision}' "$capabilities_result"

# E2E fixtures are generated artifacts and are intentionally ignored by Git.
# Generate them as part of the gate so a clean checkout exercises the same
# deterministic upload flows as a developer checkout that has run the fixture
# helper already.
bun scripts/create-e2e-fixtures.ts >"$fixture_log" 2>&1

cleanup() {
  if [[ -n "$preview_pid" ]] && kill -0 "$preview_pid" 2>/dev/null; then
    kill -TERM "$preview_pid" 2>/dev/null || true
    wait "$preview_pid" 2>/dev/null || true
  fi
}

trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM

suites=(
  "tests/e2e/office-docs-navigation.acl"
  "tests/e2e/playground-template-discoverability.acl"
  "tests/e2e/presentation-animation.acl"
  "tests/e2e/pdf-page-organization.acl"
  "tests/e2e/spreadsheet-maximum-sparse.acl"
  "tests/e2e/spreadsheet-cell-fill.acl"
  "tests/e2e/spreadsheet-font-size-border-shortcuts.acl"
  "tests/e2e/spreadsheet-diagonal-borders.acl"
  "tests/e2e/spreadsheet-underline-styles.acl"
  "tests/e2e/spreadsheet-ribbon-orientation-visibility.acl"
  "tests/e2e/spreadsheet-font-colors-shortcuts.acl"
  "tests/e2e/spreadsheet-gradient-fill.acl"
  "tests/e2e/spreadsheet-pattern-fill.acl"
  "tests/e2e/spreadsheet-date-time.acl"
  "tests/e2e/spreadsheet-1904-date-system.acl"
  "tests/e2e/spreadsheet-copy-from-above.acl"
  "tests/e2e/spreadsheet-font-dialog-shortcuts.acl"
  "tests/e2e/spreadsheet-rich-text.acl"
  "tests/e2e/spreadsheet-table.acl"
  "tests/e2e/spreadsheet-table-totals.acl"
  "tests/e2e/spreadsheet-custom-sort.acl"
  "tests/e2e/spreadsheet-custom-list-sort.acl"
  "tests/e2e/spreadsheet-appearance-sort.acl"
  "tests/e2e/spreadsheet-row-sort.acl"
  "tests/e2e/spreadsheet-text-sort.acl"
  "tests/e2e/spreadsheet-sort-range.acl"
  "tests/e2e/spreadsheet-owned-range-sort.acl"
  "tests/e2e/spreadsheet-data-validation.acl"
  "tests/e2e/spreadsheet-hyperlink.acl"
  "tests/e2e/spreadsheet-paste-special.acl"
  "tests/e2e/collaboration-playground-entry.acl"
  "tests/e2e/collaboration-document-suggestions.acl"
  "tests/e2e/word-formatting-revision.acl"
  "tests/e2e/word-paragraph-formatting-revision.acl"
  "tests/e2e/word-document-comparison.acl"
  "tests/e2e/word-emphasis.acl"
  "tests/e2e/word-hidden-text.acl"
  "tests/e2e/word-legacy-text-effects.acl"
  "tests/e2e/word-kerning.acl"
  "tests/e2e/word-character-scale.acl"
  "tests/e2e/word-character-position.acl"
  "tests/e2e/word-character-spacing.acl"
  "tests/e2e/word-proofing-languages.acl"
  "tests/e2e/word-table-of-contents.acl"
  "tests/e2e/word-document-index.acl"
  "tests/e2e/word-script-fonts.acl"
  "tests/e2e/word-opentype-typography.acl"
  "tests/e2e/word-strike-styles.acl"
)

if [[ -n "${A3S_TEST_SUITE:-}" ]]; then
  if [[ "$A3S_TEST_SUITE" != tests/e2e/*.acl || ! -f "$A3S_TEST_SUITE" ]]; then
    echo "A3S_TEST_SUITE must name an existing tests/e2e/*.acl file." >&2
    exit 1
  fi
  suites=("$A3S_TEST_SUITE")
fi

for suite in "${suites[@]}"; do
  suite_name="$(basename "$suite" .acl)"
  check_result="$artifact_root/$suite_name-check.json"
  "$a3s_test" check "$suite" --json >"$check_result"
  jq '{name, scenarios: [.scenarios[].id]}' "$check_result"
done

if curl --fail --silent http://127.0.0.1:4175/ >/dev/null 2>&1; then
  echo "Port 4175 is already serving another process." >&2
  exit 1
fi

bun run playground:preview >"$preview_log" 2>&1 &
preview_pid=$!

preview_ready=false
for _ in {1..200}; do
  if ! kill -0 "$preview_pid" 2>/dev/null; then
    echo "Playground preview exited before it became ready." >&2
    tail -n 120 "$preview_log" >&2
    exit 1
  fi
  if curl --fail --silent http://127.0.0.1:4175/ >/dev/null; then
    preview_ready=true
    break
  fi
  sleep 0.1
done

if [[ "$preview_ready" != true ]]; then
  echo "Playground preview did not become ready within 20 seconds." >&2
  tail -n 120 "$preview_log" >&2
  exit 1
fi

run_suite() {
  local suite="$1"
  local suite_name
  local result_path
  local run_id
  local run_status
  local diagnostics=()

  suite_name="$(basename "$suite" .acl)"
  result_path="$artifact_root/$suite_name-run.json"

  run_status=0
  "$a3s_test" run "$suite" \
    "${browser_arguments[@]}" \
    --command-timeout-ms 120000 \
    --max-parallel-scenarios 1 \
    --infrastructure-retries 2 \
    --retry-backoff-ms 500 \
    --json >"$result_path" || run_status=$?

  jq '{run_id, suite, status, scenarios: [.scenarios[] | {id, status}]}' \
    "$result_path" || true

  if ((run_status != 0)); then
    return "$run_status"
  fi

  run_id="$(jq -er '.run_id' "$result_path")"
  shopt -s nullglob
  diagnostics=("$repository_root/.a3s-test/runs/$run_id"/*/diagnostics/*.json)
  shopt -u nullglob

  if ((${#diagnostics[@]} == 0)); then
    echo "A3S Test produced no browser diagnostics for $suite." >&2
    return 1
  fi

  if ! jq -s -e '
    all(.[];
      if .data.messages? != null then
        .success == true and .error == null and (.data.messages | length == 0)
      elif .data.errors? != null then
        .success == true and .error == null and (.data.errors | length == 0)
      else
        false
      end
    )
  ' "${diagnostics[@]}" >/dev/null; then
    echo "Browser console or page errors were captured for $suite:" >&2
    jq -c '{path: input_filename, data: .data, error: .error}' \
      "${diagnostics[@]}" >&2
    return 1
  fi
}

for suite in "${suites[@]}"; do
  run_suite "$suite"
done
