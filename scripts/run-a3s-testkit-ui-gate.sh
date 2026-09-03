#!/usr/bin/env bash
set -euo pipefail

repository_root="$(cd "$(dirname "$0")/.." && pwd)"
artifact_root="${A3S_TESTKIT_ARTIFACT_ROOT:-$repository_root/.a3s-test/testkit-gate}"
run_result="$artifact_root/office-testkit-ui-run.json"
check_result="$artifact_root/office-testkit-ui-check.json"
server_log="$artifact_root/playground.log"
warmup_log="$artifact_root/warmup.log"
server_pid=""
owns_server=false

cd "$repository_root"
mkdir -p "$artifact_root"

resolve_a3s_test() {
  local configured="${A3S_TEST_BIN:-}"
  local local_release="$repository_root/../../crates/test/target/release/a3s-test"
  local local_debug="$repository_root/../../crates/test/target/debug/a3s-test"
  local candidate
  if [[ -n "$configured" ]]; then
    [[ -x "$configured" ]] || { echo "A3S_TEST_BIN is not executable: $configured" >&2; return 1; }
    printf '%s' "$configured"
    return 0
  fi
  for candidate in "$local_release" "$local_debug"; do
    if [[ -x "$candidate" ]]; then
      printf '%s' "$candidate"
      return 0
    fi
  done
  candidate="$(command -v a3s-test || true)"
  if [[ -n "$candidate" && -x "$candidate" ]]; then
    printf '%s' "$candidate"
    return 0
  fi
  echo "Required executable not found: a3s-test" >&2
  return 1
}

resolve_agent_browser() {
  local configured="${A3S_TEST_AGENT_BROWSER:-}"
  local candidate
  if [[ -n "$configured" ]]; then
    printf '%s' "$configured"
    return 0
  fi
  candidate="$(command -v agent-browser || true)"
  if [[ -n "$candidate" && -x "$candidate" ]]; then
    printf '%s' "$candidate"
    return 0
  fi
  if [[ -x "$HOME/.bun/bin/agent-browser" ]]; then
    printf '%s' "$HOME/.bun/bin/agent-browser"
    return 0
  fi
  echo "Required executable not found: agent-browser" >&2
  return 1
}

a3s_test="$(resolve_a3s_test)"
agent_browser="$(resolve_agent_browser)"
[[ "$("$a3s_test" --version)" == "a3s-test 1.0.0" ]] || {
  echo "The Test Kit gate requires a3s-test 1.0.0 (build crates/test)." >&2
  exit 1
}
[[ "$("$agent_browser" --version)" == "agent-browser 0.26.0" ]] || {
  echo "The Test Kit gate requires agent-browser 0.26.0." >&2
  exit 1
}
for required_tool in bun curl jq node; do
  command -v "$required_tool" >/dev/null 2>&1 || {
    echo "Required executable not found: $required_tool" >&2
    exit 1
  }
done

if [[ -z "${AGENT_BROWSER_EXECUTABLE_PATH:-}" ]]; then
  chromium_path="$(node -e "const { chromium } = require('playwright'); process.stdout.write(chromium.executablePath())")"
  if [[ -x "$chromium_path" ]]; then
    AGENT_BROWSER_EXECUTABLE_PATH="$chromium_path"
  else
    AGENT_BROWSER_EXECUTABLE_PATH="$("$agent_browser" doctor --json | jq -er '.checks[] | select(.id == "chrome.installed" and .status == "pass") | (.message | capture(" at (?<path>.*)$") | .path)')"
  fi
  export AGENT_BROWSER_EXECUTABLE_PATH
fi
[[ -x "$AGENT_BROWSER_EXECUTABLE_PATH" ]] || {
  echo "Chromium executable is missing: $AGENT_BROWSER_EXECUTABLE_PATH" >&2
  exit 1
}

"$a3s_test" capabilities \
  --browser-driver standalone \
  --browser-executable "$agent_browser" \
  --command-timeout-ms 90000 \
  --json >"$artifact_root/capabilities.json"
jq -e '.integration == "standalone" and .version == "0.26.0" and .protocol_revision == 15' \
  "$artifact_root/capabilities.json" >/dev/null
jq '{integration, version, protocol_revision}' "$artifact_root/capabilities.json"

bun scripts/create-e2e-fixtures.ts >"$artifact_root/fixtures.log" 2>&1
testkit_url="http://127.0.0.1:3000/playground/"

if ! curl --fail --silent "$testkit_url" >/dev/null 2>&1; then
  bun run playground >"$server_log" 2>&1 &
  server_pid=$!
  owns_server=true
  server_ready=false
  for _ in {1..300}; do
    if ! kill -0 "$server_pid" 2>/dev/null; then
      echo "Playground development server exited before becoming ready." >&2
      tail -n 120 "$server_log" >&2
      exit 1
    fi
    if curl --fail --silent "$testkit_url" >/dev/null 2>&1; then
      server_ready=true
      break
    fi
    sleep 0.2
  done
  [[ "$server_ready" == true ]] || {
    echo "Playground development server did not become ready within 60 seconds." >&2
    tail -n 120 "$server_log" >&2
    exit 1
  }
fi

cleanup() {
  if [[ "$owns_server" == true && -n "$server_pid" ]] && kill -0 "$server_pid" 2>/dev/null; then
    kill -TERM "$server_pid" 2>/dev/null || true
    wait "$server_pid" 2>/dev/null || true
  fi
}
trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM

# Rsbuild compiles the lazy editor workspace on first pointer entry in dev.
# Warm those chunks before the ACL runner clicks a template; otherwise the
# first HMR rebuild can reload the page between the click and the editor's
# readiness assertion.
A3S_TESTKIT_BASE_URL="$testkit_url" \
  AGENT_BROWSER_EXECUTABLE_PATH="$AGENT_BROWSER_EXECUTABLE_PATH" \
  bun scripts/warm-playground-editor.ts >"$warmup_log" 2>&1

"$a3s_test" check tests/e2e/office-testkit-ui.acl --json >"$check_result"
jq '{name, scenarios: [.scenarios[].id]}' "$check_result"

run_status=0
"$a3s_test" run tests/e2e/office-testkit-ui.acl \
  --browser-driver standalone \
  --browser-executable "$agent_browser" \
  --command-timeout-ms 90000 \
  --idle-timeout-ms 300000 \
  --max-parallel-scenarios 1 \
  --infrastructure-retries 2 \
  --retry-backoff-ms 500 \
  --json >"$run_result" || run_status=$?
jq '{run_id, suite, status, scenarios: [.scenarios[] | {id, status}]}' "$run_result"
if ((run_status != 0)); then
  exit "$run_status"
fi

snapshot_count="$(jq '[.scenarios[].steps[] | select(.output.page_context != null)] | length' "$run_result")"
if ((snapshot_count < 9)); then
  echo "Test Kit gate captured only $snapshot_count valid page-context snapshots; expected at least 9." >&2
  exit 1
fi
jq -e '
  .status == "passed" and
  all(.scenarios[]; .status == "passed") and
  ([.scenarios[].steps[] | select(.output.page_context != null) | .output.page_context] |
    all(.[];
      .present == true and
      .protocol == "a3s.test.page-context/1" and
      .sdk_version == "0.6.2" and
      .snapshot.ui != null and
      .snapshot.ui.protocol == "a3s.test.ui-understanding/1" and
      (.snapshot.components | length) > 0
    )
  )
' "$run_result" >/dev/null

for boundary_id in \
  office-playground \
  office-editor-document \
  office-editor-spreadsheet \
  office-editor-presentation \
  office-editor-markdown \
  office-editor-pdf; do
  jq -e --arg boundary_id "$boundary_id" '
    any(
      ([.scenarios[].steps[] | select(.output.page_context != null) | .output.page_context.snapshot.components[]?.id][]);
      . == $boundary_id
    )
  ' "$run_result" >/dev/null || {
    echo "Test Kit gate did not publish the required component boundary: $boundary_id" >&2
    exit 1
  }
done

run_id="$(jq -er '.run_id' "$run_result")"
diagnostic_count="$(find "$repository_root/.a3s-test/runs/$run_id" -type f -path '*/diagnostics/*.json' | wc -l | tr -d ' ')"
if ((diagnostic_count == 0)); then
  echo "A3S Test produced no browser diagnostics for the Test Kit gate." >&2
  exit 1
fi
while IFS= read -r diagnostic; do
  jq -e '
    if .data.messages? != null then
      .success == true and .error == null and
      all(.data.messages[]?; ((.type // "info") != "error" and (.type // "info") != "warning"))
    elif .data.errors? != null then
      .success == true and .error == null and (.data.errors | length == 0)
    else false end
  ' "$diagnostic" >/dev/null || {
    echo "Browser console or page errors were captured: $diagnostic" >&2
    jq -c '{data: .data, error: .error}' "$diagnostic" >&2
    exit 1
  }
done < <(find "$repository_root/.a3s-test/runs/$run_id" -type f -path '*/diagnostics/*.json' | sort)

echo "A3S Test Kit UI gate passed: $snapshot_count valid UI snapshots across the editor surfaces."
