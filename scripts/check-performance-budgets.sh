#!/usr/bin/env bash
# Phase 3.4 — compare a benchmark JSON report against performance-budgets.json.
# Usage:
#   bun run performance:check-budgets -- path/to/report.json
# Exits 0 when every reported scenario metric is within budget (or absent).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
BUDGETS="$ROOT/.a3s-test/performance/performance-budgets.json"
REPORT="${1:-}"

if [[ -z "$REPORT" ]]; then
  echo "Usage: $0 <benchmark-report.json>" >&2
  exit 2
fi
if [[ ! -f "$REPORT" ]]; then
  echo "Benchmark report not found: $REPORT" >&2
  exit 1
fi
if [[ ! -f "$BUDGETS" ]]; then
  echo "Budgets file not found: $BUDGETS" >&2
  exit 1
fi

bun -e '
import { readFile } from "node:fs/promises";

const budgets = JSON.parse(await readFile(process.argv[1], "utf8"));
const report = JSON.parse(await readFile(process.argv[2], "utf8"));
const runs = Array.isArray(report) ? report : report.runs ?? report.results ?? [report];
const failures = [];

for (const run of runs) {
  const kind = run.kind ?? run.scenario ?? run.name;
  if (!kind || !budgets.scenarios[kind]) continue;
  const budget = budgets.scenarios[kind];
  const milestones = run.milestones ?? {};
  const checks = [
    ["editorVisibleMs", milestones.shellMountedMs ?? milestones.contentMountedMs ?? run.editorVisibleMs],
    ["paginationReadyMs", milestones.paginationReadyMs ?? run.paginationReadyMs],
    ["longestTaskBeforeVisibleMs", run.importLongTasks?.maximumMs ?? run.longestTaskBeforeVisibleMs],
    ["jumpToLastSlideMs", run.jump?.durationMs ?? run.jumpToLastSlideMs],
    ["viewerReadyMs", run.viewerReadyMs],
    ["firstPageBitmapMs", run.firstPageBitmapMs],
  ];
  for (const [key, value] of checks) {
    if (budget[key] == null || value == null || !Number.isFinite(value)) continue;
    if (value > budget[key]) {
      failures.push(`${kind}.${key}=${value} exceeds budget ${budget[key]}`);
    }
  }
  const p95 = run.interaction?.p95Ms ?? run.scroll?.frameInterval?.p95Ms;
  if (p95 != null && p95 > budgets.interaction.p95MainThreadTaskMs * 4) {
    // Scroll frame interval is not the same as typing task p95; allow 4× headroom
    // until interaction-specific reports are standardized.
    failures.push(`${kind}.frameIntervalP95Ms=${p95} exceeds relaxed interaction budget`);
  }
}

if (failures.length) {
  console.error("Performance budget failures:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
console.log(`Performance budgets OK for ${runs.length} run(s).`);
' "$BUDGETS" "$REPORT"
