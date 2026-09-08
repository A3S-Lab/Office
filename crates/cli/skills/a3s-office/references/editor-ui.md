# Editor UI operation contract

Use this reference when the request is about the A3S Office browser editors,
their WPS-like interaction model, or UI/UX evidence. The native file CLI and
the browser editor are different surfaces: use typed native commands for file
mutations and the local UI operator for interaction evidence. Do not replace a
UI regression with a semantic file preview.

## One operator for five surfaces

From an Office source checkout, discover the matrix first:

```bash
bun run office:ops -- plan all --json
bun run office:ops -- capabilities --json
bun run office:ops -- doctor --json
```

The matrix is the single source of truth for the focused visual specs, A3S
Test ACLs, deterministic fixtures, and the WPS reference flag:

| Surface | Focused local contract | Typical evidence |
| --- | --- | --- |
| Writer | `writer` | ribbon state (incl. Ctrl+F1 collapse), WPS shortcuts (incl. Ctrl+Enter page break, Ctrl+K hyperlink, Ctrl+F/Ctrl+H find-replace, Ctrl+Shift+D/Ctrl+Shift+W underline, Ctrl+Shift+X double strikethrough, Ctrl+= / Ctrl+Shift+= sub/superscript, Ctrl+]/Ctrl+[ font size, Ctrl+M / Ctrl+Shift+M, Shift+Alt+. / Shift+Alt+,, and Alt+Shift+→ / Alt+Shift+← indent, Alt+Shift+↑ / Alt+Shift+↓ paragraph move, Ctrl+L/E/R/J align, Ctrl+Shift+J distribute, Ctrl+Shift+L bullets, Ctrl+1/5/2 line spacing, Ctrl+Alt+1/2/3 headings, Ctrl+Shift+N Normal/正文, Shift+F3 case cycle, Ctrl+Space clear formatting, Ctrl+Shift+G word count, Ctrl+Shift+F9 unlink, Ctrl+F11 / Ctrl+Shift+F11 lock), Insert/Page Layout chrome, field settings/`MERGEFORMAT`, field-code toggle (Alt+F9 document-wide / Shift+F9 selected), imported DOCX, pagination, font/grid metrics, undo/redo, WPS drawing boundary |
| Spreadsheet | `spreadsheet` | ribbon IA (incl. Ctrl+F1 collapse), freeze panes, View workbook chrome (编辑栏 / 显示公式 Ctrl+` / 网格线 / 标题), sheet Ctrl+PageUp/PageDown navigation, sort/filter, format cells, paste special, phone task panes |
| Presentation | `presentation` | object selection, animation/task panes, cut/paste focus, slide sorter Enter→normal, View 备注 notes pane, collapsible ribbon Ctrl+F1, Ctrl+M / ⌘⇧N new slide, Ctrl+D slide/object duplicate, Delete/Backspace slide delete on thumbnails (visual + ACL), Ctrl+G / Ctrl+Shift+G group/ungroup, slideshow F5 / Shift+F5 start, B/W blank screens, digit+Enter slide jump, Home/End first/last slide, compact canvas |
| Markdown | `markdown` | source/visual/split modes, link dialogs (Ctrl+K), source bold/italic (Ctrl+B / Ctrl+I) with visual + ACL, read-only preview, focus return, collapsible ribbon Ctrl+F1 |
| PDF | `pdf` | thumbnail rail, page organization, keyboard navigation (PageUp/PageDown, Space/Shift+Space, Ctrl+Home/End), search focus (Ctrl+F), WPS zoom (Ctrl+0/1/2), compact page drawer, save/reopen |

`plan <surface> --json` is the machine-readable workflow manifest. It expands
the selected matrix row into typed fixture, ACL, A3S Test gate, desktop/compact
visual, exploratory-agent, and (for Writer) WPS COM reference commands. Keep
the manifest as the handoff between Codex/CLI and the UI evidence runner.
`doctor --json` must report `a3sTestCompatible: true` before a browser gate is
treated as runnable; stale 0.x binaries are surfaced as infrastructure failure.
Interactive `run` and `agent start` apply the same version gate before opening
a browser, while static `check` remains available for diagnosing ACL syntax
with an older binary. Windows dispatch keeps selectors and action JSON in typed
argv rather than shell tokenization. When `--cdp-port` is supplied on Windows,
the operator compiles and uses a native `.exe` CDP adapter under
`.a3s-test/office-ops/`; this avoids passing state-changing A3S Test arguments
through a `.cmd` shim. The adapter prefers a Playwright-managed Chromium build
under `%LOCALAPPDATA%\ms-playwright`, auto-launches it when the CDP port is
empty, sets viewport size through `Browser.setWindowBounds`, and waits for CSS
selectors directly over CDP so spaced attribute selectors are not split. Set
`A3S_TEST_CHROME_EXECUTABLE` only when an explicit reviewed browser binary is
required. Set `A3S_TEST_AGENT_BROWSER` only when an explicit reviewed adapter is
required.

Run one bounded surface gate while developing. `check` is the fast ACL parse;
add `--run` when a configured A3S Test browser should execute the primary
interaction suite:

```bash
bun run office:ops -- gate writer --run \
  --browser-driver standalone \
  --cdp-port 9345
bun run office:ops -- gate spreadsheet
bun run office:ops -- gate presentation
bun run office:ops -- gate markdown
bun run office:ops -- gate pdf
```

`gate` regenerates ignored fixtures and checks the selected ACLs. With `--run`
it executes those ACLs through A3S Test; unless `--no-visual` is supplied it
then runs Playwright only as a supplemental desktop/compact pixel baseline. It
writes evidence under `.a3s-test/office-ops/` and does not update committed
snapshots or wait for CI.

Run the primary A3S Test path directly against one ACL when narrowing a
failure. A loopback preview can be reused; the operator materializes an
ignored manifest copy when the preview uses a different port:

```bash
bun run office:ops -- a3s run tests/e2e/word-connector-editor.acl \
  --base-url http://127.0.0.1:4175/playground/ \
  --browser-driver standalone \
  --cdp-port 9345 --json
```

For exploratory UI work use the persistent A3S Test agent lifecycle. Keep it
bounded as `start → observe → one act → observe → finish/abort`; refs are not
reused after a state-changing action:

```bash
bun run office:ops -- a3s agent start writer \
  --url http://127.0.0.1:4175/playground/ \
  --cdp-port 9345 --json
bun run office:ops -- a3s agent observe --session <id> --interactive --json
bun run office:ops -- a3s agent act --session <id> --observation <n> \
  --action-json '{"type":"click","target":{"type":"ref","value":"@e7"}}' --json
bun run office:ops -- a3s agent finish --session <id> --status passed \
  --summary "Observed the editor and captured the final state." --json
```

For common actions, use the typed wrappers instead of manually escaping JSON.
They forward one schema-validated A3S Test action and preserve the same
observation/ref safety boundary:

```bash
bun run office:ops -- a3s agent click --session <id> \
  --observation <n> --target 'role=button|保存' --json
bun run office:ops -- a3s agent fill --session <id> \
  --target 'label=标题' --value 'A3S Office' --json
bun run office:ops -- a3s agent select --session <id> \
  --target 'css=#status' --value draft --value review --json
bun run office:ops -- a3s agent viewport --session <id> \
  --width 390 --height 844 --scale 1 --json
bun run office:ops -- a3s agent assert-visible --session <id> \
  --target 'css=.work-markdown-workspace' --json
```

The target grammar is `@e7`, `css=<selector>`, `role=<role>|<name>`,
`label=<text>`, `placeholder=<text>`, `testid=<id>`, `automation=<id>`, or
`text=<text>`. This grammar is intentionally explicit so a malformed target
fails before dispatch rather than silently selecting a different editor
control. Use `--action-json` only for an action outside the typed wrapper set.

Inspect the locked CUA Driver/MCP matrix before attempting native GUI testing:

```bash
bun run office:ops -- a3s cua certification --json
```

The locked CUA Driver 0.10.0 matrix currently marks both Windows profiles as
`unsupported` because there is no reviewed Windows application backend. The
operator therefore fails closed for Windows native-GUI claims and uses A3S Test
Web/CDP for browser-editor evidence. Do not report a CUA pass from an
unsupported profile. On a contract-tested platform, pass the approved CUA
policy and proxy options to `a3s-test run`.

When the certification matrix reports a reviewed native profile, the matching
CUA Driver/MCP lifecycle can be exercised explicitly with
`a3s cua certify --gui-policy-file <policy> --cua-proxy-executable <proxy>`.
On this Windows checkout that command must remain an explicit, fail-closed
experiment because the locked profiles are unsupported.

Run only one part when narrowing a failure:

```bash
bun run office:ops -- fixtures
bun run office:ops -- check writer
bun run office:ops -- visual writer --project compact-768
bun run office:ops -- visual visual-tests/markdown-menu.functional.spec.ts
```

The visual operator builds `playground-dist/playground` before launching the
static preview. `playground:preview` fails with an actionable message when the
bundle is absent instead of serving the documentation home at
`/playground/`. To inspect a running development bridge, use port `3000` and
the Test Kit suite separately:

```bash
bun run playground
bun run playground:testkit
bun run test:e2e:testkit:check
```

## Observe, act, prove

For every UI change, keep the loop bounded and evidence-based:

1. Open the real template or deterministic fixture for the target surface.
2. Observe a stable semantic target (`role`, accessible name, test id, or the
  documented editor boundary) before acting. Prefer the A3S Test agent or ACL
  action schema; do not add an ad-hoc switch/`if` parser around CSS branches.
3. Perform one user-level action, then observe/assert the state change. Avoid
  arbitrary sleeps and DOM implementation handles that the pagination/
  windowing engine may recycle.
4. Assert the visible result, keyboard focus, responsive geometry, undo/redo,
   accessibility, and empty console/page-error diagnostics as appropriate.
5. Capture one final screenshot and retain the run JSON/diagnostics. A
   screenshot alone does not prove semantics; an ACL alone does not prove the
   visual hierarchy.

Classify a failure before changing product code:

- **Product**: the route is loaded, the editor boundary is ready, and a
  semantic or visual assertion is wrong.
- **Test contract**: the action targets a stale label, fixture, route, or
  unsupported browser project.
- **Infrastructure**: the preview is missing, the browser executable is
  unavailable, the CDP adapter cannot attach, or the locked CUA profile is
  unsupported. Preserve the artifact and fix/classify the operator path first.

## WPS COM reference, Windows only

WPS COM is a bounded reference probe for discovering native OOXML and UI
semantics; it is not an A3S Office runtime dependency and must not be used as a
CI prerequisite. From the source checkout:

```powershell
bun run office:ops -- wps-probe --connector
```

For a named output:

```powershell
bun run office:ops -- wps-probe --connector --output .a3s-test/office-ops/wps/connector.docx
```

The checked-in `scripts/probe-wps-shapes.ps1` starts the installed WPS Writer
COM server in a hidden, bounded session, creates a rectangle and optionally a
straight connector with the bounded line/arrow reference, saves a DOCX, and
closes every COM object. Inspect the
OOXML or open the result through the A3S editor; do not infer browser fidelity
from COM properties alone. Remove only the exact temporary probe artifacts
after the comparison.

## Native CLI versus editor UI

Use `a3s-office native validate/view/query/get` to inspect and mutate saved
DOCX/XLSX/PPTX files. Use `a3s-office collab` for durable Markdown,
Document, Spreadsheet, Presentation, or PDF replica mutations and event
cursors. Use A3S Test for browser interaction and UI proof; use Playwright only
for a supplemental pixel baseline. Keep
these boundaries explicit:

- A typed native mutation must have a validate/readback check.
- A browser mutation must have a state, focus, and diagnostics assertion.
- A WPS comparison must record the WPS version and the exact fixture path.
- A compatibility fallback is allowed only after the native capability and the
  browser boundary have been checked.

Never poll GitHub Actions as part of a local edit loop. Commit and push after
the focused local evidence is green; CI remains an independent publication
signal.
