# Local A3S Test E2E

These deterministic browser scenarios use the separately installed
[`a3s-test`](https://github.com/A3S-Lab/Test) CLI. They are local release gates:
GitHub Actions does not install or run A3S Test. Local development may use the
A3S Browser adapter or the standalone `agent-browser` adapter.
Evidence is written under `.a3s-test/`, which is ignored by Git.

The current local release contract pins `a3s-test 1.0.0`, standalone
`agent-browser 0.26.0`, and Web protocol revision 15. The gate prefers the
Chromium executable provided by the repository's Playwright install and falls
back to the already-installed agent-browser Chromium when that cache is
unavailable; it never downloads or installs a browser. A missing or
incompatible executable fails before any suite starts.

Build the Playground and start its static preview in one terminal:

```bash
bun run playground:build
bun run playground:preview
```

Validate and run the ACL suites from another terminal:

```bash
bun run test:e2e:check
bun run test:e2e
```

## Authoring contract

These files are deterministic ACL regressions for workflows that have already
been observed. Use an A3S Test agent session first when the next action still
depends on the current page state, then promote only the smallest stable path
to `tests/e2e/`.

- Run `a3s-test check <suite> --json` before launching the browser.
- Prefer `role`, `label`, `testid`, and `placeholder` for actions that accept
  semantic targets. With the current Web adapter, visibility checks require
  `css` or a snapshot `ref`; keep those selectors tied to an accessibility or
  route contract, and use implementation CSS only for third-party or purely
  visual state.
- Follow a state-changing action with a typed `wait` or `expect`; do not add
  arbitrary sleeps. Use a snapshot only when the observed intermediate state
  is itself required before the next action, such as an animated third-party
  navigation layer.
- Capture one final-state screenshot when visual proof matters, accessibility
  only for the semantic contract under test, and console plus page errors for
  browser failures. Keep every artifact path relative.
- The pinned standalone `agent-browser 0.26.0` release gate uses Playwright for
  pixel evidence and ACL accessibility output for semantic evidence. Its CDP
  screenshot command can exceed the bounded command deadline with both the
  existing Playwright Chromium and its version-matched local Chrome, so release
  ACL suites must not depend on that command until the adapter contract fixes
  the timeout.
- Let A3S Test own and clean up each browser surface. Do not terminate shared
  browser processes by name.

## Test Kit UI gate

The Playground publishes the dev-only `@a3s-lab/testkit@0.6.2` page-context
bridge. It is the source of UI-understanding evidence for the five editor
surfaces: component boundaries and ownership, computed style/layout, motion,
responsive conditions, and state diffs after an interaction. The bridge is
disabled in the production bundle, so this gate intentionally runs against the
development Playground on port `3000`.

Run the bridge tests and the standalone A3S Test Kit gate with:

```bash
bun run playground:testkit
bun run test:e2e:testkit:check
bun run test:e2e:testkit
```

`playground:testkit` is a Playwright functional/visual suite with desktop and
compact projects. It asserts the live handshake, SDK and protocol versions,
component facts, bounded UI snapshots, toolbar/grid/view interactions, dialog
geometry, focus, and responsive layout. `test:e2e:testkit` runs
`tests/e2e/office-testkit-ui.acl` with the repository-compatible `a3s-test`
binary and standalone `agent-browser 0.26.0`, then verifies every captured page
context, UI-understanding protocol, component boundary, screenshot,
accessibility tree, and browser diagnostics. The ACL covers the home, Writer,
Spreadsheet, Presentation, Markdown, and PDF surfaces; the PDF stability path
uses the A3S Test accessibility/screenshot evidence while the Playwright bridge
test additionally proves organizer state changes.

The gate prewarms the lazy editor chunks before ACL execution. This keeps the
first Rsbuild development compilation (and its possible HMR reload) outside a
scenario's interaction and readiness assertions.

The regular static Web gate skips this dev-only suite because the `4175`
production preview aliases Test Kit to its no-op implementation. The
`test:e2e` script chains the static release suites and this dedicated live UI
gate. Set `A3S_TEST_INCLUDE_TESTKIT=true` only when a compatible dev bridge is
available to the static runner.

The documentation navigation and collaboration homepage have focused gates:

```bash
bun run test:e2e:docs:check
bun run test:e2e:docs
bun run test:e2e:collaboration-playground:check
bun run test:e2e:collaboration-playground
```

The documentation gate also opens the bilingual **What's new / 更新日志**
timeline on desktop and at 390 px, verifies localized deep links, and proves
that a frozen documentation version cannot expose later releases.

The product home also has a focused visual and interaction gate. It waits for
the shipped Document, Spreadsheet, and Presentation components, exercises the
stage tabs on desktop and phone viewports, and records the live-editor
accessibility tree plus console/page-error diagnostics:

```bash
bun run test:e2e:homepage:check
bun run test:e2e:homepage
```

The release-facing documentation and Playground template grid share one focused
discoverability gate:

```bash
bun run test:e2e:playground-templates:check
bun run test:e2e:playground-templates
```

It opens the Writer table of contents, native document index, character shading,
proofing languages, Spreadsheet data validation, bounded Spreadsheet structured
references, and Presentation animations from the Playground template grid;
follows the documentation-home links to the full Writer, Spreadsheet, and PDF
references; repeats the public discovery check at 390 px; and captures
accessibility plus empty console/page-error diagnostics. A hidden `?e2e=` fixture
or a deep documentation paragraph alone cannot satisfy this gate.

Presentation entrance and exit animations have a separate focused local gate:

```bash
bun run test:e2e:presentation-animation:check
bun run test:e2e:presentation-animation
bun run playground:visual:presentation-animation
```

The ACL opens the public **进入与退出动画** entry, selects the fly-in/fly-out
object, verifies class, effect, trigger, direction, timing, ordering, and preview
controls, and consumes both click cues before the only slide can advance. The
complementary Playwright workflow switches classes, authors exit direction and
duration, and verifies entrance and final-hidden exit cue states in both desktop
and compact projects. The release gate remains local-only and reuses the
repository's existing Playwright Chromium.

Presentation Chinese IME publication has a focused local gate:

```bash
bun run test:e2e:presentation-chinese-ime:check
bun run test:e2e:presentation-chinese-ime
```

The component regression dispatches a real composition lifecycle and proves
phonetic pre-edit text is never published to the controlled host. The A3S Test
workflow then enters Chinese title text in the built Playground, requires the
rendered value to match exactly once, reopens the presentation to prove
persistence, and captures accessibility plus empty console/page-error evidence.

The PDF page organizer has its own focused local gate:

```bash
bun run test:e2e:pdf-page-organization:check
bun run test:e2e:pdf-page-organization
```

It inserts, undoes, redoes, rotates, deletes, reorders, merges, extracts, and
checks split availability through the public PDF surface. The complementary
Playwright suite performs real drag and drop, independently parses downloaded
PDF bytes, and verifies save/reopen order, page counts, and rotations.

Run the focused suites against the built static preview through
`scripts/run-a3s-test-web-gate.sh` before a release. The local gate performs
admission first, generates the ignored deterministic upload fixtures, owns the
preview lifecycle, and fails when captured console or page-error diagnostics
are non-empty. Local environments with the A3S Browser integration can run the
gate without the standalone adapter:

```bash
A3S_TEST_BROWSER_DRIVER=a3s bash scripts/run-a3s-test-web-gate.sh
```

The source-level Spreadsheet 1904 date-system path has a focused import,
dynamic-filter, controlled-reopen, accessibility, and diagnostics gate:

```bash
bun run test:e2e:spreadsheet-1904-date-system:check
bun run test:e2e:spreadsheet-1904-date-system
```

The first-open focus regression also has a focused gate:

```bash
bun run test:e2e:initial-focus:check
bun run test:e2e:initial-focus
```

The collaboration participant roster has a focused desktop/phone gate:

```bash
bun run test:e2e:collaboration-participants:check
bun run test:e2e:collaboration-participants
```

The Document suggestion workflow has a focused desktop/phone gate:

```bash
bun run test:e2e:collaboration-document-suggestions:check
bun run test:e2e:collaboration-document-suggestions
```

The Writer character-formatting revision workflow has a focused gate:

```bash
bun run test:e2e:writer-formatting-revision:check
bun run test:e2e:writer-formatting-revision
```

The Writer paragraph-formatting revision workflow has a focused gate:

```bash
bun run test:e2e:writer-paragraph-formatting-revision:check
bun run test:e2e:writer-paragraph-formatting-revision
```

The Writer native character-spacing workflow has a focused local gate:

```bash
bun run test:e2e:writer-character-spacing:check
bun run test:e2e:writer-character-spacing
```

The Writer native kerning workflow has a focused local gate:

```bash
bun run test:e2e:writer-kerning:check
bun run test:e2e:writer-kerning
```

The Writer native East Asian emphasis-mark workflow has a focused local gate:

```bash
bun run test:e2e:writer-emphasis:check
bun run test:e2e:writer-emphasis
```

It covers `Cmd/Ctrl+D`, the semantic emphasis selector, live circle and explicit
none previews, native DOM projection, exact editor-focus restoration, one-step
Undo, accessibility evidence, and empty console/page-error diagnostics. The
suite uses A3S Test only through the local release gate and is not referenced by
GitHub Actions.

The Writer native hidden-text workflow has a focused local gate:

```bash
bun run test:e2e:writer-hidden-text:check
bun run test:e2e:writer-hidden-text
```

It covers `Cmd/Ctrl+D` authoring, the mixed-safe hidden-text checkbox, dotted
preview, hidden-by-default editing, the View-tab reveal command, the standard
`Cmd/Ctrl+Shift+H` toggle, one-step Undo, accessibility evidence, and empty
console/page-error diagnostics. Preview and PDF output stay suppressed. The
suite uses the latest pinned local A3S Test release gate and is not referenced
by GitHub Actions.

The Writer native outline, shadow, emboss, and imprint workflow has a focused
local gate:

```bash
bun run test:e2e:writer-legacy-text-effects:check
bun run test:e2e:writer-legacy-text-effects
```

It opens the public **文字效果** Playground template, proves the valid
outline-shadow pair plus standalone emboss and imprint samples, selects real
document text, and opens the shared `Cmd/Ctrl+D` dialog. Enabling emboss must
clear outline, shadow, and imprint in the same transaction; Apply must retain
Worker/WASM layout eligibility, restore editor focus, and Undo must restore the
original pair in one step. The suite captures accessibility and empty
console/page-error diagnostics through pinned local A3S Test 1.0.0,
agent-browser 0.26.0, and protocol revision 15. GitHub Actions and Pages do not
install or invoke A3S Test; desktop and 390 px Playwright own the visual proof.

The Writer native character-border workflow has a focused local gate:

```bash
bun run test:e2e:writer-run-border:check
bun run test:e2e:writer-run-border
```

It opens the public **字符边框** Playground template, selects a real native
line-border sample, and uses the shared `Cmd/Ctrl+D` dialog to author a double
wave border with exact eighth-point width, integer spacing, shadow, and frame
semantics. Apply restores editor focus, one Undo restores the original border,
and the suite captures accessibility plus empty console/page-error evidence.
A3S Test remains a pinned local release gate and is not installed or invoked by
GitHub Actions or Pages.

The Writer native proofing-language workflow has a focused local gate:

```bash
bun run test:e2e:writer-proofing-languages:check
bun run test:e2e:writer-proofing-languages
```

It opens the public **校对语言** Playground template, verifies independent
Latin, East Asian, and bidi language slots plus explicit proofing exclusion,
authors a mixed-safe language and proofing-state patch from Review, restores
editor focus, and proves one-step Undo. A second scenario keeps the same dialog
semantically usable at 390 px. The suite captures accessibility and empty
console/page-error diagnostics through the pinned local A3S Test gate; GitHub
Actions and Pages use the corresponding desktop and 390 px Playwright flow and
never install or invoke A3S Test.

The Writer native table-of-contents workflow has a focused local gate:

```bash
bun run test:e2e:writer-table-of-contents:check
bun run test:e2e:writer-table-of-contents
```

It opens the public **可更新目录** template, verifies typed heading targets,
customizes hyperlinks and the native leader, applies with exact editor-focus
restoration, refreshes current page numbers, and proves one-step Undo. A second
scenario keeps the complete dialog usable at 390 px. The suite captures
accessibility and empty console/page-error diagnostics through pinned local
A3S Test 1.0.0, agent-browser 0.26.0, and protocol revision 15; GitHub Actions
and Pages never install or invoke A3S Test.

The Writer native document-index workflow has a focused local gate:

```bash
bun run test:e2e:writer-document-index:check
bun run test:e2e:writer-document-index
```

It opens the public **原生索引** template, navigates from a generated page
number to its stable marker, edits page-number emphasis, customizes columns,
layout, and leaders, refreshes a deliberately stale generated cache, and proves
that customization and refresh each have one-step Undo without discarding a
separate marker edit. A second scenario keeps both dialogs usable at 390 px.
The suite captures accessibility and empty console/page-error diagnostics
through pinned local A3S Test 1.0.0, agent-browser 0.26.0, and protocol revision
15; GitHub Actions and Pages never install or invoke A3S Test.

The Writer document compare/combine workflow has a focused local gate:

```bash
bun run test:e2e:writer-document-comparison:check
bun run test:e2e:writer-document-comparison
```

It opens the public **文档比较** template, imports a revised HTML file through
Review, and proves that insertion, deletion, paragraph-formatting, and inserted
block revisions appear in the normal Changes pane. The desktop scenario makes
one real review decision; the 390 px scenario verifies the same comparison
dialog remains usable. Both capture accessibility and empty console/page-error
diagnostics through pinned local A3S Test 1.0.0, agent-browser 0.26.0, and
protocol revision 15. GitHub Actions and Pages never install or invoke A3S
Test.

The Writer native character-scale workflow has a focused local gate:

```bash
bun run test:e2e:writer-character-scale:check
bun run test:e2e:writer-character-scale
```

The Writer native character-position workflow has a focused local gate:

```bash
bun run test:e2e:writer-character-position:check
bun run test:e2e:writer-character-position
```

The native Spreadsheet cell lifecycle has a focused gate:

```bash
bun run test:e2e:collaboration-spreadsheet-cells:check
bun run test:e2e:collaboration-spreadsheet-cells
```

The maximum-size sparse Spreadsheet workflow has a focused gate:

```bash
bun run test:e2e:spreadsheet-maximum-sparse:check
bun run test:e2e:spreadsheet-maximum-sparse
```

The Spreadsheet four-direction fill workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-cell-fill:check
bun run test:e2e:spreadsheet-cell-fill
```

It proves relative-formula and style propagation, one-step Undo and Redo,
sparse blank-row safety, and ownership of `Cmd/Ctrl+D` and `Cmd/Ctrl+R` without
turning the latter into a browser refresh. This suite remains local-only and is
not referenced by a GitHub Actions workflow.

The Traditional Office Find and Select / Go To workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-go-to:check
bun run test:e2e:spreadsheet-go-to
```

It covers the keyboard-accessible Find and Select menu, `Ctrl+G`, `F5`,
direct and named cross-sheet ranges, validation, exact focus restoration, and
the view-only contract that leaves content and Undo unchanged. It is local-only
and is not referenced by a GitHub Actions workflow.

The Traditional Office Spreadsheet hyperlink workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-hyperlink:check
bun run test:e2e:spreadsheet-hyperlink
```

It covers the Insert ribbon command, grid-scoped `Cmd/Ctrl+K`, safe Web,
cell-range, and worksheet targets, hidden-sheet and unsafe-URL validation,
Insert/Edit/Remove dialog semantics, exact ribbon and grid focus restoration,
and one-step Undo. The suite runs only through the local A3S Test release gate
and is not referenced by GitHub Actions.

The Traditional Office Spreadsheet data-validation workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-data-validation:check
bun run test:e2e:spreadsheet-data-validation
```

It first opens the public Playground **新建 → 数据验证** template and verifies
that its list, input-message, blank, dropdown, and error-alert metadata is
actually reachable by users. It then covers multiple captured ranges, invalid
and valid list sources, complete input/error settings, compact Apply and Remove,
exact focus restoration, and one-step Undo. The suite runs only through the
local A3S Test release gate and is not referenced by GitHub Actions.

The Traditional Office Spreadsheet Paste Special workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-paste-special:check
bun run test:e2e:spreadsheet-paste-special
```

It covers the rich same-editor clipboard, the Paste split menu,
`Cmd/Ctrl+Alt+V`, content and arithmetic choices, formula translation, exact
grid-focus restoration, and one-step Undo. The suite runs only through the
local A3S Test release gate and is not referenced by GitHub Actions.

The Traditional Office built-in Cell Styles gallery has a focused local gate:

```bash
bun run test:e2e:spreadsheet-cell-style:check
bun run test:e2e:spreadsheet-cell-style
```

It covers the grouped desktop/compact gallery, preview semantics, keyboard
focus, native styling of populated and blank cells, one-step Undo, and clean
browser diagnostics. The equivalent Playwright regression remains the browser
coverage used by GitHub Actions.

The Traditional Office Spreadsheet font-size and border shortcuts have a focused local gate:

```bash
bun run test:e2e:spreadsheet-font-size-border-shortcuts:check
bun run test:e2e:spreadsheet-font-size-border-shortcuts
```

It covers both Grow Font and Shrink Font aliases, visible ribbon commands,
Outside Borders and Clear Borders, shortcut metadata in the border menu,
one-step Undo, accessibility, and clean browser diagnostics. The suite runs
only through the local A3S Test release gate; GitHub Actions keeps the
equivalent Rstest and desktop/compact Playwright coverage.

The Traditional Office Spreadsheet independent diagonal-border workflow has a focused local
gate:

```bash
bun run test:e2e:spreadsheet-diagonal-borders:check
bun run test:e2e:spreadsheet-diagonal-borders
```

It covers Ribbon diagonal-down, Format Cells diagonal-up, crossed borders,
persistence, Undo, compact Ribbon layout, screenshot and accessibility
evidence, and empty console and page-error diagnostics. The suite uses the
pinned local A3S Test 1.0.0,
standalone agent-browser 0.26.0, Web protocol revision 15, and the existing
Playwright Chromium. GitHub Actions and Pages run the equivalent Rstest and
desktop/compact Playwright coverage and never install or invoke A3S Test.

The Traditional Office Spreadsheet underline styles have a focused local gate:

```bash
bun run test:e2e:spreadsheet-underline-styles:check
bun run test:e2e:spreadsheet-underline-styles
```

It covers the five-state split menu, initial menu focus, double-accounting
application, active-style-aware `Cmd/Ctrl+U`, one-step Undo, compact Ribbon
layout, accessibility, screenshot evidence, and clean browser diagnostics.
The suite runs only through the local A3S Test release gate; GitHub Actions
keeps the equivalent Rstest and desktop/compact Playwright coverage.

The Traditional Office Spreadsheet text-orientation and row/column visibility workflow has a
focused local gate:

```bash
bun run test:e2e:spreadsheet-orientation-visibility:check
bun run test:e2e:spreadsheet-orientation-visibility
```

It covers all six orientation choices, menu radio state, exact grid-focus and
one-step Undo behavior, the four hide/unhide shortcut declarations, runtime
focus ownership, accessibility, and clean browser diagnostics. Playwright owns
the desktop/compact screenshots and Canvas pixel comparisons because the pinned
standalone adapter's CDP screenshot command is not a reliable bounded gate.
The suite remains local-only and is not referenced by GitHub Actions.

The Traditional Office Spreadsheet direct-color resets and font-emphasis aliases have a
focused local gate:

```bash
bun run test:e2e:spreadsheet-font-colors-shortcuts:check
bun run test:e2e:spreadsheet-font-colors-shortcuts
```

It covers Automatic Color, No Fill, `Ctrl+2`, `Ctrl+3`, and `Ctrl+4`, exact
shortcut metadata, real pressed-state changes, final worksheet-grid focus,
both desktop and compact Ribbon densities, accessibility, and clean browser
diagnostics. The suite runs only through the local A3S Test release gate;
GitHub Actions keeps the equivalent Rstest and Playwright coverage.

The Traditional Office Spreadsheet static date and time workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-date-time:check
bun run test:e2e:spreadsheet-date-time
```

It covers the grid-scoped `Ctrl+;` date shortcut, the discoverable Date and
Time menu, `Ctrl+Shift+;` metadata, local date/time number formats, one
controlled update per insertion, one-step Undo, compact Ribbon focus,
accessibility, and clean browser diagnostics. The Playground fixture reports
revision evidence without freezing a real user's local clock. The suite runs
only through the local A3S Test 1.0.0 release gate; GitHub Actions and Pages
use equivalent Rstest and Playwright coverage and never install or invoke A3S
Test.

The Spreadsheet copy-from-above shortcuts have a focused local gate:

```bash
bun run test:e2e:spreadsheet-copy-from-above:check
bun run test:e2e:spreadsheet-copy-from-above
```

It covers exact formula copying with `Ctrl+'`, cached-value copying with
`Ctrl+Shift+'`, formula-bar synchronization, target-style preservation, one
controlled update per command, and one-step Undo. The dedicated Playground
fixture exposes revision and cell-state evidence; desktop and compact
Playwright keep the same workflow under the browser release gate. The A3S Test
1.0.0 suite remains local-only and is never installed or invoked by GitHub
Actions or Pages.

The Spreadsheet font-dialog shortcuts have a focused local gate:

```bash
bun run test:e2e:spreadsheet-font-dialog-shortcuts:check
bun run test:e2e:spreadsheet-font-dialog-shortcuts
```

It covers `Cmd/Ctrl+Shift+F` opening the Font tab on the font-family control,
`Cmd/Ctrl+Shift+P` opening the same tab on font size, catalog-backed
`aria-keyshortcuts`, exact grid-focus restoration, accessibility, screenshot
evidence, and clean browser diagnostics. Desktop and compact Playwright use
the same public Playground path. The A3S Test 1.0.0 suite remains local-only;
GitHub Actions and Pages do not install or invoke it.

Native Spreadsheet rich-text cells have a focused local gate:

```bash
bun run test:e2e:spreadsheet-rich-text:check
bun run test:e2e:spreadsheet-rich-text
```

It covers the controlled Playground fixture's three ordered native runs,
cell-wide Bold formatting through one revision, preservation of run text and
count, exact grid-focus restoration, screenshot and accessibility evidence,
and clean console and page-error diagnostics. Import/export/reopen Rstest
fixtures separately cover shared strings, inline strings, exact whitespace,
semantic colors, and fail-closed budgets. Desktop and compact Playwright use
the same public path. The A3S Test 1.0.0 suite remains local-only; GitHub
Actions and Pages do not install or invoke it.

The native Spreadsheet Table lifecycle has a focused local gate:

```bash
bun run test:e2e:spreadsheet-table:check
bun run test:e2e:spreadsheet-table
```

It covers grid-scoped `Cmd/Ctrl+T`, captured-range and header semantics, the
contextual Table Design ribbon, all three built-in style families, responsive
two-dimensional keyboard navigation, exact style-trigger focus restoration,
table-style options, Convert to Range, one-step Undo, accessibility, and clean
browser diagnostics. The suite is local-only; GitHub Actions keeps the
equivalent Playwright regression and never installs or invokes A3S Test.

The native Spreadsheet totals-row authoring workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-table-totals:check
bun run test:e2e:spreadsheet-table-totals
```

It opens the public structured-reference workbook, verifies the enabled totals
row and native function selector, changes a column to `AVERAGE`, opens a custom
formula editor, and records accessibility plus empty browser diagnostics. The
suite is local-only; GitHub Actions and Pages do not install or invoke A3S
Test.

The offline Spreadsheet multi-key Custom Sort workflow has a focused local
gate:

```bash
bun run test:e2e:spreadsheet-custom-sort:check
bun run test:e2e:spreadsheet-custom-sort
```

It captures the quarterly pipeline selection, verifies the editable header and
two ordered value keys, moves a relative formula with its row, and confirms
that one Undo restores both values and formulas. The suite also records focus,
accessibility, console, and page-error evidence. It uses the pinned local A3S
Test 1.0.0 binary and remains absent from GitHub Actions and Pages.

The offline Spreadsheet custom-list Sort workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-custom-list-sort:check
bun run test:e2e:spreadsheet-custom-list-sort
```

It rejects a duplicate authored sequence, accepts and locally saves a bounded
status order, applies that order with a secondary value key, verifies reuse in
the mounted editor and after a full-page reload, and confirms formula-safe row
movement plus one-step Undo. Accessibility, focus, console, and page-error
evidence are captured by the pinned local A3S Test 1.0.0 binary; the suite
remains absent from GitHub Actions and Pages.

The offline Spreadsheet appearance Sort workflow has a focused local gate:

```bash
bun run test:e2e:spreadsheet-appearance-sort:check
bun run test:e2e:spreadsheet-appearance-sort
```

It applies two same-column cell-color priorities plus font-color and
conditional-format icon levels to one captured range, verifies the resulting
four-level priority and moved relative formula, then confirms one Undo restores
the source rows and formula.
The pinned local A3S Test 1.0.0 binary also records invoker/grid focus,
interactive accessibility, console, and page-error evidence. The suite remains
absent from GitHub Actions and Pages.

`spreadsheet-row-sort.acl` switches the Traditional Office-compatible sort options to
left-to-right, applies stable multi-row keys to complete columns, verifies
relative formula translation and focus restoration, then undoes the single
batch mutation:

```bash
bun run test:e2e:spreadsheet-row-sort:check
bun run test:e2e:spreadsheet-row-sort
```

`spreadsheet-text-sort.acl` selects Traditional Office-compatible stroke and
case-sensitive comparison, applies a lowercase-first tie breaker, verifies a
moved relative formula and immediate post-remount keyboard focus, then undoes
the single batch mutation:

```bash
bun run test:e2e:spreadsheet-text-sort:check
bun run test:e2e:spreadsheet-text-sort
```

The Traditional Office-style Spreadsheet partial-range Sort Warning has a focused local gate:

```bash
bun run test:e2e:spreadsheet-sort-range:check
bun run test:e2e:spreadsheet-sort-range
```

It selects only the formula column in the public quarterly workbook, verifies
the default expanded current region and independently safe exact option, walks
both Custom Sort transitions, and confirms formula-safe row movement plus one
Undo. The suite records invoker focus, accessibility, console, and page-error
evidence with the pinned local A3S Test 1.0.0 binary. It remains absent from
GitHub Actions and Pages.

The Spreadsheet table and AutoFilter-owned Sort workflow has a focused local
gate:

```bash
bun run test:e2e:spreadsheet-owned-range-sort:check
bun run test:e2e:spreadsheet-owned-range-sort
```

It starts from one cell inside a native table, requires the complete owned
range, locks the structural header and top-to-bottom orientation, verifies
formula-safe row movement, and confirms one Undo restores the original row.
The complementary Playwright suite repeats table and native value-filter paths
at desktop and compact widths, including filter-row remapping after sort and
Undo. The A3S Test gate records accessibility and browser diagnostics locally
and remains absent from GitHub Actions and Pages.

The native Presentation scene-element and z-order lifecycle has a focused gate:

```bash
bun run test:e2e:collaboration-presentation-elements:check
bun run test:e2e:collaboration-presentation-elements
```

The suites cover focused Word page-color, a desktop Writer quick-access and
adaptive, collapsible-ribbon workflow covering comfortable, priority-compacted,
persistent collapsed, and temporary tab states without document layout
movement, plus executable Traditional Office formatting and review shortcuts scoped to the
document, including permission-free format copy and paste, accessibility,
console, and page-error evidence;
the complete phone Page Setup flow
with grouped header/footer formatting controls, phone paragraph-spacing and
pagination popovers with committed document semantics and exact invoker-focus
restoration, a phone list-gallery workflow covering bounded ribbon paging,
square bullets, upper-Roman numbering, a custom start value, viewport-safe
touch controls, TipTap selection and editor-focus recovery, and
Escape-to-invoker focus restoration, compact comments-drawer, a phone citation
workflow with
field-level validation, save, insertion, and close-to-invoker focus
restoration, full-text navigation, the complete phone Find/Replace workflow,
page-preview, bounded page-button virtualization for a
real 120-page DOCX with animation-frame-independent Home/End keyboard focus,
bounded heading-outline and full-text-result windows for the same real DOCX,
including sparse current/selection pins, physical spacer geometry, exact
selection of result 120, and same-frame alignment with the final document page,
selected-text AI question-composer regressions, and durable picture insertion
from a project-relative fixture after the browser-managed file input is reset,
plus a phone-width Picture Properties workflow covering centimeter size,
aspect-ratio locking, tight contour wrapping, wrap-side selection, alignment,
text distance, quarter-turn rotation, horizontal/vertical reflection,
alternative text,
retained image selection, and exact invoker restoration after apply and cancel,
plus a desktop and phone Writer text-box workflow covering isolated insertion,
contextual layout/geometry/placement/fill/outline controls, accessibility,
viewport containment, and clean browser diagnostics,
plus a desktop and phone Writer content-control workflow covering the
responsive Insert dialog, plain/rich type and appearance metadata, direct
paragraph editing, one-step Undo/Redo, accessibility, and clean browser
diagnostics. The native `w:sdt` round-trip boundary is covered by the focused
Rstest suite; unsupported bindings and structural variants remain explicitly
diagnosed rather than promoted to an editable promise,
and a
phone tracked-changes workflow with in-pane recording controls, truthful empty
states, exact decision focus, and close-to-invoker restoration, plus a real
120-comment native OOXML DOCX workflow covering bounded card and connector
mounting, Home/End reachability, document-selection and physical-page
synchronization, the 120-to-119 deletion transition, adjacent focus continuity,
and non-blocking status feedback, plus a real
120-revision DOCX workflow covering bounded mounted rows, physical spacer
geometry, Home/End reachability, keyboard acceptance, the 120-to-119 count
transition, and same-action focus continuity, plus a phone
table workflow covering touch-sized dimension controls, exact 3 × 3 insertion,
row extension, compact-ribbon deletion, and focused editing recovery, plus a
desktop table-border workflow proving that the reusable border pen does not
flatten existing edges, outside borders retain independent cell-edge data in
edit and preview, and browser diagnostics stay clean, plus a real theme-backed
DOCX table workflow proving tint/shade resolution for cell fill and four
independent border edges across edit and preview with screenshot,
accessibility, console, and page-error evidence, plus a real inherited-table-
style DOCX workflow proving based-on resolution, first-row fill, white bold
header text and double border, horizontal banding, last-row presentation, and
conditional paragraph alignment, spacing, pagination, and indent semantics
across edit and preview. The same workflow verifies truthful 62.5% centered
geometry in the unified Table Properties dialog, compact 390 px touch controls,
all four Table, Row, Column, and Cell tabs, and one atomic change covering 72%
left placement with a 0.5 cm indent, a 1.2 cm exact row, disabled row splitting
and repeated headings, a 4 cm column, centered cell content, and a 0.4 cm left
cell margin. It also proves cancellation and Escape safety, invoker-focus
restoration, exact retention of untouched imported 8 px and 16 px margin
edges, and matching preview geometry, with screenshot, accessibility, console,
and page-error evidence. A real multi-page
DOCX table workflow proves that a splittable
row continues at paragraph boundaries across three physical pages, repeated
heading rows remain visible
on both continuation pages, and the final row content and following paragraph
remain keyboard reachable, plus a
phone caption and cross-reference workflow covering dialog focus, accessible
numbering, keyboard deletion, and the live missing-reference transition;
the phone bookmark workflow covering focused name entry, a complete paired
bookmark range, live bookmark-REF insertion, current-target deletion, the
missing-reference transition, and exact editor-focus recovery;
the phone footnote workflow covering two-way reference/definition pairing,
reference-order renumbering, paired deletion, atomic undo, accessibility, and
browser diagnostics;
the phone body-field workflow covering physical PAGE and NUMPAGES results,
F9 refresh, one-action undo and redo, phone ribbon overflow, accessibility,
and browser diagnostics;
the Spreadsheet ribbon workflow with command-catalogued native strikethrough
and its Traditional Office `Cmd/Ctrl+5` shortcut, plus colored cell-border selection, focus
restoration, and undo through the local-only ACL gate; the phone
worksheet-rename workflow and
viewport-safe, touch-sized Find controls,
exact-result navigation, and grid-focus restoration, a safe-area-aware bottom
context menu with 44 px actions, bounded scrolling, screenshot and
accessibility evidence, together with modal
workbook task panes that isolate the ribbon, grid, and worksheet footer,
contain forward and reverse focus, and restore the exact ribbon invoker; a
maximum-size Spreadsheet workflow that keeps 1,048,576 rows by 16,384 columns
sparse through `Control+End` navigation and materializes exactly one far row
after a formula-bar edit, with accessibility and browser-diagnostic evidence;
Presentation focus
recovery across object cut and paste, presenter-view keyboard navigation and
phone layout, plus a phone chart-inspector workflow covering responsive modal
semantics, background isolation, forward and reverse focus containment, dirty
draft cancellation, screenshot and accessibility evidence, and exact
selected-chart focus restoration, plus phone comment review covering full-
editor modal semantics, touch-sized review controls, background isolation,
forward and reverse focus containment, dirty-comment cancellation, and exact
New Comment and View Comments invoker restoration; PDF import
from a project-relative fixture plus focus-synchronized thumbnail keyboard
navigation, plus a phone page-drawer workflow proving toolbar ownership,
selection, modal close behavior, and an unobstructed document canvas; the
phone Markdown source, visual, and split-mode workflow with
unambiguous accessible controls, a source-first full-workspace phone switch,
synchronized preview content, and touch-sized actions; and the public
Playground-to-documentation-center navigation contract, including the
documentation center at the website root, Playground in the top navigation,
Simplified Chinese as the default language, a prominent homepage and sidebar route to the
five-format real-time collaboration guide, collaboration accessibility
evidence, and language and release-version switching.
The responsive Office-shell workflow also crosses from a persistent desktop
sidebar into a 390 px workspace, proves the phone drawer stays closed without
moving search focus, and restores the prior persistent sidebar on return.
The collaboration roster workflow projects human and agent Awareness state in
edit and preview chrome, keeps the phone popover in view, and restores the
exact trigger after Escape.
The Document suggestion workflow uses two independent Yjs documents, proves
attributed insertion synchronization, keeps accept and reject controls out of
the suggester surface, records the editor's final decision, and verifies the
converged audit trail on desktop together with phone reachability. It asserts
the pending state and captures final accessibility plus clean browser
diagnostics.
The Writer character-formatting revision workflow enters through the public
Playground demo, verifies the dedicated Formatting card and native revision
mark, rejects that revision, and proves that the document text remains while
both the applied bold style and revision wrapper are removed while the separate
paragraph revision remains pending. The paragraph-formatting workflow opens the
same public fixture, verifies the Paragraph Formatting card, rejects the
paragraph revision, and proves that the original alignment, indentation,
spacing, and line height return without changing text or the character
revision. Both assert pending and rejected states and capture final
accessibility plus clean browser diagnostics.
The Writer native strikethrough workflow selects real document text, authors
double strike from the accessible split menu, proves the command catalog does
not claim `Mod+Shift+S`, verifies that key chord leaves formatting unchanged,
then checks explicit `none`, Undo restoration, native DOM metadata, a final
accessibility snapshot, and clean browser diagnostics. Visual evidence remains
covered by the focused Playwright functional test. Run it locally with
`bun run test:e2e:writer-strike`; the focused gate reuses the repository's
Playwright Chromium and is intentionally not part of GitHub Actions or Pages.
The Writer native character-spacing workflow selects real document text, opens
the accessible advanced font dialog through `Cmd/Ctrl+D`, authors exact
expanded spacing, restores editor focus, and proves one-step Undo together with
accessibility and clean browser diagnostics. Desktop and compact Playwright
cover the same public path. The A3S Test suite remains local-only and is not
referenced by GitHub Actions or Pages.
The Writer native kerning workflow uses the same `Cmd/Ctrl+D` dialog to author
an exact effective threshold, verifies native DOM metadata and focus recovery,
then clears only direct kerning and proves one-step Undo restoration. The
focused A3S Test suite captures accessibility, console, and page-error evidence
locally; desktop and 390 px Playwright capture the responsive dialog and verify
the same workflow. It is not referenced by GitHub Actions or Pages.
The Writer native character-position workflow uses the same advanced font
dialog to change exact spacing and raised position in one Apply action, proves
both native DOM attributes, focus restoration, and one-step Undo, then captures
accessibility and clean browser diagnostics. Desktop and compact Playwright
cover the live preview and responsive controls on the same public path. The
A3S Test suite remains local-only and is not referenced by GitHub Actions or
Pages.
The Writer native character-scale workflow authors an exact `1..600` integer
percentage through the same accessible dialog, combines scale, spacing, and
position in one transaction, verifies the live preview and native DOM metadata,
then proves selection focus and one-step Undo. The focused A3S Test suite also
captures accessibility, console, and page-error evidence locally; it is not
referenced by GitHub Actions or Pages.
The native Spreadsheet collaboration workflow applies the Rust CLI's real Yjs
updates to an initialized browser workbook, verifies the visible A2 value and
formula change, then covers sparse cell creation and exact deletion with
screenshot, accessibility, console, and page-error evidence.
The native Presentation collaboration workflow applies the Rust CLI's real Yjs
updates to an initialized browser deck, verifies a visible scene-element update
and creation, moves that object to the first z-order-array position through a
stable predecessor mutation, then covers exact deletion through a durable
tombstone with screenshot, accessibility, console, and page-error evidence.
The phone Word font workflow captures the grouped picker, traverses the full
list with End, applies the final font with document-focus recovery, and checks
Escape-to-combobox restoration together with clean browser diagnostics.
Each suite owns only its browser surface. Keep the preview process under the
terminal that started it and stop that process separately when testing is
complete.
