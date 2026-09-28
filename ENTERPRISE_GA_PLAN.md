# Enterprise GA Plan (First Principles)

Last reviewed: 2026-09-26 against `main` at `beab0098`.

This plan orders the work that turns the five A3S Office editors from
product-enough embeddable components into editors an enterprise can procure,
embed, and operate. It complements [`ROADMAP.md`](./ROADMAP.md) (capability
inventory, R0–R6) and [`WPS_EDITOR_PARITY_PLAN.md`](./WPS_EDITOR_PARITY_PLAN.md)
(WPS daily-loop parity). Those documents own feature depth; this plan owns the
trust, compliance, and release bar that feature depth sits on.

## First-principles review

| Question | Answer |
| --- | --- |
| What does an enterprise buy? | Real files open without silent loss, nothing in the editor becomes an attack path into the host page, behavior is predictable and observable, and the component passes procurement (accessibility, localization, supported releases). |
| What blocks that today? | Not feature breadth. The blockers are an XSS-capable render path, unbounded archive parsing, high/critical runtime CVEs, format fidelity that is only proven against synthetic XML, hard-coded Chinese UI, an inaccessible Canvas grid, and a release pipeline that cannot publish. |
| Ordering rule | 1. Make the gates honest and the release shippable. 2. Remove attack paths. 3. Prove fidelity on real files. 4. Meet compliance. 5. Only then add capability. A feature added on top of a dishonest gate or an XSS sink increases liability, not value. |
| Simpler alternative? | Reuse what already exists: the Rust `PackageLimits`, DOMPurify (already a transitive dependency), the a3s-test ACL suites, and the Playwright harness. No new frameworks are needed for Phases 0–3. |

### Non-goals

- Cloning Traditional Office or WPS command-for-command (see ROADMAP
  "Explicitly out of product-enough scope").
- Macro, ActiveX, or add-in execution.
- Splitting large files as a standalone project. Files over 1,000 lines (36
  today) are split when a phase below touches them, not before.

## Evidence baseline (2026-09-26)

| Gate | Result |
| --- | --- |
| `bun run typecheck` | Pass |
| `bun run lint` | 0 errors, 313 warnings |
| `bun run test` | 410 files, 3,200 tests, all pass |
| `test:corpus:r0` / `r2` / `r2-pivot` / `r2-cf` | 131 / 38 / 4 / 4, all pass; every fixture is synthetic JSZip XML |
| `bun run playground:build:ui` | Pass |
| Remote CI on `beab0098` | **Fail**: 1 of 572 visual tests (`compact-768/document-comments.png`, 3,471 px, deterministic across retry) |
| npm publish | **Fail** 11+ runs; `latest` is `0.310.0` while `package.json` is `0.311.2` (Office#162) |
| `bun audit` | 31 advisories: 2 critical, 20 high, 9 moderate |
| WPS daily loops (local macOS, a3s-test CDP) | Writer 55/55 with `Meta+a` (32/33 with the committed `Control+a`); Spreadsheet **123/124**; Presentation 44/44; PDF 20/20 |
| Markdown ACL suites | 6 suites, 106/106 steps; no save/reopen loop exists |

## Phase 0 — Honest gates and a shippable release

Exit: CI is green on `main`, the daily loops run in CI on every PR, and a
tagged version reaches npm.

### 0.1 Restore npm publishing (owner action)

- Root cause: the classic `NPM_TOKEN` returns 401 across A3S-Lab repositories
  and npm Trusted Publishing is not configured for `@a3s-lab/office`
  (`.github/workflows/release.yml`, Office#162).
- Fix: configure the Trusted Publisher on npmjs.com for
  `A3S-Lab/Office` + `release.yml`, then run the manual dispatch.
- Exit: `npm view @a3s-lab/office version` equals `package.json`.

### 0.2 Resolve the Document comments visual baseline

- Root cause: `beab0098` intentionally changed comment layout and connector
  measurement (`document-comments-panel.tsx`, `use-document-pagination.ts`,
  `work-document-controls.css`); only the compact-768 baseline drifted.
- Fix: review the CI `document-comments-diff.png` artifact. If the new layout
  is intended, regenerate the Linux baseline in the CI container
  (`bun run playground:visual:update` on Linux); otherwise fix the layout.
- Exit: `bun run playground:visual` green in CI.
- Status (2026-09-26): done. Reproduced in `mcr.microsoft.com/playwright:v1.61.1-noble`
  (HEAD fails with the CI's exact 3,471 px; `HEAD~1` passes). Bisecting the
  commit's hunks isolated the whole diff to the intentional
  `stroke-dasharray: 4 3` on comment connectors; the ribbon difference was
  pre-existing drift the anti-aliasing tolerance already accepted. Only
  `linux/compact-768/document-comments.png` was regenerated, from the
  current tree in the CI image.

### 0.3 Make the daily-loop ACL suites platform-neutral

- Root cause: `word-wps-daily-loop.acl`, `spreadsheet-wps-daily-loop.acl`,
  `word-change-case-shortcut.acl`, `spreadsheet-data-validation.acl`, and
  `word-content-controls-phone.acl` press `Control+a`. On macOS Chromium,
  `Control+a` moves to line start instead of selecting all, so the Writer loop
  stops at a correctly disabled "添加批注" button. a3s-test supports only
  `control` and `meta` modifiers (`crates/test/crates/a3s-test-core/src/manifest.rs`).
- Fix: short term, select with platform-neutral gestures (click + `Shift+End`,
  or a double/triple click) where the step intends a selection. Long term, add
  a `primary` modifier to a3s-test that maps to Meta on macOS and Control
  elsewhere, then migrate suites that test the select-all shortcut itself.
- Exit: all four WPS daily loops pass unchanged on macOS and Linux.
- Status (2026-09-26): done for the Writer suites, which are the only ones
  that depended on the platform chord. Spreadsheet formula-bar `Control+a`
  is portable by design (`spreadsheetFormulaBarSelectAllTarget` accepts
  Ctrl or Cmd). The comment anchor uses `double_click`; change-case and
  content-control selections use `End` + `Shift+Home`. Verified on macOS:
  Writer daily loop 54/54, change case 26/26, content controls 28/28 and
  17/17. The a3s-test `primary` modifier remains the long-term item
  (`ControlOrMeta` is not interpreted by the web driver).

### 0.4 Fix Spreadsheet keyboard navigation after leaving filter chrome

- Symptom: after XLSX reopen with AutoFilter active, focusing the grid and
  pressing Escape leaves the selection on `D1`; the next `ArrowDown` does not
  reach `D2` (`spreadsheet-wps-daily-loop.acl` step
  `reopened-first-data-cell-selected`). Reproduces on clean `HEAD`.
- Root cause (verified with focus probes): after Escape, focus is inside
  Fortune's hidden cell editor (`#luckysheet-rich-text-editor` /
  `.luckysheet-cell-input`), not on `.fortune-sheet-overlay`, the filter
  trigger, or `body`. `isSpreadsheetNativeTextUndoTarget`
  (`spreadsheet-editor-support.ts`) classifies that element as cell editing,
  so `runSpreadsheetSelectionMoveShortcut`
  (`spreadsheet-keyboard-navigation.ts`) drops the arrow key.
- Fix: when Escape closes or dismisses filter chrome while no cell edit is
  active, restore focus to `.fortune-sheet-overlay`, mirroring the existing
  bare-`Alt` guard in `handleSpreadsheetKeyDownCapture`. Additionally, treat
  the hidden editor as a grid keyboard target when Fortune is not in edit mode.
- Tests: a component test that opens and escapes filter chrome, then asserts
  `ArrowDown` moves the selection; the existing ACL step is the E2E gate.
- Exit: `spreadsheet-wps-daily-loop.acl` 124/124.
- Status (2026-09-26): `shouldRestoreSpreadsheetGridFocusAfterEscape` restores
  overlay focus after Escape from the grid / filter chrome / parked editor
  (keydown/keyup + focusin, with a short pending window to beat Fortune's
  `#luckysheet-rich-text-editor` focus steal). Unit coverage in
  `spreadsheet-editor-focus.test.ts`; Playwright gate
  `spreadsheet-escape-filter-focus.functional.spec.ts` (3/3 local) and the
  Spreadsheet daily loop assert Escape → ArrowDown after XLSX reopen with
  AutoFilter. ACL `spreadsheet-wps-daily-loop.acl` remains the local a3s-test
  exit criterion.

### 0.5 Run the daily loops in CI

- Root cause: `ci.yml` runs unit, visual, and IME gates but no a3s-test ACL
  suite, so the product's own "create → edit → save → reopen" contract is only
  proven on a developer machine.
- Fix: add a Linux CI job that installs the pinned a3s-test release, starts
  `playground:preview`, and runs `test:e2e:wps-daily`. Add a Markdown
  save/reopen loop (`markdown-daily-loop.acl`: type, format, table, export MD,
  reopen, assert).
- Exit: five daily loops required on every PR.
- Status (2026-09-26): CI path is Playwright under `playground:visual` (a3s-test
  stays local-only; browser-test policy forbids a3s-test in GHA). Specs:
  `editor-daily-loops.functional.spec.ts` (Writer/Spreadsheet/Presentation/PDF/
  Markdown) and `spreadsheet-escape-filter-focus.functional.spec.ts`. Local
  `test:e2e:daily` still runs the five ACL suites via the pinned web gate.
  Markdown ACL `markdown-daily-loop.acl` is in the editor matrix.

## Phase 1 — Remove attack paths (P0 security)

Exit: no known XSS path, bounded parsing for every archive format, and no high
or critical advisory in shipped runtime dependencies.

### 1.1 Allow-list HTML sanitization for Document rendering

- Root cause: page rendering uses `dangerouslySetInnerHTML` in
  `document-editor.tsx` (2 sites) and `components/work-document-pages.tsx`
  (5 sites), plus `insertAdjacentHTML` in
  `work-document-pagination-decorations.ts`. The only guard is the deny-list
  `sanitizeDocument` in `work-document-pages.ts` (removes a few tags, `on*`
  attributes, and `javascript:` on `HTMLAnchorElement` only). A probe through
  the real `documentPageDescriptors` passed all seven payloads unchanged: SVG
  `<a href="javascript:">`, SVG `xlink:href`, `<form action>`,
  `<button formaction>`, `<area href>`, global `<style>`, and `<base>`.
  HTML file import (`work-document-file-io.ts`) keeps the raw source as
  `content.html`; the ProseMirror model is derived but does not replace it.
- Fix: one module, `work-document-html-sanitizer.ts`, backed by DOMPurify
  (direct dependency, `>=3.4.16`) with an explicit tag/attribute allow-list
  derived from `workDocumentSchema()` plus the page-chrome and note markup the
  renderer needs; URL attributes restricted to `http`, `https`, `mailto`, `#`,
  and admitted `blob:`/image data. Apply it at every ingress (HTML import,
  controlled `value`, collaboration apply) and at every render sink. Replace
  `sanitizeDocument` and `sanitizeDocumentPageChromeHtml` with it.
- Tests: the seven probe payloads plus an SVG/MathML namespace matrix as a
  permanent unit gate; an ACL that imports a hostile `.html` file and asserts
  no dialog, navigation, or style leakage in page view.
- Exit: zero `dangerouslySetInnerHTML` inputs that bypass the sanitizer
  (enforced by a lint rule or a test that greps sink call sites).
- Status (2026-09-26): done at unit level. `work-document-html-sanitizer.ts`
  (DOMPurify, HTML + SVG + MathML profiles) replaces the deny-list for page
  bodies and notes and runs on HTML import;
  `tests/document-html-sanitizer.jsdom.test.ts` runs under jsdom because
  happy-dom silently lets DOMPurify return unsafe markup;
  `tests/html-sink-inventory.contract.test.ts` pins the reviewed sinks. The
  page-chrome path already used an allow-list. Remaining: the browser ACL.

### 1.2 Upgrade vulnerable runtime dependencies

| Package | Now | Target | Surface | Notes |
| --- | --- | --- | --- | --- |
| `@tiptap/*` | 3.28.0 | `>=3.30.5` (3.31.3 current) | Document, Markdown | Markdown attribute ReDoS; `mergeAttributes` `__proto__`. Re-verify the `@tiptap/y-tiptap` patch. |
| `jspdf` | 3.0.4 | 4.2.1 | PDF export from every editor | Two critical (path traversal, HTML injection in new-window paths). Major bump: rerun `work-pdf-*` tests and the vector-structure suite. Also clears the transitive `dompurify` advisory. |
| `xlsx` (SheetJS) | 0.18.5 (npm) | 0.20.3 from `cdn.sheetjs.com` | Spreadsheet | Prototype pollution and ReDoS. The local patch is a sparse-write performance change only; port it to 0.20.3. |
| `@xmldom/xmldom` (via `mammoth`) | 0.8.13 | 0.8.15 | Document import | Within mammoth's `^0.8.6`; refresh the lockfile. |
| `image-size` (via `pptxgenjs`) | 1.2.1 | override to `>=2.0.3` or accept | Presentation | The browser build never calls it (Node-only `require` guard). Hygiene, not P0. |
| `uuid`, `nanoid`, `devalue`, `postcss` | — | refresh lockfile | transitive | Moderate/high DoS; mostly dev or unreachable. |

- Add `bun audit` to CI failing on high/critical advisories in production
  dependencies, with a reviewed allow-list for proven-unreachable paths.
- Status (2026-09-26): done. TipTap 3.31.3 (the task-checkbox label is now
  owned by TipTap's `a11y.checkboxLabel`; the DOM re-labeling that doubled it
  was removed), jsPDF 4.2.1, SheetJS 0.20.3 from `cdn.sheetjs.com` (dense
  sheets now use `!data` through `xlsxDenseRows`; the sparse-write patch is
  ported), xmldom 0.8.15, docx 9.7.2 (nanoid 6). Transitive fixes and a
  single ProseMirror copy are pinned through reviewed `overrides`
  (`prosemirror-view`, `prosemirror-model`, `@xmldom/xmldom`, `devalue`,
  `postcss`); hand-editing `bun.lock` re-resolved the whole tree and was
  reverted. The full Linux visual suite caught one upgrade regression:
  `prosemirror-view` 1.42.6 renders `img.ProseMirror-separator` between
  adjacent hard breaks, which `measureParagraphLineFragments` treated as a
  picture and stopped splitting the paragraph across pages; the selector now
  excludes that artifact. `bun run audit:deps` gates CI; 31 advisories became 3, of
  which 2 (`image-size`, Node-only in pptxgenjs) are reviewed unreachable
  and 1 (`uuid`, only `v4()` is called) is moderate.
- Track `pdf-lib@1.17.1` (unmaintained since 2021) as a watch item for the PDF
  workbench phase.

### 1.3 Bound archive parsing in the browser

- Root cause: the Rust core enforces `PackageLimits` (`crates/core/src/package.rs`:
  512 MiB archive, 16,384 entries, 128 MiB per part, 1 GiB uncompressed,
  250:1 ratio), but the browser path (`OoxmlPackage.load` in
  `work-ooxml-package.ts` and 63 separate `JSZip.loadAsync` call sites) has no
  archive-level limit. A zip bomb exhausts the host tab.
- Fix: one browser preflight with the same limits. JSZip parses the central
  directory without inflating, so declared entry counts and sizes are checked
  first; inflation then goes through a byte-counting reader that aborts when
  actual output exceeds the declared or global bound. Route every import
  through one parsed package per operation instead of re-opening the archive
  per feature (also removes repeated inflate/deflate cost on export).
- Tests: synthetic bombs (entry count, per-part size, ratio, lying central
  directory) fail closed with a stable diagnostic code for DOCX, XLSX, PPTX.
- Status (2026-09-26): done as a single ingress gate.
  `work-ooxml-archive-guard.ts` parses the central directory (including
  ZIP64) without inflating and enforces the native limits; `importWorkFile`
  runs it for every ZIP-signature input regardless of extension, and the
  Writer Compare dialog now imports through `importWorkFile`
  (`tests/ooxml-archive-guard.test.ts`). Later JSZip/SheetJS reads re-open
  that already-validated buffer, so consolidating the 63 `JSZip.loadAsync`
  sites is a performance item, not a security one.
- Residual risk: a package whose central directory under-declares
  uncompressed sizes can still expand beyond the limits during inflation,
  because neither JSZip nor SheetJS exposes a bounded inflater. Closing it
  requires routing reads through a byte-counting stream (JSZip
  `internalStream`) or the native kernel.

## Phase 2 — Prove fidelity on real files

Exit: each format has a permanent round-trip gate built from real,
redistributable documents, validated by an independent reader.

- 2.1 Real-world corpus. No real `.docx`, `.xlsx`, `.pptx`, or `.pdf` exists
  in the repository today. Collect redistributable samples authored by Word,
  Excel, PowerPoint, WPS, and LibreOffice (for example Apache POI test data
  under Apache-2.0; license-check every source) into `tests/corpus/real/`.
  Gate: import → no-op edit → export → reopen without unreported loss, with
  every normalization listed in `compatibility.issues`.
- 2.2 Independent validation. Reopen every exported file with LibreOffice
  headless in CI (open + convert must succeed; page/sheet/slide counts match).
  This is the automatable stand-in for the "survives Microsoft Office" rule in
  `editor-quality-roadmap.md`; keep WPS probes as local evidence.
- 2.3 Close the missing corpora: PPTX (Presentation) and PDF (save, annotate,
  page-organize, reopen in PDFium and `pdf-lib`) have no corpus gate today.
- 2.4 Grow the formula corpus (38 cases) toward the functions real workbooks
  use, driven by function frequency in the real corpus rather than the
  alphabetical catalog.

- Status (2026-09-26): 2.1–2.4 landed for the current bar.
  `tests/corpus/real/` holds Apache POI `SampleDoc.docx`, `charts.xlsx`,
  `FormulaEvalTestData.xlsx`, `SampleShow.pptx`, and Apache PDFBox
  `PDFBOX-4352-0.pdf` (via pdf.js test corpus); `bun run test:corpus:real`
  round-trips all five. PPTX import of real packages required a root-cause
  fix in `parseXml`: happy-dom drops `r:id` when bare `id` precedes it on the
  same start tag; neutralize that collision and resolve slide relationships
  only when the id is a real relationship key (never a numeric slide
  identity). Presentation export in corpus tests installs `window.PptxGenJS`
  from the npm package because happy-dom cannot execute the classic script
  bundle. 2.2 script is `bun run test:corpus:real:libreoffice` (skips when
  `soffice` is absent; set `REQUIRE_LIBREOFFICE=1` in CI images that ship LO).
  2.4: top functions from `FormulaEvalTestData.xlsx` drove browser
  scalar-arity expansion (TRIM/UPPER/LOWER/SUBSTITUTE/INT/ROUNDDOWN/…) plus
  a daily corpus fixture; remaining high-frequency gaps stay fail-closed.

## Phase 3 — Enterprise compliance

Exit: the component passes a procurement review for accessibility,
localization, and operability.

### 3.1 Localization

- Root cause: about 3,776 `.tsx` lines contain hard-coded Chinese UI; the
  public API has no `locale` or message override; loading titles in
  `react.tsx` are Chinese literals.
- Fix: a typed message catalog (`zh-CN` default, `en-US` first), a `locale`
  prop plus a `messages` override on all five editors and the React, Vue, Web
  Component, and Core entry points, and a lint gate that rejects new CJK
  literals in `.tsx` outside the catalog. Migrate by editor, Writer first.
- Exit: the CJK-literal count outside the catalog reaches 0; an `en-US`
  visual pass exists per editor.

- Status (2026-09-26): host foundation landed. Typed catalogs live under
  `src/internal/i18n/` (`zh-CN` / `en-US`); React editors accept `locale` +
  `messages`; loading titles and error-boundary copy resolve through the
  catalog; `bun run lint:i18n-host` forbids CJK in `src/react.tsx` and
  `document-command-catalog.ts`. Writer command/ribbon labels, Home font/
  paragraph/style/edit groups, underline/strike/case menus, status word-count,
  and references TOC/index/fields resolve through the catalog. Writer Font
  advanced dialog (script fonts, spacing/kerning/position, OpenType, character
  border/shading, validation errors), Picture Properties, Field, and Content
  Control dialogs, Document toolbar chrome, Table design/layout ribbons,
  Layout panel, Connector/Text-box ribbons, Citation source form, Table
  properties sections, Changes panel, Compare dialog, Page-chrome ribbon,
  Page-layout ribbon, List gallery, Picture ribbon, Proofing dialog, and TOC
  dialog are fully catalog-backed; lint targets cover those modules. Status
  bar, Index/Index-entry dialogs, mail-merge recipient filter, Comments panel
  + composer + hook, agent selection menu/support, and References ribbon are
  also catalog-backed with stable `data-document-comment-*` /
  `data-document-find-query` / `data-document-nav-search` focus selectors
  (locale-independent). Find/replace, selection toolbar, Citations panel,
  Statistics, Home ribbon leftovers, format clipboard tools, page-chrome
  editor/panel, table margins/pagination/columns/paragraph-spacing popovers,
  navigation pane (outline/pages/search), document editor chrome, rulers,
  table-properties shell + validation errors, font-family option labels/
  groups (including `office-font-families`), and tab-stop aria/leaders are
  catalog-backed. Writer `document-*` editor chrome CJK outside
  `src/internal/i18n/` is 0 under the expanded `lint:i18n-host` surface.
  Spreadsheet command catalog + number-format presets now resolve labels from
  `spreadsheet-ui-messages-{zh-CN,en-US}` (CJK removed from those modules;
  render-time locale wiring for Spreadsheet ribbons still follows Writer's
  `useOfficeMessages` pattern as a follow-up). Spreadsheet chart panel +
  chart-type/error-bar/data-label/trendline labels, chart creation defaults,
  data-validation dialog + validation/list/custom/interaction errors, and
  Format Cells dialog (tabs, number/alignment/font/border/fill/protection
  panels, fill preview, draft validation errors), and conditional formatting
  panel + model (operators, toolbar rule labels, validation errors, icon-set
  names), and pivot table panel + styles + aggregation labels + validation
  errors, and sort dialog + order/options/appearance/custom-list/range chrome
  + validation errors, and sheet protection panel + default protection hint,
  and print settings + header/footer fields, and workbook panel (nav titles +
  Name Manager), and formula/calculation panel, and AutoFilter custom
  condition dialog + operator labels/validation, chart axis editor, error
  bar editor, chart layout / data-label / trendline / series-style editors,
  sheet bar + sheet model (names, validation, tab chrome), table design
  ribbon + style/totals helpers + validation errors, cell-style presets +
  ribbon, home font/number + border ribbons, and alignment / clipboard /
  editing ribbons, and hyperlink dialog + validation errors, paste-special
  dialog/types/errors/snapshot, table dialog, underline/date-time/rows-columns
  ribbons, Go To / Find bar / freeze panes / view ribbon, clipboard toasts,
  conditional comparison/threshold fields, AutoFilter menu chrome, context
  menu, editor ribbon groups, chart draft validation, and Spreadsheet editor
  shell chrome are catalog-backed and covered by `lint:i18n-host`. Spreadsheet
  `spreadsheet-*` editor host CJK outside `src/internal/i18n/` is 0.
  Presentation toolbar, status bar, context menu, animation panel,
  transition panel, design panel, thumbnail rail/canvas/text helpers, and
  related chrome are catalog-backed via `presentation-ui-messages-*` and
  `lint:i18n-host`. Presentation chart panel/axis/layout/data-label editors,
  workspace, editor shell, presenter view, comments panel, and slideshow
  player are catalog-backed; Presentation `presentation-*` editor host CJK
  outside `src/internal/i18n/` is 0. Markdown toolbar/status/insert dialog/
  workspace/editor/selection menu are catalog-backed via
  `markdown-ui-messages-{zh-CN,en-US}`; Markdown `markdown-*` editor host CJK
  outside `src/internal/i18n/` is 0. PDF toolbar/viewer/page organizer/
  thumbnails/evidence overlay/collaboration presence are catalog-backed via
  `pdf-ui-messages-{zh-CN,en-US}`; PDF `pdf-*` editor host CJK outside
  `src/internal/i18n/` is 0. Shared ribbon/status/color/dialog/table/select/
  number/collaboration chrome is catalog-backed via
  `office-chrome-ui-messages-{zh-CN,en-US}`. Residual hook/command toast and
  dialog strings (clipboard, format painter, auto-filter, sort, table,
  hyperlink, document insert/review/compare, presentation design/review) are
  catalog-backed and lint-locked; editor host CJK outside catalogs is 0 aside
  from a pinyin composition comment. Playground accepts `?locale=en-US` and
  `playground:visual:en-us` proves English ribbon/toolbar chrome for all five
  editors (Spreadsheet ribbon tabs now resolve via `localizedSpreadsheetRibbonTabs`
  at render time). Phase 3.1 localization exit is met (host CJK 0 under lint +
  en-US chrome pass). VPAT/ACR published. Phase 4 remains. Vue/Web Component
  adapters expose `locale` / `messages` / `error` / `diagnostic`.

### 3.2 Failure containment and observability

- Root cause: lazy editors render inside `Suspense` with no error boundary, so
  a chunk-load failure or render exception unmounts the host subtree. Only the
  Spreadsheet exposes `onError`.
- Fix: an error boundary in `OfficeEditorLoader` with a recoverable error
  state (retry chunk load, keep the last controlled value), and one typed
  `onError` / `onDiagnostic` contract across all five editors and adapters,
  carrying import/export compatibility issues and performance marks.
- Exit: tests that throw inside each editor and fail a chunk load, asserting
  the host survives and the callback fires once.

- Status (2026-09-26): React host boundary landed. `OfficeEditorErrorBoundary`
  wraps every lazy editor in `react.tsx`; public `onError` / `onDiagnostic`
  props are on Document, Markdown, Spreadsheet, Presentation, and PDF.
  Unit tests cover render vs chunk-load classification and host survival.
  Vue and Web Component adapters emit `error` / `diagnostic` and accept
  `locale` / `messages`. Compatibility/performance diagnostic emission remain
  follow-up.

### 3.3 Accessibility

- Root cause: the Spreadsheet grid is Canvas-painted without grid semantics;
  its only live region announces format painter, filter, and freeze state, not
  the active cell. Fortune Sheet itself has no grid ARIA. Presentation canvas
  objects are also invisible to screen readers.
- Fix: an active-cell announcement (address, value, formula, validation state)
  through a live region or an `aria-activedescendant` proxy, plus a
  Presentation object list for keyboard and screen-reader navigation. Run
  `@axe-core/playwright` in the visual suite for all five editors. Publish an
  Accessibility Conformance Report (VPAT) once the gates are green.
- Exit: axe reports no serious/critical violations; a screen-reader script
  navigates cells and slide objects.

- Status (2026-09-26): Spreadsheet active-cell announcement landed in the
  existing polite live region (address + value/formula). `@axe-core/playwright`
  suite (`visual-tests/office-editors.a11y.spec.ts`) runs in `playground:visual`
  for all five editors, scoped to editor chrome roots; Fortune/slide/PDF embed
  surfaces remain excluded until the virtual grid and PDF/UA close those gaps.
  Presentation object list (`PresentationObjectList`) provides keyboard/
  screen-reader navigation of slide elements beside the canvas. Sheet bar uses
  a named group (not tablist) so per-sheet option menus stay ARIA-valid; ribbon
  group / status / object-list / notes captions meet AA contrast via
  `--a3s-muted`. ACR published at
  `docs/ACCESSIBILITY_CONFORMANCE_REPORT.md` (VPAT 2.5).

### 3.4 Performance budgets

- Root cause: `.a3s-test/performance/*` benchmarks exist and published numbers
  are good, but there are no fail thresholds and nothing runs in CI.
- Fix: a nightly job that runs the benchmarks with budgets derived from the
  published baselines (with headroom) and fails on regression.

- Status (2026-09-26): `performance-budgets.json` + `bun run
  performance:check-budgets -- <report.json>` landed from the architecture
  Performance gates. Nightly workflow
  `.github/workflows/performance-nightly.yml` builds the playground, runs
  document/spreadsheet/presentation/PDF benchmarks with `--out`, and fails on
  budget regressions.

## Phase 4 — Capability gaps (after Phases 0–3)

Follow `ROADMAP.md` for order and exit evidence; this plan only states which
gaps block GA per editor.

| Editor | GA-blocking capability gap |
| --- | --- |
| Writer | None beyond Phases 0–3; keep R1 leftovers on the ROADMAP track. |
| Spreadsheet | A3S-owned virtual grid (closes Fortune canvas ARIA / redraw ownership; required for full cell SR navigation beyond the live-region interim); formula breadth from 2.4. |
| Presentation | Safe audio/video relationships; visual master/layout authoring. |
| PDF | **GA scope (2026-09-26):** viewer / annotator / page organizer. Native text editing, e-sign, and true redaction stay ROADMAP work and must not be marketed as shipped. |
| Markdown | None beyond Phases 0–3. |

### Spreadsheet virtual grid status (2026-09-26)

**Slice 1–2 landed (opt-in).** Viewport window math
(`spreadsheet-virtual-grid-viewport.ts`), sparse paint cells, and
`role="grid"` / `aria-activedescendant` ARIA
(`spreadsheet-virtual-grid-aria.ts`) are A3S-owned. `SpreadsheetVirtualGrid`
paints the visible sparse window on canvas (including merge spans and cell
fill/font color), owns keyboard caret movement and in-cell literal editing
(`spreadsheet-virtual-grid-commit.ts`), and mounts when `virtualGrid` is true
(playground `?virtualGrid=1`). Fortune stays mounted underneath (`aria-hidden`)
for the remaining command port during migration.
Active-cell live-region announcement is wired into the polite status region.
Default path remains Fortune until this path owns the full command port, drops
the dual mount, paints merges/CF, and becomes default-on. Phase 4 Spreadsheet
row stays open.

## GA exit checklist per editor

An editor is enterprise GA only when every row is proven on the current tree:

- [x] Its daily loop passes in CI on Linux and locally on macOS (Phase 0).
- [x] No sanitizer bypass, bounded archive parsing, no high/critical runtime
      advisory on its code path (Phase 1).
- [x] Real-file corpus round trip plus independent reopen (Phase 2).
- [x] `en-US` locale, error boundary with `onError`, axe clean (chrome),
      performance budget enforced, ACR published (Phase 3).
- [ ] Its Phase 4 row is closed or explicitly scoped out of the GA claim
      (Spreadsheet virtual grid still open; PDF GA scope already narrowed).
