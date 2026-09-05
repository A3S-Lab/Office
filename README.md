<p align="center">
  <img src="./assets/readme/hero.svg" width="100%" alt="A3S Office brings five format-native editors into a host-owned product boundary">
</p>

<p align="center">
  <strong>Open-source browser editors and deterministic native automation for real Office files.</strong>
</p>

<p align="center">
  Document · Spreadsheet · Presentation · Markdown · PDF
</p>

<p align="center">
  <a href="https://a3s-lab.github.io/Office/playground/"><strong>Open the Playground</strong></a>
  ·
  <a href="https://a3s-lab.github.io/Office/docs/">Read the documentation</a>
  ·
  <a href="#quick-start">Embed an editor</a>
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI status" src="https://img.shields.io/github/actions/workflow/status/A3S-Lab/Office/ci.yml?branch=main&amp;style=flat-square&amp;label=CI"></a>
  <a href="https://www.npmjs.com/package/@a3s-lab/office"><img alt="npm version" src="https://img.shields.io/npm/v/@a3s-lab/office?style=flat-square&amp;color=1456f0"></a>
  <a href="#project-status"><img alt="Project status: pre-1.0" src="https://img.shields.io/badge/status-pre--1.0-7457c8?style=flat-square"></a>
  <a href="./LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/license-MIT-16845b?style=flat-square"></a>
</p>

<p align="center">
  <a href="#proof-not-promises">Product</a> ·
  <a href="#the-core-design">Design</a> ·
  <a href="#quick-start">Quick start</a> ·
  <a href="#collaboration-without-a-bundled-cloud">Collaboration</a> ·
  <a href="#automation-without-ui-scraping">Automation</a> ·
  <a href="#capabilities-and-boundaries">Boundaries</a> ·
  <a href="./ROADMAP.md">Roadmap</a> ·
  <a href="./CONTRIBUTING.md">Contributing</a>
</p>

---

A3S Office is an open-source Office engine for teams building editing into
their own products. It provides complete browser surfaces for documents,
spreadsheets, presentations, Markdown, and PDFs, plus a separate Rust
automation plane for files and coding agents.

The product boundary is deliberate: A3S Office owns format-aware editing,
layout, import, export, and typed mutations. Your application keeps control of
content, persistence, identity, authorization, collaboration transport, and AI
providers. Core editing and file workflows require no A3S backend or bundled
cloud service.

## Proof, not promises

These are committed visual-regression baselines from the real
[Playground](https://a3s-lab.github.io/Office/playground/), not concept art.

<p align="center">
  <a href="./visual-tests/__snapshots__/linux/desktop-1280/document.png">
    <img src="./visual-tests/__snapshots__/linux/desktop-1280/document.png" alt="A3S Office Document editor with an Office-style ribbon and paginated project brief" width="100%">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png" alt="A3S Office Spreadsheet editor showing a quarterly execution plan">
      </a>
      <br><sub><strong>Spreadsheet</strong> — formulas, tables, formatting, sort, filter, and print workflows</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png" alt="A3S Office Presentation editor with slide thumbnails and a slide canvas">
      </a>
      <br><sub><strong>Presentation</strong> — structured slides, objects, animations, and presenter flows</sub>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png" alt="A3S Office Markdown editor in synchronized source and preview mode">
      </a>
      <br><sub><strong>Markdown</strong> — GFM source, visual editing, and synchronized preview</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png" alt="A3S Office PDF editor with search, annotation, page organization, save, and download controls">
      </a>
      <br><sub><strong>PDF</strong> — PDFium rendering, forms, annotations, page organization, and save</sub>
    </td>
  </tr>
</table>

## Why A3S Office exists

Embedding Office work is not one problem. A dependable product needs four
things at the same time:

1. **A complete interaction surface** — ribbons, dialogs, panes, shortcuts,
   responsive layouts, accessibility, and predictable focus.
2. **Format-native behavior** — DOCX, XLSX, PPTX, Markdown, and PDF cannot be
   flattened into one lowest-common-denominator model without losing meaning.
3. **A clear ownership boundary** — the editor should not dictate where files
   live, who users are, how permissions work, or which AI provider is allowed.
4. **Deterministic automation** — agents and backend jobs need typed,
   conflict-aware mutations instead of UI scraping.

A3S Office keeps those concerns separate while exposing them through one
package and one set of bounded contracts.

## Five surfaces, one product boundary

- **Document** (`DOCX`, `HTML`, `TXT`) — structured authoring, live
  pagination, tables, equations, references, review, and PDF output.
- **Spreadsheet** (`XLSX`, `XLS`, `ODS`, `CSV`) — sparse worksheets,
  formulas, tables, formatting, sort/filter, validation, pivots, charts, and
  print workflows.
- **Presentation** (`PPTX`) — slides, typed scene objects, masters/layouts,
  transitions, bounded entrance/exit animations, notes, slideshow, and
  presenter view.
- **Markdown** (`MD`) — GFM source, visual mode, split preview, direct round
  trips, and native automation.
- **PDF** — PDFium rendering, search, forms, annotations, history, save, and
  page organization.

The surface is requested lazily. Large or expensive work is isolated behind
cancellable Workers, Rust WebAssembly, viewport-bounded rendering, and an
explicit PDFium runtime. React, Vue 3, Web Components, and the framework-neutral
Core API use the same controlled content models.

## The core design

<p align="center">
  <a href="./assets/readme/architecture.svg">
    <img src="./assets/readme/architecture.svg" width="100%" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

The two execution planes solve different jobs:

- **Browser editing** provides complete interactive surfaces through React,
  Vue 3, Web Components, and the Core API. Workers, Rust/WASM, and PDFium keep
  layout, calculation, parsing, and rendering bounded.
- **Native automation** provides deterministic file reads, validation,
  mutation, batching, CLI commands, a standard MCP server, and an Office Skill
  without launching desktop Office.

Both planes are controlled by the host:

| A3S Office owns | Your product owns |
| --- | --- |
| Format-native models and commands | Content persistence and version history |
| Import, export, layout, and rendering | Identity, authorization, and policy |
| Editor UI and responsive interaction | Application shell and navigation |
| Typed collaboration and agent ports | Rooms, transport, model providers, and AI lifecycle |

Neither plane requires an A3S backend. Read the
[browser architecture](./docs/latest/en/browser-editor-architecture.md) and
[native engine design](./docs/latest/en/native-office-engine.md) for the exact
boundaries.

## Quick start

### Try the complete product

The fastest first success is the
[live Playground](https://a3s-lab.github.io/Office/playground/). It exposes
normal document templates, recent capabilities, file import, and every editor
without a local install.

To run the same Playground locally, use Node.js 20+, Bun 1.3+, and Rust 1.85+:

```bash
git clone https://github.com/A3S-Lab/Office.git
cd Office
bun install --frozen-lockfile
bun run playground
```

### Embed a controlled React editor

```bash
bun add @a3s-lab/office react react-dom
```

Import the stylesheet once, give the editor an explicit-height host, and retain
the complete value emitted by `onChange`:

```tsx
import { useState } from 'react';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

const initialContent: DocumentContent = {
  type: 'document',
  html: '<h1>Project brief</h1><p>Start editing here.</p>',
  pageSize: 'a4',
  pageColor: '#ffffff',
};

export function ProjectBrief() {
  const [content, setContent] = useState(initialContent);

  return (
    <main style={{ height: '100dvh', minHeight: 0 }}>
      <DocumentEditor
        content={content}
        onChange={setContent}
        theme="system"
      />
    </main>
  );
}
```

The editor owns the editing transaction. Your application decides when, where,
and how the controlled value is stored.

### Choose an entry point

| Entry point | Use it for |
| --- | --- |
| `@a3s-lab/office/react` | Lazy React editor components and preload helpers |
| `@a3s-lab/office/vue` | Vue 3 adapters with `v-model:content` |
| `@a3s-lab/office/web-component` | Framework-agnostic Custom Elements |
| `@a3s-lab/office/core` | Models, templates, import/export, file workflows, and Yjs bindings |
| `@a3s-lab/office/styles.css` | Shared editor and interaction-system styles |

Copyable integrations live in the
[component documentation](https://a3s-lab.github.io/Office/docs/components/).

## Files stay files

The Core API can import and export files without mounting an editor:

```ts
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '@a3s-lab/office/core';

const shell = createArtifact('blank-document');
const artifact = await importOfficeFile(file, {
  artifactId: shell.id,
  onProgress: ({ stage, progress }) => {
    console.info(stage, Math.round(progress * 100));
  },
});

const output = await createArtifactBlob(artifact);
```

Imported DOCX artifacts are source-backed. Safe, unedited package parts and
stable native identities are preserved under bounded rules; unsupported or
unsafe structures produce compatibility diagnostics or fail explicitly rather
than being attached to the wrong content. Structured Document values can also
cross process boundaries through the versioned snapshot codec.

See the
[Document](./docs/latest/en/components/document.mdx),
[Spreadsheet](./docs/latest/en/components/spreadsheet.mdx),
[Presentation](./docs/latest/en/components/presentation.mdx), and
[PDF](./docs/latest/en/components/pdf.mdx) references for file-specific
contracts.

## Collaboration without a bundled cloud

Every editor exposes the same transport-neutral collaboration boundary through
Yjs/Yrs content and Awareness. Browser users, native replicas, CLI sessions,
MCP clients, and A3S Code can participate in the same host-owned document.

A3S Office provides format-specific bindings, local undo, validated presence,
remote selections or locations, comments, suggestions, and conflict-local
typed mutations. The host provides rooms, authentication, authorization,
delivery, offline buffering, persistence, and the `Y.Doc`.

This separation makes collaboration optional infrastructure rather than a
mandatory account or storage service. Start with the
[collaboration guide](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
or the runnable
[A3S Boot example](./examples/collaboration-server/).

## Automation without UI scraping

The Rust CLI, standard MCP server, typed Rust API, and packaged Office Skill
share the same bounded file contracts:

```bash
# Validate and inspect a file.
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- view report.docx outline --json

# Apply an exact guarded mutation.
cargo run -p a3s-office-cli -- set report.docx /body \
  --find Draft --replace Final --json

# Expose the same contracts over standard MCP.
cargo run -p a3s-office-cli -- mcp
```

Native collaboration replicas can exchange standard Yjs updates and state
vectors, then apply typed Document, Markdown, Spreadsheet, Presentation, and
PDF mutations without interpreting Office's private CRDT schema.

Read the [automation guide](https://a3s-lab.github.io/Office/docs/automation/)
and [CLI reference](./docs/latest/en/cli-reference.md).

## Current release

Version `0.53.1` closes the next WPS drawing boundary without pretending that
legacy VML connectors are editable text boxes:

- A Windows WPS 12.0 COM probe records `Shapes.AddConnector` as a VML
  `v:shape` (`o:spt="32"`, `#_x0000_t32`). A dedicated `docx.connectors`
  diagnostic reports that evidence and keeps endpoint, routing, arrowhead, and
  floating-anchor semantics on the compatibility path.
- A deterministic connector DOCX fixture, focused Rstest, and local A3S Test
  ACL verify that import keeps the Writer editor usable, does not create a
  `data-document-text-box` node, and captures accessibility plus clean browser
  diagnostics. The live browser run remains dependent on the pinned local
  adapter; static ACL validation is part of the release contract.

Version `0.53.0` extends Writer's bounded native text-box workflow with
WPS-referenced shape geometry and closes the import-to-editor model loop:

- The contextual Text Box ribbon exposes rectangle, rounded rectangle, ellipse,
  diamond, and triangle presets. One typed shape value drives the live page,
  preview, PDF capture, native DOCX export, and one-step Undo/Redo.
- Isolated WPS `mc:AlternateContent` text-bearing shapes preserve their native
  geometry, placement, fill, outline, padding, vertical anchor, and drawing
  identity. Mixed paragraphs, connectors, arbitrary shapes, malformed bodies,
  and unsupported DrawingML branches stay diagnosed compatibility boundaries.
- A structured-model parsing regression was fixed so imported shape attributes
  are not replaced by default text-box values. Desktop and 390/768 px browser
  flows cover all five shape controls, WPS import, contextual discovery,
  accessibility, viewport containment, and clean diagnostics; the local A3S
  Test ACL is checked against the pinned release contract.

Version `0.52.0` adds bounded, WPS-referenced whole-paragraph revision
fidelity to Writer:

- Exact text-only paragraph insertions and deletions now import as one atomic
  review item, preserve author and date, accept or reject the complete block,
  and round-trip through native `w:pPr/w:rPr` paragraph-mark records.
- The importer accepts the separate paragraph-mark and body revision IDs used
  by WPS when author, timestamp, and structure agree. Isolated paragraph-break
  merges/splits, mixed or relationship-bound content, malformed metadata, and
  over-limit input remain explicit fail-closed diagnostics.
- The WPS Office 12.1.0.22215 COM probe confirmed `Bravo\r` deletion and
  `Delta\r` insertion ranges for the reference document. Desktop and 390 px
  review flows add real DOCX import, atomic decisions, keyboard and focus
  checks, accessibility evidence, clean diagnostics, and touch-sized compact
  actions.

Version `0.51.0` extends Writer Compare with bounded, WPS-referenced
cross-paragraph text moves:

- A unique lexical range can move between aligned simple text paragraphs or
  headings in the same section and become one paired `move` review item.
  Separators travel with the range; accepting, rejecting, undoing, or reopening
  the DOCX preserves exact source and revised text.
- The Review ribbon and Changes pane keep one localized **移动** card with
  source/destination marks, destination navigation, author guidance, and
  move-aware toast counts. Duplicate or mark-mismatched candidates, section
  boundaries, rich or relationship-bound content, tables, and over-limit input
  remain fail-closed ordinary revisions or diagnostics.
- The WPS 12.0 COM/UIA probe observed `CompareDocuments` returning ordinary
  delete/insert records for the tested reorder. A3S therefore documents this
  paired inference as a bounded local enhancement, with no claim that WPS
  exposed a native move type through that API.

Version `0.50.0` extends Writer Compare with a bounded, WPS-referenced text
move workflow:

- Deterministic lexical ranges that move within one simple paragraph or heading
  become one paired `move` review item. Separators travel with the range, and
  accepting, rejecting, undoing, or reopening the DOCX preserves exact source
  and revised text.
- The Review ribbon and Changes pane expose one localized **移动** card with
  source/destination marks, destination navigation, updated author guidance,
  and move-aware toast counts. Ambiguous duplicates, rich or relationship-bound
  content, cross-paragraph moves, and over-limit inputs remain fail-closed
  boundaries.
- The WPS 12.0 COM/UIA probe observed `CompareDocuments` returning ordinary
  delete/insert records for the tested reorder. A3S therefore documents its
  paired inference as a bounded local enhancement, with no claim that WPS
  exposed a native move type through that API.

Version `0.49.0` adds bounded native Writer move revisions and hardens the
local PDF workflow:

- Strict and transitional DOCX `w:moveFrom`/`w:moveTo` text pairs import as one
  atomic `move` review item. Accept/reject resolves both source and destination
  sides together, the immutable collaboration audit records the single final
  decision, and export/reopen restores native move records with no private
  marker leakage.
- Rich, range-marker, relationship-bound, malformed, and unpaired moves remain
  visible as diagnosed compatibility boundaries. The review pane labels moves
  distinctly and navigates to the destination range without selecting the text
  between the two sides.
- The Playground now serves the pinned PDFium WebAssembly asset from a stable
  origin-relative path, keeping direct PDF imports ready in desktop and compact
  layouts.

Version `0.48.1` fixes the Writer floating selection toolbar's split emphasis
controls without changing the document model or public command contract:

- Underline and strikethrough controls now reset browser-native button
  appearance and use the same compact themed states as the surrounding
  actions, including one-pixel grouping dividers and keyboard focus treatment.
- The fix is covered by a browser visual contract in both desktop and compact
  viewports; formatting commands, accessibility names, and split menus remain
  unchanged.

Version `0.48.0` adds bounded native Writer content controls without
introducing a second document model or a remote service:

- The Insert ribbon and responsive dialog author inline plain-text or rich-text
  controls with aliases, program tags, multiline text, border/tag/hidden
  appearance, and an optional theme color. Each accepted change is one typed
  update and one Undo step, and the control exposes an accessible textbox name.
- Content and shell locks are enforced at the transaction boundary, including
  direct editor transactions. Locked content cannot be changed accidentally;
  typed unlock/delete commands remain explicit and auditable.
- Direct paragraph DOCX `w:sdt` controls round-trip through strict and
  transitional WordprocessingML with collision-free native IDs, rich run
  formatting, lock/multiline metadata, and Word 2012 appearance/color. Active
  bindings, placeholders, repeating regions, form controls, nested or
  relationship-bound structures remain safe editable text with compatibility
  diagnostics instead of an inexact editable promise.

Version `0.47.0` adds a bounded common-field workflow for Writer without
introducing a second document model or a remote service:

- The Insert ribbon exposes live word and character counts. PAGE, NUMPAGES,
  SECTION, SECTIONPAGES, DATE, and TIME continue to refresh from the measured
  page model, while generated field results stay out of document statistics.
- The Cross-reference dialog can insert a bookmark-backed `PAGEREF` with a
  stable target identity. Supported native DOCX switches round-trip as fields;
  unsupported, malformed, nested, or missing-target instructions remain cached
  text with explicit compatibility diagnostics.

Version `0.46.0` adds native, bounded Writer text boxes without introducing a
second document model or a remote service:

- The Insert ribbon adds an editable **Text Box** block. Its contextual ribbon
  controls inline/floating placement, millimeter geometry, relative offsets,
  fill, outline, padding, and vertical alignment with one controlled update and
  one Undo record per intent.
- Isolated WordprocessingML text boxes round-trip through native DOCX geometry,
  body properties, and fill/outline state and share the same live, preview, and
  PDF projection. Mixed or malformed drawings remain an explicit diagnostic
  boundary instead of being silently flattened.

Version `0.45.0` adds bounded Writer picture transforms without introducing a
second document model or a remote service:

- The contextual **Picture** ribbon now exposes quarter-turn rotation plus
  horizontal and vertical reflection with accessible Lucide icons. **Picture
  Properties** keeps the same controls in a responsive dialog, and each
  confirmed change is one undoable, controlled update.
- The transform is projected consistently into editing, preview, PDF capture,
  and the native DOCX `a:xfrm` fields (`rot`, `flipH`, and `flipV`). Imported
  90-degree/reflection values reopen as editable state, while arbitrary-angle
  or malformed values produce a compatibility diagnostic and normalize safely.

Version `0.44.0` adds dependent local dropdowns to Spreadsheet Data Validation
without introducing a remote service:

- **Data → Data Tools → Data Validation** accepts a bounded
  `=INDIRECT(...)` source composed of quoted text and single-cell references.
  Relative drivers are re-evaluated for every target cell, so a Region column
  can drive a Regional owner list through workbook-local named ranges or
  one-row/one-column areas.
- Empty drivers show an empty list. External books, whole-row/column ranges,
  missing sheets, uncached formulas, and malformed expressions fail closed;
  formulas stay within 255 Unicode characters and local range reads stay
  bounded.
- The controlled model keeps the authored formula compact while the grid gets
  a bounded runtime projection. Native XLSX import/export/reopen preserves the
  list formula and names, and the public **新建 → 数据验证** template includes
  the Region → Regional owner example.

Version `0.43.0` makes Spreadsheet formula conditional formatting a local,
editable workflow instead of a file-only preservation path:

- **Home → Conditional Formatting → Custom formula** authors rules with
  relative/absolute references, finite cross-sheet references, independent
  text/fill colors, ordered precedence, and Stop-if-true behavior. Imported
  XLSX expression rules reopen in the same editor.
- The shared synchronous evaluator reads only cached workbook values, caps
  formulas at 255 Unicode characters and each decision at 1,024 cells, scans
  bounded blank ranges without densifying the workbook, and fails closed for
  external, whole-row/column, missing-sheet, or uncached-formula references.
- Native XLSX differential styles, `sqref`, priorities, and formulas round-trip
  through export/reopen. The public **新建 → 公式条件格式** template shows a
  blocking rule and a cross-sheet threshold without any remote service.

Version `0.42.0` extends Spreadsheet data validation with a local custom-formula
workflow that remains bounded, inspectable, and native-file compatible:

- Author a formula with an optional `=` prefix in the same accessible dialog;
  relative references are anchored per selected range and the proposed value
  is evaluated before it enters the controlled workbook.
- Common local functions, cell/range references, and sheet-qualified references
  are evaluated synchronously with a 255-character formula and 1,024-cell read
  budget. External, whole-row/column, missing-sheet, and uncached-formula
  references fail closed instead of silently accepting an unsafe edit.
- Custom rules round-trip through XLSX and are demonstrated in the public
  **新建 → 数据验证** template, alongside the existing list, date, and alert
  branches. Paste and object-level batch writes remain separate preflight
  boundaries.

Version `0.41.0` aligns Spreadsheet data-validation alerts with the local,
testable interaction model used by Traditional Office:

- Stop blocks an invalid edit with an accessible notice. Warning and
  Information ask whether the invalid value should be kept, using the authored
  title and message plus the current input so the decision is explicit.
- Keeping a Warning or Information value commits it once through the controlled
  workbook API; cancelling restores the original value. Selection, focus,
  Undo, and collaboration remain on the same bounded edit path, and native
  XLSX `errorStyle` metadata remains lossless on reopen.

Version `0.40.0` extends Writer's local, testable review workflow with atomic
ordered-list numbering revisions:

- Tracked numbering-style and starting-number changes now appear as one
  Numbering review card per ordered-list range. Accept keeps the current list;
  reject restores the exact bounded baseline; Undo, Yjs/Yrs collaboration, and
  immutable decisions share the same intent identity. Common single-level
  decimal, letter, and Roman `w:numberingChange` records round-trip through
  DOCX, while malformed or structurally ambiguous records fail closed.

- A slide object can now own one entrance and one exit animation. The shared
  object-centric Animation tab authors appear/disappear, fade in/out, fly in/out,
  and zoom in/out effects with one ordered trigger, timing, direction, preview,
  collaboration, clipboard, and Undo model. Browser playback and native PPTX
  timing-tree import/export retain both classes, while malformed or overlapping
  sequences fail closed with diagnostics.

- The documentation now includes a first-class bilingual, version-aware
  [What's new](https://a3s-lab.github.io/Office/docs/changelog.html) timeline.
  It presents editor scope, user outcome, compatibility evidence, and detailed
  references while frozen documentation hides releases that did not yet exist.

- Writer's shared advanced Font dialog now edits the complete bounded Office
  2010 OpenType typography set: 16 ligature combinations, numeral form and
  spacing, style sets 1-20, and contextual alternates. Mixed selections,
  tracked formatting, Undo, body/page-chrome stories, exact DOCX reopen, and
  malformed-input diagnostics share one model with structured equations.
- Spreadsheet XLSX import/export now retains the workbook's native 1900 or
  1904 date system, exact typed date serials, dynamic date filters, and
  epoch-correct current-date authoring across controlled and collaborative
  reopen flows.
- Spreadsheet Custom Sort now recognizes native table and AutoFilter-owned
  ranges, keeps headers and totals structurally safe, and reapplies active
  filter visibility after one formula-safe, undoable sort.
- Spreadsheet Custom Sort now includes a responsive, keyboard-accessible
  preference manager for creating, editing, deleting, and reordering bounded
  user sequences while keeping the seven built-in sequences read-only.
- Document, visual Markdown, and Presentation text composition is safe for
  Chinese and other IMEs: pre-edit text remains local, controlled replacement
  waits for settlement, and only the committed value reaches the host.
- Spreadsheet sort and AutoFilter workflows run locally, including bounded
  multi-key sorting, native wildcard filters, Top/Bottom filters, custom lists,
  and Simplified Chinese collation choices.
- Structured-reference calculation, calculated-column fill, and native totals
  rows share bounded Rust/WASM and JavaScript paths.
- The shared File menu now has explicit action icons, readable labels, bounded
  keyboard navigation, and a distinct destructive-action treatment.
- PDF page organization performs insert, delete, rotate, reorder, extract,
  merge, and split in a dedicated Worker with Blob-level Undo/Redo.

Browse the product-focused [What's new](https://a3s-lab.github.io/Office/docs/changelog.html)
page or the exhaustive engineering [changelog](./CHANGELOG.md). The Playground
exposes these workflows from the release-labelled template grid.

## Capabilities and boundaries

A3S Office favors explicit fidelity boundaries over silent approximation.
`Supported` means a path has deterministic behavioral or native round-trip
evidence. `Partial` means a useful path exists with a documented boundary.

### Document

**Strong paths:** structured authoring, pagination, tables, references, review,
large plain-document windows, and source-aware DOCX round trips.

**Boundary:** long-tail DrawingML, fields, exact layout parity, and searchable
vector PDF remain partial.

### Spreadsheet

**Strong paths:** sparse editing, dependency-aware formulas, tables,
sort/filter, validation, charts, pivots, and XLSX round trips.

**Boundary:** formula breadth, external data, macros, add-ins, and specialist
analysis remain incomplete.

### Presentation

**Strong paths:** scene editing, masters/layouts, transitions, bounded entrance
animations, presenter workflows, and PPTX round trips.

**Boundary:** broader shapes, effects, media, animation families, and exact
print/video output remain partial or unsupported.

### Markdown

**Strong paths:** GFM source and visual editing, synchronized preview,
collaboration, direct round trips, and native automation.

**Boundary:** Markdown remains format-native rather than acting as a conversion
layer for every rich-document construct.

### PDF

**Strong paths:** PDFium rendering, search, forms, annotations, save, and
bounded page organization.

**Boundary:** existing content-stream editing, OCR, signatures, optimization,
and trustworthy redaction need explicit providers or future engines.

The complete capability inventory, priorities, and exit evidence live in the
[capability roadmap](./ROADMAP.md). Collaboration delivery has its own
[roadmap](./COLLABORATION_ROADMAP.md).

## Project status

A3S Office is pre-1.0. Public models and component props are usable, but a
minor release may still include breaking model changes. Required migrations
are documented in the [changelog](./CHANGELOG.md).

The goal is predictable browser editing and honest file preservation, not
pixel parity with every desktop-suite feature. Unsupported semantics remain
visible compatibility boundaries. The
[editor quality roadmap](./docs/latest/en/editor-quality-roadmap.md) records
the evidence required to move each boundary.

## Development

Run the standard checks from the repository root:

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

Focused local A3S Test gates cover first-open focus and IME composition:

```bash
bun run test:e2e:initial-focus:check
bun run test:e2e:initial-focus
bun run test:e2e:presentation-chinese-ime:check
bun run test:e2e:presentation-chinese-ime
```

The controlled Document IME lifecycle also runs against Playwright's pinned
WebKit engine in CI:

```bash
bun run playground:ime:webkit
```

GitHub Actions does not require A3S Test. See
[CONTRIBUTING.md](./CONTRIBUTING.md),
[the E2E guide](./tests/e2e/README.md), and
[the visual-test guide](./visual-tests/README.md) for the complete validation
matrix.

## Documentation

- [Live Playground](https://a3s-lab.github.io/Office/playground/)
- [Documentation center](https://a3s-lab.github.io/Office/docs/)
- [What's new](https://a3s-lab.github.io/Office/docs/changelog.html)
- [React, Vue, Web Component, and Core API](https://a3s-lab.github.io/Office/docs/components/)
- [Real-time collaboration](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
- [CLI, MCP, and Office Skill](https://a3s-lab.github.io/Office/docs/automation/)
- [Browser editor architecture](./docs/latest/en/browser-editor-architecture.md)
- [Native Office engine](./docs/latest/en/native-office-engine.md)
- [Editor quality roadmap](./docs/latest/en/editor-quality-roadmap.md)
- [Engineering changelog](./CHANGELOG.md)

## Community and security

Contributions are welcome. Read the
[contribution guide](./CONTRIBUTING.md) and
[Code of Conduct](./CODE_OF_CONDUCT.md) before opening a change. Report
suspected vulnerabilities through the private process in
[SECURITY.md](./SECURITY.md), not through a public issue.

## License

A3S Office is available under the [MIT License](./LICENSE). Bundled PDFium and
other third-party assets carry additional notices in
[THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md).
