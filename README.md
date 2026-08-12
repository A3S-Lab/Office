<p align="center">
  <img src="assets/readme/hero.svg" width="1200" alt="A3S Office — five browser editors behind one typed boundary">
</p>

<p align="center">
  <strong>AI-native Office surfaces for the browser, backed by deterministic native automation.</strong>
</p>

<p align="center">
  Edit documents, Markdown, spreadsheets, presentations, and PDFs inside your product.<br>
  Keep persistence, identity, collaboration, authorization, and AI in your host application.
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI status" src="https://img.shields.io/github/actions/workflow/status/A3S-Lab/Office/ci.yml?branch=main&amp;style=flat-square&amp;label=CI"></a>
  <a href="https://a3s-lab.github.io/Office/"><img alt="Open the live Playground" src="https://img.shields.io/badge/Live_Playground-open-2f6fed?style=flat-square"></a>
  <a href="#project-status"><img alt="Project status: pre-1.0" src="https://img.shields.io/badge/status-pre--1.0-7a5bd6?style=flat-square"></a>
  <a href="LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/license-MIT-159469?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://a3s-lab.github.io/Office/">Playground</a> ·
  <a href="https://a3s-lab.github.io/Office/docs/">Documentation</a> ·
  <a href="#quick-start">Quick start</a> ·
  <a href="#five-format-native-surfaces">Editors</a> ·
  <a href="#native-automation">Automation</a> ·
  <a href="#architecture">Architecture</a> ·
  <a href="CONTRIBUTING.md">Contributing</a>
</p>

---

A3S Office is an open-source Office engine for product teams that need rich
browser editing and agent-ready file automation without adopting an A3S
backend. It ships complete editor surfaces, typed host contracts, browser file
workflows, and a separate Rust automation plane.

- **Embed complete editors** through React, Vue 3, Web Components, or the
  framework-neutral Core API.
- **Keep format-native behavior** instead of forcing every file through one
  lowest-common-denominator model.
- **Own the product boundary**: your application controls content, storage,
  permissions, collaboration, and model providers.
- **Automate deterministically** through the native CLI, standard MCP server,
  or packaged Office Skill.

## See it working

The images below are committed visual-regression baselines from the real
[Playground](https://a3s-lab.github.io/Office/), not conceptual mockups.

<p align="center">
  <a href="visual-tests/__snapshots__/linux/desktop-1280/document.png">
    <img src="visual-tests/__snapshots__/linux/desktop-1280/document.png" alt="A3S Office document editor with a ribbon and paginated project brief" width="1280">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png" alt="A3S Office spreadsheet editor showing a quarterly execution plan">
      </a>
      <br><sub><strong>Spreadsheet</strong> — workbook editing and formula workflows</sub>
    </td>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/presentation.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/presentation.png" alt="A3S Office presentation editor with slide thumbnails and a slide canvas">
      </a>
      <br><sub><strong>Presentation</strong> — structured slides, objects, and presenter flows</sub>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/markdown.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/markdown.png" alt="A3S Office Markdown editor in synchronized source and preview mode">
      </a>
      <br><sub><strong>Markdown</strong> — GFM source and synchronized preview</sub>
    </td>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/pdf.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/pdf.png" alt="A3S Office PDF editor with search, annotation, save, and download controls">
      </a>
      <br><sub><strong>PDF</strong> — PDFium rendering, forms, annotations, and save</sub>
    </td>
  </tr>
</table>

## Why A3S Office

- **Product-native UI** — Complete Office-style surfaces with no required
  backend, account system, or storage model.
- **Accessible responsive shell** — Compact sidebars, slide navigation, AI
  panes, responsive chart inspectors, dialogs, menus, and popovers share
  bounded keyboard navigation, background isolation, topmost Escape handling,
  and focus restoration.
  Persistent desktop navigation and the temporary phone drawer keep separate
  open states, so live breakpoint changes never turn a desktop sidebar into an
  unexpected blocking modal or steal focus from the workspace.
  Shared color palettes expand to an eight-column touch layout on phones while
  preserving spatial arrow-key navigation.
- **Editor-scoped zoom** — Status controls and Ctrl/Cmd + mouse-wheel gestures
  share each surface's bounded zoom model without changing the host browser's
  page scale.
- **Predictable state** — Controlled content values, typed callbacks, explicit
  file actions, conflict-aware document edits, and one document typography
  baseline across editing, preview, and PDF rendering. Word editing and
  read-only preview retain the same canonical TipTap tree, while browser PDF
  export captures that same Worker/WASM page layout instead of rebuilding
  content from compatibility HTML. Caption numbering and cross-reference
  validity also update in that same transaction graph, including a truthful
  missing-target state after deletion. Paired Word bookmarks can span blocks,
  retain stable native identities through edits, drive live editable REF
  fields, and keep internal links distinct from external hyperlink
  relationships; deleting a target exposes truthful missing-link and
  missing-reference states that undo repairs. Footnote and endnote references
  remain paired with one editable definition, renumber live in independent
  reference-order sequences, receive new identities when copied, and are
  deleted or restored together through one undoable transaction. Body `PAGE`,
  `NUMPAGES`, `SECTION`, `SECTIONPAGES`, `DATE`, and `TIME` fields are atomic,
  copied under fresh stable identities, and resolve from the live Worker/WASM
  page containing each field. Automatic reflow updates numeric fields without
  adding history or ticking clock fields; F9 refreshes every field in one
  undoable action. Safe inline DOCX fields round-trip natively, while nested,
  incomplete, cross-paragraph, deleted, or instructionless structures stay
  text and produce an explicit compatibility warning.
- **Framework choice** — React components, Vue 3 adapters, Custom Elements,
  and a framework-neutral Core API over the same engine.
- **Responsive computation** — Lazy editor chunks, cancellable Workers,
  CSS-compatible font-weight matching, Rust WebAssembly layout and
  calculation, and PDFium rendering.
- **AI without UI scraping** — Typed agent ports and host-defined selection
  actions receive structured context and editing commands.
- **Automation outside the browser** — The native Rust CLI, standard MCP
  server, and Office Skill share bounded mutation contracts.

## Quick start

### Try the product locally

Node.js 20+, Bun 1.3+, and Rust 1.85+ are required.

```bash
git clone https://github.com/A3S-Lab/Office.git
cd Office
bun install --frozen-lockfile
bun run playground
```

Then open the local URL printed by the development server. For a zero-install
tour, use the [live Playground](https://a3s-lab.github.io/Office/).

### Embed a controlled React editor

Install the public package with its React peers:

```bash
bun add @a3s-lab/office react react-dom
```

Import the stylesheet once, give the editor an explicit-height host, and store
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

The editor owns editing, layout, import/export, and browser rendering. The host
owns persistence and decides when, where, and how the emitted content is saved.

### Export the live Word layout to PDF

Give each mounted document a stable `artifactId`, then pass the matching,
current artifact to `downloadArtifactPdf`. Export waits for the live pagination
surface and crops the physical pages computed by the editor:

```tsx
import { useState } from 'react';
import {
  createArtifact,
  downloadArtifactPdf,
} from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';

export function DocumentWithPdfExport() {
  const [artifact, setArtifact] = useState(() =>
    createArtifact('blank-document'),
  );

  if (artifact.content.type !== 'document') return null;

  return (
    <main style={{ display: 'flex', height: '100dvh', flexDirection: 'column' }}>
      <button
        type="button"
        onClick={() => void downloadArtifactPdf(artifact)}
      >
        Export PDF
      </button>
      <section style={{ flex: 1, minHeight: 0 }}>
        <DocumentEditor
          artifactId={artifact.id}
          content={artifact.content}
          onChange={(content) =>
            setArtifact((current) => ({
              ...current,
              content,
              revision: current.revision + 1,
              updatedAt: Date.now(),
            }))
          }
        />
      </section>
    </main>
  );
}
```

The live path preserves automatic page-break decorations, shaped text, table
continuations, and page chrome. It currently rasterizes each physical page
into the PDF; searchable text and vector output remain future fidelity work.

### Choose an entry point

- `@a3s-lab/office/react` — Lazy React editor components and preload helpers.
- `@a3s-lab/office/vue` — Vue 3 adapters with `v-model:content`.
- `@a3s-lab/office/web-component` — Custom Elements for framework-agnostic UI
  composition.
- `@a3s-lab/office/core` — Typed models, templates, import, export, and browser
  file workflows.
- `@a3s-lab/office/styles.css` — Shared editor and interaction-system styles.

Copyable React, Vue, and Web Component examples live in the
[component documentation](https://a3s-lab.github.io/Office/docs/components/).

## Five format-native surfaces

Each surface keeps a canonical model that matches its file format and user
interaction model.

- **Document** — Pagination, sections, focused page-setup tabs, clean page
  margins with on-demand header/footer editing, outline navigation with
  contextual full-text results, keyboard-operated live physical-page raster
  thumbnails with viewport-bounded capture, and shared virtualized
  long-document windows for page buttons, heading rows, full-text results, and
  tracked-revision and anchored-comment review that preserve native scroll
  distance, sparse current/selection/draft pins, and Home/End reachability;
  a WPS-familiar command catalog, quick-access undo and redo, a priority-aware
  adaptive and persistently collapsible ribbon whose tabs can temporarily
  expose commands; WPS-oriented Insert groups and direct Page Layout presets
  for margins, orientation, paper size, and columns, with matching advanced
  Page Setup destinations; WPS-oriented References, Review, and View groups,
  direct tracked-change navigation and decisions, contextual object tabs, and
  viewport-calculated 100%, one-page, and page-width zoom; an actionable status
  bar with live word-count details, WPS `Ctrl+Shift+G`, keyboard-traversable
  view and zoom controls, and compact Web prioritization; editor-scoped WPS
  formatting, paragraph, style, and review shortcuts, plus a permission-free
  formatting clipboard and one-shot format painter shared with compatible
  header and footer formatting, paragraph styles and phone-sized
  paragraph-spacing and pagination controls
  with touch targets and exact invoker-focus restoration; stable compact-ribbon
  edge paging and viewport-bounded list galleries with touch-sized numbering
  actions, selection-preserving bullet and numbering commands, and explicit
  Escape-to-invoker focus restoration;
  grouped bundled, common system, monospace, imported, and host-provided font
  choices with live typeface previews, independent table layout algorithms,
  auto, percentage, or pixel preferred widths, left/center/right placement,
  indentation, table-level cell margins, cell-level margin overrides, and
  rendered column measurements for autofit tables. One Table Properties dialog
  presents table, row, column, and cell tabs. It combines preferred width,
  placement, and indent with selected-row height and pagination, current-column
  width, and selected-cell vertical alignment and margins. The complete draft
  commits in one TipTap transaction and one undo record while untouched imported
  measurements and partial margin inheritance retain their exact source values.
  Cancel and Escape discard the draft and restore the exact ribbon invoker;
  compact layouts keep tabs, numeric steppers, and footer actions touch-sized. A
  contextual Picture Properties workflow combines centimeter width and height,
  a per-image aspect-ratio lock, inline/square/top-and-bottom wrapping,
  alignment, text distance, and alternative text. It commits only changed
  fields in one TipTap history entry, preserves untouched imported dimensions
  exactly, and retains the image selection and invoker focus after apply,
  cancel, or Escape on desktop and phone layouts. A
  reusable table-border pen with all, outside, inside, and individual-edge
  targets, per-edge DOCX border preservation, and bounded `basedOn` table-style
  inheritance with whole-table, banded row or column, first/last row or column,
  and corner-cell formatting. Conditional fills, borders, bold, italic,
  underline, strikethrough, fonts, and text colors are materialized before
  direct table, cell, paragraph, and run formatting. Conditional paragraph
  alignment, direction, indents, spacing and line rules, pagination rules, and
  tab stops use the same precedence chain. Theme tint and shade values resolve
  for table borders and cell fills across edit, preview, and RGB-stable export.
  Splittable table rows that exceed a full physical page
  continue at paragraph boundaries with repeated heading rows on every page,
  viewport-safe comment drafting and focus-preserving citation drafts, tracked
  changes, notes,
  pointer- and Shift+F10-accessible host-defined selection menus, and shared
  edit/preview/PDF typography and page-chrome placement. Editing, read-only
  preview, and browser PDF export share one live pagination result. Tail edits
  reuse the stable physical-page prefix, including resolved first/odd/even
  headers, footers, page numbers, and navigation descriptors.
  Real DOCX fixtures provide deterministic A3S Test evidence for inherited
  table styling across edit and preview, a centered 62.5% table with table and
  cell margin overrides, all four Table Properties tabs at 390 px, exact
  preservation of untouched cell-margin edges, a complete 390 px Picture
  Properties flow, and a 120-paragraph row across three physical pages through
  the final table content and following paragraph.
  A direct-formatting A4 fixture also runs through a real WPS Writer PDF export
  and a page-only browser comparison. Its automatic line spacing, paragraph
  gaps, centered fixed table, row heights, borders, fills, and cell margins
  stay within the checked one-pixel layout-landmark budget while preserving
  the original OOXML line multiples for DOCX export.
  Separate 30-row common-font, 36-row CJK-font, 18-row document-grid, and
  30-row multilingual-script matrices now gate text-band positions against
  WPS. DOCX import selects the Word `ascii`, `hAnsi`, `eastAsia`, or `cs` font
  slot from the run script, including complex-script bold, italic, size, and
  RTL overrides. The browser keeps the measured per-font automatic-line
  advance, while DOCX round trips retain the section `docGrid` type and line
  pitch and each run's `snapToGrid` override.
  _DOCX import/export; PDF export._
- **Markdown** — GFM source, visual editing, synchronized and resizable split
  preview, source-native undo/redo with typing coalescing and selection restore,
  source-aware ribbon formatting and shortcuts, empty-source guidance, task
  lists, tables, links, images, code, a flat reading surface with a comfortable
  line length, and keyboard-accessible host-defined selection menus across both
  editing surfaces. At phone widths, Source and Preview each use the complete
  workspace instead of being compressed into two stacked panes. _Markdown
  import/export._
- **Spreadsheet** — Multiple sheets, formulas, dependency-aware recalculation,
  live selection statistics, an A3S-owned worksheet bar with lifecycle and
  color controls, safe deletion confirmation, in-place rename validation, a
  grid-aligned single-row responsive workbook footer that keeps the active
  worksheet and its actions visible across window and phone-width changes,
  selection-preserving undo/redo, deterministic Arrow, Enter, Tab, Home,
  PageUp/PageDown, row, column, and all-cells selection shortcuts, Cmd/Ctrl
  formatting through root-scoped capture before vendor grid listeners, an
  editor-owned Cmd/Ctrl+F search bar that finds displayed values, raw values,
  formulas, and sparse cells without opening browser Find, with a viewport-safe
  phone layout, touch-sized controls, and exact grid-focus restoration,
  desktop workbook task panes that become focus-contained phone dialogs with
  an inert workbook background and exact ribbon-invoker restoration,
  a WPS-familiar command catalog and tab order, quick-access Undo and Redo,
  a priority-aware adaptive and collapsible ribbon with temporary tab
  expansion, an executable Home clipboard group whose Paste, Cut, and Copy
  commands share the permission-resilient local fallback with their WPS
  shortcuts, a WPS-style Format Painter with single-use and double-click locked
  sessions, cross-sheet range-pattern tiling, Escape cancellation, and native
  style-only writes, Conditional Formatting under Home and Styles, executable
  sorting plus WPS-style AutoFilter under Data, finite current-region discovery,
  `Cmd/Ctrl+Shift+L` toggling, an `Alt+ArrowDown` keyboard filter menu,
  a Home and Cells Rows and Columns menu for inserting above, below, left, or
  right and deleting selected rows or columns through the shared structure
  command port, a Home and Alignment Merge and Center split control with
  Merge Cells, Merge Across, Unmerge Cells, Unmerge and Fill, and the WPS
  `Ctrl+M` shortcut through one controlled native workbook batch,
  a Home and Editing Clear menu for independently removing content, formats,
  comments, hyperlinks, or all cell state while preserving merge geometry,
  with Delete and Backspace mapped to Clear Contents,
  a View and Window Freeze Panes menu for current-cell, top-row, first-column,
  and unfreeze patterns with XLSX round trips, and visible F9 workbook
  recalculation,
  Shift+F11 creation, Ctrl/Cmd+PageUp/PageDown switching in edit and read-only
  preview,
  one shared cell/worksheet context-menu surface with executable shortcut
  hints that becomes a touch-sized, scrollable bottom action sheet on phones,
  keyboard-operated cells and sheet tabs, direct type-to-edit,
  permission-resilient multi-cell cut/copy/paste, clear, F2 editing, and
  focus-safe Escape behavior, direct
  font-family, vertical-alignment, text-wrap, number/percent, and decimal
  controls, charts, validation, protection, comments, and print settings.
  _XLSX, XLS, ODS, and CSV import; XLSX and PDF export._
- **Presentation** — Typed slide scene graph, multi-selection, groups, object
  transforms, keyboard-accessible table-size insertion, Shift+F10-accessible
  native slide and object context actions, guides, comments, transitions,
  editor-scoped formatting and clipboard shortcuts that restore the selected
  object, a docked desktop chart inspector that becomes a focus-contained modal
  surface whenever it overlays the canvas, docked desktop comment review that
  becomes a readable, touch-sized, focus-contained full-editor modal on phones
  with dirty-draft cancellation and exact invoker restoration, a dismissible
  phone slide navigator, a top-aligned phone canvas with a priority-aware status
  bar, one-step slideshow from the beginning or current slide,
  keyboard-complete playback, and a responsive presenter view with one
  navigation strip, current/next context, speaker notes, and a session timer.
  _PPTX import/export; PDF export._
- **PDF** — PDFium rendering, navigation, search, form filling, annotations,
  annotation color, opacity, compatible stroke-width controls, history,
  a scrollable page-thumbnail rail with current-page synchronization and
  focus-synchronized Arrow/Home/End navigation, bounded rendering for long
  files, and a dismissible page drawer whose phone trigger stays in the page
  controls instead of covering PDF content,
  focus-safe page and search drafts, responsive search-result, navigation, and
  zoom controls, editor-scoped shortcuts, and a keyboard-operated overflow menu
  that retains secondary annotation tools and appearance settings on phones.
  _PDF open/save._

Document and Markdown accept public TipTap Extensions. Spreadsheet,
Presentation, and PDF expose stable host ports rather than their internal
command contexts. Editor engines and large runtime assets load only when that
surface is requested.

## Controlled by design

A3S Office is headless at the product boundary, not at the UI boundary. The
package includes complete toolbars, ribbons, panes, popovers, and dialogs while
leaving product infrastructure to the host.

**A3S Office owns:** editing models and commands; import, export, layout, and
rendering; editor UI and responsive interactions; typed selection and agent
ports.

**Your host owns:** persistence and version history; identity, permissions,
and collaboration; the application shell and navigation; AI providers,
prompts, policy, and request lifecycle.

Document and Markdown selection menus can be replaced with host-defined typed
actions. Each action receives an immutable selection snapshot, nearby text,
the complete controlled content, and conflict-aware `replaceText`,
`insertBefore`, `insertAfter`, and `copyText` commands. Markdown snapshots also
identify whether the source or visual surface owns the selection. Async edits
track unrelated visual-editor transactions and fail with `stale-selection`
instead of changing the wrong text. The Playground's open-ended question action
enters a focused draft before dispatch, so a host never receives an unfinished
“Question:” request; attached context stays available without dominating the
assistant surface.

## Browser file workflows

The Core API creates typed blank artifacts and performs browser-side import or
export without mounting an editor:

```ts
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '@a3s-lab/office/core';

const imported = await importOfficeFile(file);
const output = await createArtifactBlob(imported);
const blankDeck = createArtifact('blank-presentation');
```

Use `downloadArtifact` to start a browser download or
`createArtifactBlob` when your application owns upload and persistence.
Imported DOCX artifacts are source-backed: safe source-only OPC parts,
content-type registrations, and relationships survive a regenerated export.
Persist the original Blob alongside the artifact and call `registerSourceBlob`
after a browser reload. If that source is unavailable, DOCX export fails
explicitly instead of silently dropping complex package state. The artifact's
source metadata carries a SHA-256 fingerprint, so registering a different DOCX
under the same artifact ID also fails. Generated core parts remain
authoritative; compatibility diagnostics identify known inline OOXML
normalization and the deliberate removal of invalid signatures, VBA, ActiveX,
and custom-ribbon parts. Within `word/settings.xml`, relationship-free
ignorable attributes, elements, and structurally valid, non-conflicting
`mc:AlternateContent` blocks survive strict or transitional UTF-8/UTF-16
sources. Generated Word settings still win, and this preservation does not yet
restore behavior-changing settings. Regenerated `word/styles.xml` and
`word/numbering.xml` also retain relationship-free passive extensions at the
root and on uniquely matched identities. Styles match by type plus style ID;
imported abstract-numbering, concrete-numbering, and level metadata follows
regenerated IDs. Source-only or duplicate identities, malformed trees,
relationship-bound content, and ambiguous one-to-many numbering mappings are
dropped. Generated Word style and numbering semantics still win. In regenerated
document, header, footer, footnote, and endnote parts, relationship-free passive
extensions from non-OOXML ignorable namespaces also follow uniquely matched
picture drawings, using normalized anchor and drawing-property IDs. Body,
header, footer, footnote, and endnote imports retain those image identities in
sanitized editable HTML. Passive extensions also follow uniquely matched,
unchanged paragraphs and their paragraph properties by native `w14:paraId`
plus `w14:textId`. Body and page-chrome HTML
retain these identities; text edits rotate `textId`, formatting-only edits and
moves keep it, and copies or splits receive new paragraph IDs. Source-only,
duplicate, changed-text, relationship-bound, Microsoft/OOXML semantic, and
ambiguous branches are dropped; generated paragraph and drawing semantics stay
authoritative. Stable table hierarchies use native row `w14:paraId` plus
`w14:textId`, ordered directly owned row IDs for tables, and directly owned
paragraph IDs for cells. Passive extensions on `w:tbl`/`w:tblPr`,
`w:tr`/`w:trPr`, and `w:tc`/`w:tcPr` survive body or page-chrome regeneration.
Row text or structural edits rotate the row version; formatting-only edits and
moves retain it, copies receive independent IDs, and nested rows or cells are
isolated from their outer table. Duplicate or cross-kind identities and unsafe
extension branches fail closed, while generated table geometry and formatting
win. Imported footnotes and endnotes now retain their native positive `w:id`
across reorderings, while copies receive independent IDs. Signed native comment
and reply IDs, reply parentage, and resolved state also survive regeneration.
Native DrawingML pictures inside footnotes and endnotes retain their identity,
layout, wrapping, crop, and layer metadata through public import and artifact
export. Export repairs missing image relationships in generated note parts,
allocates collision-free relationship IDs, and validates each media payload.
Changed, duplicate, namespace-spoofed, relationship-bound, or semantic drawing
branches stay disconnected; generated geometry and media remain authoritative,
while legacy VML, shapes, and SmartArt normalize.

Supported native OMML equations now survive as bounded structured objects in
the document body, headers, footers, footnotes, and endnotes. Inline and display
math, Unicode runs with literal/normal-text semantics, math script/style,
manual-break, and alignment-point properties, common fractions, scripts,
left-side pre-sub/superscripts with empty script slots, radicals with optional
degrees, functions, n-ary operators, combining accents, overbars and underbars,
group characters with
explicit grouping-character placement and baseline justification, phantoms with
visible or hidden bases, independently zeroed width, ascent, or descent, and
transparent spacing, border boxes with independently visible edges and four
strike directions, semantic boxes with
operator-emulation, no-break, differential-spacing, manual-break, and alignment
properties, bounded rectangular matrices with explicit column alignment,
row-spacing and column-gap rules, and minimum column widths,
equation arrays with 1–64 rows, vertical base alignment, maximum/object
distribution, row-spacing rules, and `&` alignment/spacer markers, lower and
upper limit objects, and delimiters regenerate as `m:oMath` or `m:oMathPara` and
render an accessible MathML preview. Bar placement preserves the distinct OMML
defaults for an omitted `barPr` and an omitted `pos`. Group-character
normalization separately preserves an absent `chr` as U+23DF, an explicitly
empty `chr`, bottom `pos`, and the absent-versus-empty `vertJc` defaults.
Phantom normalization preserves the visible `show` default and disabled
`zeroWid`, `zeroAsc`, `zeroDesc`, and `transp` defaults; MathML preview uses
`mphantom` and `mpadded` without discarding the native spacing properties.
Pre-scripts preserve required `sub`, `sup`, and `e` ordering and map empty left
script slots to MathML `none` children after `mprescripts`.
Right-side `sSup`, `sSub`, and `sSubSup` objects enforce their property-first
argument order. `sSubSupPr` preserves `alnScr`, canonicalizes its absent or
disabled value to unaligned scripts, and retains the enabled state through
native export. Supported fraction, script, limit, radical, function, n-ary,
accent, bar, group-character, phantom, border-box, box, matrix,
equation-array, and delimiter property containers preserve one optional ordered
`m:ctrlPr` control format through the same bounded Word run-property model.
The control may contain a direct `w:rPr` or bounded tracked provenance rooted
at `w:ins`, `w:del`, `w:moveFrom`, or `w:moveTo`. Each revision retains a
non-negative 32-bit `w:id`, a bounded `w:author`, an optional validated
`w:date`, and optional Microsoft 365 `w16du:dateUtc` with a UTC `Z` suffix.
Word's legal `moveFrom/moveTo -> ins/del` and `ins -> del` chains are
preserved, with the optional `w:rPr` at the deepest level. Every supported
`deg`, `den`, `e`, `fName`, `lim`, `num`, `sub`, and `sup` argument slot
preserves the same direct or revision-wrapped control format after its
expressions. Empty `ctrlPr` or direct `w:rPr` values canonicalize away, while
an empty revision remains native provenance. Safe object-control values project
only onto separable MathML control/operator nodes; argument-slot formatting and
all revision provenance remain native metadata because professional MathML has
neither linear-build control characters nor Word review/move-range semantics.
Document-level move-range pairing is not inferred from an isolated equation.
Matrix properties follow the ordered
`baseJc -> plcHide -> rSpRule -> cGpRule -> rSp -> cSp -> cGp -> mcs -> ctrlPr`
grammar. Row and column rules accept single, 1.5, double, exact, and multiple
spacing. `rSp` and `cGp` are bounded to 65,535, while the minimum column width
`cSp` is bounded to 31,680 twips. If any spacing property is
present, omitted or attribute-free peers take their Word defaults and native
export emits a complete canonical spacing group. Row spacing and column gaps
project to MathML `rowspacing` and `columnspacing`; `cSp` remains native-only
for layout because MathML `columnwidth` is a fixed width rather than Word's
minimum.
Fractions enforce optional `fPr` before required `num` and `den` arguments.
An absent `type` or an attribute-free `type` canonicalizes to `bar`; `noBar`,
`skw`, and `lin` remain distinct through native export and MathML projection.
Radicals enforce `radPr`, optional `deg`, and `e` ordering. An omitted or empty
degree normalizes to a square root, while a visible nonempty degree remains an
nth root. Native export emits the canonical `radPr -> deg -> e` shape and uses
`degHide=1` with an empty degree slot for square roots.
Functions enforce optional `funcPr` before required `fName` and `e` slots. Both
required slots may be empty. Every supported `CT_OMathArg` slot may likewise be
empty and follows `argPr -> expressions -> ctrlPr`. Its optional trailing
`ctrlPr` retains one bounded direct or revision-wrapped Word control; fixed
slots use named metadata, while matrix cells, equation-array rows, and delimiter
arguments use strictly dimension-aligned metadata. Absent or empty
argument/control properties and
absent, empty, or zero `argSz` values normalize to the default. Bounded
`argSz` values from -2 through 2 round-trip as relative argument sizes. The
Word-effective `box/e`, `groupChr/e`, `limLow/lim`, `limUpp/lim`, `nary/sub`,
`nary/sup`, `rad/deg`, `sPre/sub`, `sPre/sup`, `sSub/sub`, `sSubSup/sub`,
`sSubSup/sup`, and `sSup/sup` pairs project to inverse-sign relative MathML
`scriptlevel`; valid sizes in other argument slots remain native metadata.
Out-of-range or malformed sizes, duplicate or misplaced properties, malformed
control-revision identities or nesting, and semantic argument properties fail
closed.
N-ary operators enforce optional `naryPr` before required `sub`, `sup`, and `e`
slots. An omitted `chr` defaults to U+222B, while an attribute-free `chr`
remains an explicitly empty unsupported operator; an attribute-free `limLoc`
defaults to `undOvr`. Omitted or disabled `grow` values normalize to the
non-growing default; an attribute-free or enabled `grow` round-trips and maps
to MathML `stretchy=true`. Native export always emits both limit slots with
`subHide` or `supHide` for absent scripts.
Delimiters require optional `dPr` before 1–32 `e` arguments and preserve empty
argument slots. Their properties follow
`begChr -> sepChr -> endChr -> grow -> shp -> ctrlPr`; omitted characters
normalize to `(`, U+2502, and `)`, while attribute-free character properties
remain explicitly empty. Omitted or enabled `grow` and omitted, attribute-free,
or `centered` shapes canonicalize to the growing centered defaults. Non-growing
and `match` shapes round-trip in schema order. MathML projects fixed delimiters
with `stretchy=false` and content-matched growing delimiters with
`symmetric=false`; Word ignores shape while delimiter growth is disabled.
Display equations preserve `left`, `right`, `center`, and `centerGroup`
paragraph justification. The bounded native grammar accepts one optional
`m:oMathParaPr` before one `m:oMath`, while absent properties, absent `m:jc`,
and an attribute-free `m:jc` all canonicalize to the `centerGroup` default.
Math-run properties preserve the ordered `lit`, `nor`, `scr`, `sty`, `brk`, and
`aln` grammar, canonicalize Roman/italic and disabled defaults, bound `alnAt` to
1–255, and project supported script/style combinations through MathML
`mathvariant` while retaining native break and alignment metadata.
Math runs also preserve the native `m:rPr -> w:rPr -> m:t/w:t` order. The
bounded Word run-property subset covers direct and theme font references,
Latin and complex-script bold/italic flags, strike and double-strike,
proofing/grid flags, direct and theme colors with tint/shade, half-point font
sizes, colored underline styles, RTL/complex-script flags, and Latin, East
Asian, and bidi language tags. Explicit on/off values remain distinct, export
uses canonical `m:rPr -> w:rPr -> m:t`, and the MathML preview projects safe
direct color, size, font, direction, language, emphasis, and decoration values.
Unknown, duplicated, reordered, namespace-spoofed, or relationship-bound Word
run properties fail closed instead of being silently discarded.
Border-box, box, and equation-array flags retain their semantic defaults, and
manual-break alignment indices are bounded to 1–255. Strict and transitional
UTF-8/UTF-16 math is normalized for editing. Invalid or non-combining accent
characters, malformed math-run or function structures, invalid or contradictory
fraction, radical, n-ary, delimiter, bar, group-character, phantom, border-box,
box, or equation-array properties,
malformed script-property, pre-script, math-paragraph, or lower/upper limit
structures,
malformed, duplicated, reordered, or out-of-range matrix spacing/gap
properties, ragged or over-limit matrices,
over-limit equation arrays,
malformed, misplaced, over-budget, namespace-spoofed, nested, or
relationship-bound math is flattened to bounded text and reported instead of
being trusted or silently attached to another equation.
Relationship-free passive extensions on uniquely matched `w:footnote`,
`w:endnote`, `w:comment`, and `w15:commentEx` roots are retained, and valid
`commentsIds` durable IDs are rebound to each regenerated final comment
paragraph. Duplicate or namespace-spoofed identities, deleted semantic records,
relationship-bound branches, and unsupported modern reaction/people sidecars
fail closed. Within uniquely matched note, comment, and reply records,
text-stable direct paragraphs and runs retain eligible passive extensions on
`w:p`, `w:r`, and `w:rPr`. Safe unmodeled note properties also survive, while
unchanged plain-text comments recover relationship-free source run boundaries
and formatting. Stable `w:hyperlink` wrappers now retain safe tooltips and
passive metadata. Generated note destinations remain authoritative; unchanged
comments and replies recover validated HTTP(S), `mailto`, or internal-anchor
destinations, with external relationship IDs deduplicated or rewritten after
collisions. Text-stable static rich-text and plain-text content controls also
recover eligible inline or contiguous block wrappers, aliases, tags, locking,
signed IDs, Word 2013 appearance and color, end-character formatting, passive
metadata, and stable runs. Editable footnote and endnote tables now regenerate
as native `w:tbl` blocks rather than flattened row text. Rich-text block
controls can safely span structurally stable paragraphs, tables, and nested
tables; matching includes row/cell shape, grid spans, merge state, nested block
shape, and exact paragraph text while generated table geometry wins. Colliding
control IDs are rewritten while unconflicted source IDs remain stable. Text or
table-structure edits, duplicate paragraphs or properties, missing or malformed
hyperlink relationships, wrong target types or modes, unsafe or relative
targets, combined external-plus-anchor destinations, namespace spoofing,
active data bindings or placeholder state, form or nested controls,
relationship-bound content, math, drawing-bearing control wrappers, and other
unsupported wrappers fail closed or normalize instead of being attached to the
wrong content. Source font-table metadata and source-only internal obfuscated
font payloads are also retained, with relationship references rewritten after
ID collisions. External fonts, wrong relationship or content types, duplicate
identities, and paths that collide with generated payloads are not reconnected.
Native DOCX consumers can use the retained embedded fonts; the browser editor,
preview, and PDF renderer do not load document-embedded font binaries and may
substitute fonts or wrap text differently.

## Native automation

The repository also contains a native Rust engine for deterministic reads,
validation, mutation, batch operations, exact natural-unit inventories,
sibling-isolated semantic previews, screenshots, file watching, and a
source-bound exact-layout raster boundary across Office packages. The first
layout-authoritative route supports image-only PPTX slides whose single opaque
PNG exactly covers the declared slide surface. The optional Rust `pdfium`
feature adds bounded, one-based PDF page inventory and exact page PNGs through
an explicit host-supplied PDFium 7881 library. It records media/crop boxes,
rotation, physical and pixel geometry, source and engine hashes, and never
downloads a runtime or introduces a Browser dependency. Consumers can inspect
selected pages from one previously validated complete inventory without
rescanning the full document; render still revalidates the immutable source and
actual page profile before publication. The same retained inventory now
authorizes bounded native PDF text-layer extraction with source-order Unicode,
exact UTF-8/UTF-16 ranges, optional glyph boxes, and PDFium-native same-line,
same-style runs with exact PDF-coordinate bounds. Independent character and run
limits keep the result bounded. Document outlines retain exact page targets.
Text and outline calls revalidate
the immutable source, reuse inventory authority, and return typed limit or
unsupported failures without OCR or Browser access. Richer slides and formats
without an authoritative provider remain typed unsupported instead of being
relabeled semantic previews.
Native DOCX and PPTX table reads also normalize merged cells into one-based
logical row and column coordinates, row and column spans, and stable anchor
references for covered physical cells.

```bash
# Run from the repository root
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- view report.docx outline --json
cargo run -p a3s-office-cli -- set report.docx /body --find Draft --replace Final --json

# Start the standard MCP server
cargo run -p a3s-office-cli -- mcp
```

CLI, MCP, the typed Rust API, and the packaged Office Skill share the same
bounded contracts. They inspect and modify files without launching desktop
Office or scraping editor UI.

Read the [native engine design](docs/latest/en/native-office-engine.md), the
complete [CLI reference](docs/latest/en/cli-reference.md), or the published
[CLI and Skill guide](https://a3s-lab.github.io/Office/docs/automation/).

## Architecture

<p align="center">
  <a href="assets/readme/architecture.svg">
    <img src="assets/readme/architecture.svg" width="1200" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

The browser plane combines controlled editor surfaces, framework adapters,
Workers, Rust WebAssembly, and PDFium. The native plane keeps filesystem and
OOXML package concerns in a separate Rust core, with a host-injected optional
PDFium provider for read-only PDF page evidence. Both planes expose typed
contracts; neither requires an A3S backend.

For engine ownership, Worker/WASM boundaries, delivery stages, and performance
gates, see
[Browser editor architecture](docs/latest/en/browser-editor-architecture.md).

## Project status

A3S Office is pre-1.0. Content models and component props are public, but a
minor release may still include breaking model changes. Required migrations
will be called out in the [changelog](CHANGELOG.md).

The project targets predictable browser editing and file preservation, not
pixel parity with every Microsoft Office or WPS feature. Unsupported OOXML
semantics and fidelity gaps remain explicit compatibility boundaries instead
of being silently approximated. Track product depth and release evidence in
the [editor quality roadmap](docs/latest/en/editor-quality-roadmap.md).

## Development

Run focused checks from the repository root:

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

Start the integration Playground with `bun run playground`. The full pull
request checklist also covers the Rust workspace, browser bundle budget, and
committed visual contracts; see [CONTRIBUTING.md](CONTRIBUTING.md).

## Documentation

The published website keeps the interactive Playground at its root and builds
the searchable Rspress documentation center under `/docs/`. Simplified Chinese
is the stable default, English remains available from the language menu, and
the version menu switches between `latest` and frozen release documentation.
Search is scoped to the active language and version. Both surfaces use the
same deployment base, so preview, Pages, and fork deployments keep working
without hard-coded return URLs.

- [Live Playground](https://a3s-lab.github.io/Office/)
- [Documentation center](https://a3s-lab.github.io/Office/docs/)
- [A3S Office 0.3.0 documentation](https://a3s-lab.github.io/Office/docs/0.3.0/)
- [A3S Office 0.2.0 documentation](https://a3s-lab.github.io/Office/docs/0.2.0/)
- [A3S Office 0.1.0 documentation](https://a3s-lab.github.io/Office/docs/0.1.0/)
- [React, Vue, and Web Component integration](https://a3s-lab.github.io/Office/docs/components/)
- [Office CLI and coding-agent Skill](https://a3s-lab.github.io/Office/docs/automation/)
- [Browser editor architecture](docs/latest/en/browser-editor-architecture.md)
- [Native Office engine](docs/latest/en/native-office-engine.md)
- [CLI reference](docs/latest/en/cli-reference.md)
- [Editor quality roadmap](docs/latest/en/editor-quality-roadmap.md)
- [Changelog](CHANGELOG.md)

## Community and security

Contributions are welcome. Read the [contribution guide](CONTRIBUTING.md) and
[Code of Conduct](CODE_OF_CONDUCT.md) before opening a change. Report suspected
vulnerabilities through the private process in [SECURITY.md](SECURITY.md), not
through a public issue.

## License

A3S Office is available under the [MIT License](LICENSE). Bundled PDFium and
other third-party assets carry additional notices in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
