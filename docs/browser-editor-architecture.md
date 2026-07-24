# Browser editor architecture

A3S Office uses an editor-specific interaction surface above a shared,
browser-safe Office kernel. TipTap is the right document engine for flowing
rich text, but it is not the canonical interaction model for a spreadsheet
grid, slide canvas, or PDF page.

## Product architecture

| Product | Interaction surface | TipTap responsibility | Kernel responsibility |
| --- | --- | --- | --- |
| Document | TipTap and ProseMirror | Complete logical document, selection, commands, history, comments, and collaboration boundary | Page layout, style and font resolution, OOXML semantics, and serialization |
| Spreadsheet | Virtualized grid and canvas | Rich text inside a cell or floating text object | Formula calculation, workbook semantics, print layout, OOXML, and serialization |
| Presentation | Scene graph and slide canvas | Text inside an individual text box | Masters, layouts, themes, object geometry, OOXML, and serialization |
| PDF | PDFium page surface | None | PDF parsing, rendering, annotation serialization, and document save |

This separation prevents an editor framework from becoming a false abstraction
over products with different selection, layout, and performance requirements.

## Shared platform layers

The products share infrastructure without sharing a false universal document
model:

1. The Office shell owns the ribbon, status bar, dialogs, keyboard routing,
   accessibility, file actions, theming, and responsive behavior.
2. Each editor owns its selection model and exposes typed commands to the
   shell. A command never searches rendered text or infers intent from labels.
3. Interactive editing stays on the main thread. Only the active TipTap text
   surface is mounted in Presentation; inactive text boxes use lightweight
   scene previews.
4. A versioned Worker client schedules bounded jobs, cancels superseded jobs,
   and rejects stale responses.
5. Rust WebAssembly kernels own deterministic calculation: layout, formulas,
   geometry, package parsing, OOXML semantics, and serialization.
6. Versioned structured models are the persistence boundary. HTML, canvas, and
   DOM measurements are views or migration inputs, not the long-term source of
   truth for OOXML documents.

This split keeps typing, selection, drag, and resize responsive while moving
CPU-heavy and memory-bounded work away from the UI event loop.

## Current implementation status

| Product | Implemented browser surface | Implemented kernel boundary | Next fidelity gate |
| --- | --- | --- | --- |
| Document | One TipTap/ProseMirror body tree, controlled TipTap header/footer surfaces with direct paper-margin editing and a contextual ribbon, typed physical-page and section-page descriptors, repeated first/default/even page chrome, a versioned structured model with an HTML compatibility representation, prefix-reused visual-line measurement and pages, page decorations, page-aware horizontal and vertical rulers for page margins, paragraph indents and typed tab stops, structured list-item pagination, explicit paragraph and list-item direction, compact spacing and pagination controls, typed inline/square/top-and-bottom image layout, imported style-inherited paragraph properties, structured inline tabs, and theme-aware run font/size/color/background import | Worker plus resumable Rust/WASM flow pagination and Rustybuzz shaping across exact registered text runs, including eligible list paragraphs, Unicode bidi level segmentation, ordered per-grapheme font fallback, packaged Latin/CJK/Arabic/Hebrew faces, and structured left-to-right tabs, with explicit DOM and JavaScript fallbacks for text affected by supported floats | Language-complete font substitution, complete Word style and numbering coverage, locale-complete and bidirectional tabs, arbitrary floating-object offsets and layering, complex table flow, and loss-preserving OOXML package state |
| Markdown | TipTap visual editing with source and split views | No kernel required for normal editing | CommonMark/GFM compatibility fixtures and large-file profiling |
| Spreadsheet | Fortune Sheet grid integrated with the shared Office shell and a typed command boundary for selected-range formatting, merge state, recalculation, gridlines, and zoom | Native Office formula and OOXML primitives exist outside the browser kernel; no browser calculation kernel yet | Dedicated virtual grid and Worker/WASM dependency graph, calculation, and print layout |
| Presentation | Scene canvas with one TipTap instance for the selected text box and one typed dispatcher for ribbon commands | Revisioned, cancellable Worker/Rust-WASM slide-relative alignment with a JavaScript fallback | Snapping, guides, grouping, connectors, theme resolution, text fitting, and slide serialization |
| PDF | PDFium-backed page rendering with an A3S-owned toolbar and typed capability controllers for navigation, zoom, search, basic annotations, history, and save | PDFium WebAssembly | Annotation styling, forms, redaction review, page organization, and reopen fixtures |

The table is a fidelity statement, not a marketing capability list. The
current Document path shapes text-flow paragraphs in Rust/WASM when every
explicit CSS family resolves to an exact registered face. Each run may carry up
to eight ordered faces; the kernel selects one face per grapheme and coalesces
adjacent selections before shaping. Unicode bidi levels split mixed-direction
text into logical shaping runs, including left-to-right number runs nested in
right-to-left text. Structured left-to-right tabs use normalized stops and
pre-shaped following-segment widths in the kernel. Glyphs missing from the
complete stack, unsupported OpenType behavior, inline objects, tab paragraphs
that resolve any right-to-left run, and unregistered faces deliberately retain
browser line measurement.
Spreadsheet does not yet use a browser Rust/WASM calculation kernel.
Presentation uses Rust/WASM only for alignment to slide bounds; the remaining
geometry and layout operations stay on the main thread until the later stages
below.

Spreadsheet, Presentation, and PDF commands cross explicit typed boundaries.
The shell never searches visible labels, scrapes rendered text, or synthesizes
clicks to infer product intent. This keeps localized UI copy independent from
behavior and gives Worker/WASM operations a stable request contract.

## Document editing flow

```text
Controlled WorkDocumentContent
             |
 versioned structured model
 schema + revision + HTML fingerprint
             |
      TipTap / ProseMirror
      one logical document
             |
 text-run collection and format-specific measurement
             |
       dedicated Worker
             |
 Rustybuzz shaping + Rust WebAssembly layout
             |
 versioned page and break result
             |
 ProseMirror widget decorations
  no document or history mutation
```

The document remains one logical ProseMirror tree. Automatic page boundaries
are widget decorations, not page nodes. Reflow therefore does not split the
content model or corrupt undo, selection mapping, copy and paste, or a future
collaboration protocol.

The current browser-kernel slice collects contiguous geometry-affecting text
runs from eligible paragraphs. Each run carries an ordered, exact registered
font stack, size, line height, letter spacing, ligature, and kerning behavior.
The kernel selects a face for each grapheme, joins adjacent selections that use
the same face, and includes every used face in line ascent and descent.
Rustybuzz shapes those segments and the kernel applies Unicode and
grapheme-safe line breaking across run boundaries. Unsupported paragraphs
retain the existing DOM range path, which maps each browser visual-line start
back to a ProseMirror position. The Worker/WASM kernel chooses page breaks,
keeps a minimum number of line fragments on either side of a break, and
returns mapped decorations.
Tables, images, code blocks, and other complex content use explicit
format-specific measurement. Top-level tables are row flows; eligible rows can
additionally expose synchronized direct-cell block fragments. Ordered and bullet
lists recurse through semantic list items, expose each direct item block at its
stable ProseMirror position, and distribute list-container height without
double-counting nested content.

Document image nodes carry typed layout, alignment, wrap-distance, alternative
text, and size attributes. The selected image exposes the same operations
through a contextual Picture ribbon. Inline images stay in normal flow;
supported left- or right-aligned square images use browser floats, and
top-and-bottom images clear surrounding text. Paragraphs following a square
float remain on the DOM visual-line path because the available line width
changes while the float is active. The paginator reserves and observes the
image block height, then sends the measured text fragments to the same
Worker/Rust-WASM page-layout protocol. Supported square and top-and-bottom
metadata round-trips through DOCX `wp:anchor`. Arbitrary offsets, crop geometry,
contour wrapping, z-order, and drawing-layer relationships are not represented
yet.

The page-layout panel mounts two bounded TipTap surfaces for the active
header/footer variant. Pagination joins each kernel page placement to typed
section metadata and derives its physical page, section page, displayed page
number, and resolved first/default/even chrome variant. Page view fixes the
first paper header and final paper footer to their physical descriptors and
paints the previous footer plus next header inside every automatic page gap.
Those repeated surfaces are non-editable decorations and never enter document
history. Accessible activation targets on the outer paper margins replace the
static HTML with one toolbar-free TipTap surface in place and add a contextual
Header and Footer ribbon. That ribbon uses the same typed command functions for
formatting, alignment, links, images, page-number visibility, switching between
header and footer, and closing the mode. Escape and web-view transitions close
the mode and restore body focus. The body is visually de-emphasized without
being copied or removed from the canonical tree. Both surfaces preserve
semantic paragraphs, lists, tables, links, marks, alignment, color, and inline
raster images while exposing a smaller command set than the document body.
Toolbar controls derive active state from the current selection; they do not
inspect localized labels or use `document.execCommand`. The controlled HTML
values still pass through the page-chrome sanitizer and the
default/first/even-page model, so direct paper editing does not create a second
persistence format or bypass DOCX header/footer relationships.

The editor records the earliest position changed by each ProseMirror
transaction and reuses measured blocks before that position. Runtime
diagnostics report measured and reused block counts, shaped paragraphs, and
submitted and accepted text-run counts. The incremental planner also retains
complete stable pages before the affected region. It rewinds at least one page
and continues rewinding when a paragraph flow or `keepWithNext` relationship
crosses the proposed boundary. The Worker/WASM kernel receives only the safe
suffix with an absolute start-page index; the integration then merges the
retained pages with a newly calculated boundary break. Layout requests carry
both a monotonic document revision and an independent layout revision,
allowing font, viewport, or page-size work to supersede an older layout without
changing the document identity.

It also supports explicit page and section breaks, keep-together blocks,
keep-with-next headings, cancellation, stale-revision rejection, and automatic
reflow after editing, resizing, font loading, or zoom changes. Page view uses
the result for visual paper gaps and status-bar page counts. Its horizontal
ruler is bound to the active section's paper width and left/right margins and
to the active paragraph's left, right, first-line, and hanging-indent
attributes. A vertical ruler controls the active section's top and bottom
margins. Pointer and keyboard changes update the same section node used by
pagination; compact layouts hide the vertical ruler while retaining Page
Layout controls. The paragraph attributes map to and from `w:ind` formatting
during DOCX import and export. Paragraph spacing keeps before/after values plus Word's
`auto`, `exact`, and `atLeast` line rules. Direct `keepLines`,
`keepWithNext`, `pageBreakBefore`, and `widowControl` attributes also
round-trip through their matching OOXML paragraph properties. The kernel
applies those rules to visual-line flows, including reserving the first two
lines of the next paragraph when a keep-with-next boundary fits on a page.
Paragraph direction is a typed `ltr` or `rtl` node attribute rendered as
semantic `dir`, passed to the text-layout request, and controlled from the Home
ribbon. DOCX import resolves `w:bidi` through document defaults, the default
paragraph style, bounded and cycle-safe `basedOn` chains, and direct
overrides, including numbered paragraphs. Export maps explicit paragraph or
list-item direction to `w:bidi` and right-to-left text runs to `w:rtl`. DOCX
import resolves
indents, spacing, and pagination through the same style chain. One shared style
index serves these paragraph import passes. Marker application scans converted
text once per property family instead of once per paragraph. Web view removes
both automatic pagination and the paper ruler.

Run-format import uses the same style index. It merges document run defaults,
paragraph-style run properties, direct paragraph run properties, bounded
character-style `basedOn` chains, and direct run properties. Unique sentinels
survive Mammoth conversion and become ordinary inline spans before the
TipTap/ProseMirror model is created. This currently preserves visible font
families, half-point sizes, direct hexadecimal colors, Word highlights, and
run shading. The importer also reads the DOCX font and color schemes, resolves
major/minor theme font references, system or sRGB theme colors, and tint/shade
transforms. Export maps background colors back to OOXML run shading.

Lists remain semantic ProseMirror structures. Pagination measures each direct
list-item block independently and maps a page widget before the first block
inside the target `li`; eligible list paragraphs are collected for the same
Rustybuzz text-layout request as ordinary paragraphs. DOCX export emits one
paragraph with `w:numPr` per item, supports levels zero through eight, positive
ordered-list starts, common HTML ordered-list types, nesting, bullets, and RTL
paragraph/run properties. Table cells, notes, headers, and footers use the same
export path. DOCX import uses bounded numbering metadata to restore positive
list starts and common decimal, letter, and Roman formats after Mammoth
conversion, while direction markers are applied to the semantic list item.
Reversed lists, per-item `value` overrides, arbitrary restart/continuation
semantics, and loss-preserving custom Word numbering formats remain later
fidelity gates.

This slice is not yet a Microsoft Word or WPS line-layout fidelity claim.
Language-complete font substitution, variable font axes, the remaining
character and table style properties, arbitrary floating-object offsets,
cropping, contour wrapping and layering, row-internal table splitting inside a
single paragraph or an oversized multi-page row, nested tables, footnote
balancing, multi-column flow, and mixed-size sections require the later layout
stages below.

Top-level tables now enter pagination as row flows. DOM measurement supplies
row heights and stable ProseMirror positions; the Worker/Rust-WASM kernel keeps
leading header rows with the first body row and reserves the measured header
height on continuation pages. A table-aware widget creates the paper gap and
an `aria-hidden`, non-editable copy of the rendered header with the original
column group. The editable row remains the only canonical content. TipTap row
attributes preserve `cantSplit` and `tblHeader` through structured HTML and
DOCX import/export. Ordinary rows whose full height fits a continuation page
can split at synchronized boundaries between direct cell blocks. One
non-editable widget per cell aligns the page gap and paints a clipped slice of
the same measured header overlay. The leading cell alone extends the paper-gap
paint across the page margins. `cantSplit` rows, single long paragraphs, and
oversized rows are still atomic and may overflow.

Paragraph tab stops are typed node attributes with normalized positions,
left/center/right/decimal alignment, and leader styles. A leaf
`documentTab` node preserves each Tab keystroke in the ProseMirror model
outside tables and lists; table navigation and list nesting keep priority.
DOM layout resolves default and custom stops after browser font measurement,
normalizes zoomed geometry back to document coordinates, and runs in edit,
preview, and PDF surfaces before pagination or capture. DOCX import merges
`w:tabs` through bounded paragraph-style inheritance, applies `clear` entries
by position, and replaces `w:tab` with sentinels during Mammoth conversion.
DOCX export restores paragraph tab definitions and inline tabs. Deterministic
font shaping now crosses multiple exact registered text runs in one paragraph,
and structured left-to-right tabs share the browser's default/custom stop
selection policy. Center, right, and decimal offsets use Rustybuzz-shaped
following-segment widths. Locale-specific decimal separators, bidirectional tab
paragraphs, and loss-preserving unsupported tab kinds remain later fidelity
gates.

## WebAssembly boundary

`crates/web-kernel` is deliberately independent of the DOM, filesystem,
network, and an async runtime. It accepts bounded JSON requests through a
small raw WebAssembly ABI:

- `office_kernel_abi_version`
- `office_kernel_alloc` and `office_kernel_dealloc`
- `office_kernel_register_font`
- `office_kernel_layout`
- `office_kernel_text_layout`
- `office_kernel_presentation_geometry`
- `office_kernel_result_pointer` and `office_kernel_result_length`

The protocol is versioned and carries independent layout and document
revisions plus an absolute start-page index for resumable suffix layout.
Font registration crosses the ABI as bounded raw bytes rather than base64 in
JSON. Batched text-layout requests carry contiguous UTF-16 run ranges, ordered
font IDs, CSS-pixel metrics, whitespace, ligature and kerning behavior,
direction, and first-line width. The Unicode bidi algorithm resolves one
paragraph level and directional levels before Rustybuzz shapes every
intersection of direction, style, and registered font. Unicode line breaking
and grapheme-safe emergency breaks span run boundaries and produce UTF-16
offsets that map directly to ProseMirror.
Presentation geometry requests carry stable element IDs, slide-relative
coordinates, and an explicit alignment enum. The Worker ignores cancelled
requests, and the React integration ignores stale results. Matching JavaScript
implementations preserve editing if Worker or WebAssembly loading is
unavailable. The JavaScript text fallback explicitly reports unsupported
paragraphs so the editor uses DOM line measurement instead of estimated font
metrics.

The browser pagination implementation is split by capability instead of
accumulating in one editor module. `work-document-pagination.ts` owns only the
TipTap extension and public facade; dedicated modules own block measurement,
table fragmentation, text-layout request collection, visual-line measurement,
DOM geometry, page-break decorations, and shared contracts. The Rust kernel
uses the same separation: the crate root owns the ABI and flow layout, while
text validation, UTF-16 offset mapping, shaping, and their tests live in
focused modules. This keeps the fallback and WASM paths independently
testable without duplicating their public protocol.

The npm package emits `office-kernel.worker.js`, `office-kernel.wasm`, and
default Noto Sans, Noto Sans Hans, Noto Naskh Arabic, and Noto Sans Hebrew
regular faces beside the public JavaScript entries. Hosts may override the
WASM URL and provide typed `layoutFonts` when assets are served from a CDN or a
nonstandard path. The browser loads the same faces through `FontFace`; a
paragraph enters the deterministic path only when every explicit family in its
CSS stack matches a
successfully loaded and registered asset in the same order.

## Structured document boundary

`WorkDocumentContent` remains controlled by the host and keeps `html` as a
required compatibility representation. It may also carry a versioned
`WorkDocumentModel` containing:

- the `a3s.office.document` schema identifier and schema version;
- a monotonically increasing document revision;
- the structured TipTap/ProseMirror root; and
- a fingerprint of the synchronized HTML representation.

`DocumentEditor` emits the structured model and HTML together after an edit.
It loads the model only when the fingerprint still matches normalized HTML.
If a legacy host changes `html` without updating the model, the stale model is
ignored and removed on the next synchronization instead of overwriting the
host's change. Model trees are bounded and validated before they cross into the
editor. React, Vue, and Web Component wrappers continue to share the same
controlled value.

This is a transitional persistence boundary, not loss-preserving OOXML yet.
DOCX still converts through Mammoth, but DOCX, HTML, and text imports create the
structured model immediately through the same extension schema used by the
editor. DOCX export materializes synchronized HTML from the structured model
before OOXML generation, so a stale HTML cache cannot override model changes.
Preview still reads synchronized HTML. Unsupported OOXML package parts are not
yet represented by the model.

Migration to browser-native OOXML is staged:

1. Keep editing behavior stable while moving pagination off the React thread.
2. Persist the versioned logical document model while retaining checked HTML
   compatibility for existing hosts.
3. Replace eligible browser-measured visual-line flows with deterministic
   shaped line boxes, font metrics, configurable widow and orphan rules, notes,
   fields, and full floating-object geometry beyond the supported image-anchor
   subset.
4. Compile the A3S OOXML, relationship, style, numbering, theme, and package
   layers for the browser with explicit memory and archive limits.
5. Extend the structured model with loss-preserving OOXML package state and
   round trips without making HTML the only source of truth.
6. Share style, theme, formula, geometry, and serialization primitives across
   Document, Spreadsheet, and Presentation while retaining their distinct
   editing surfaces.

Each stage needs compatibility fixtures, deterministic layout goldens, large
document performance budgets, and real Microsoft Office and WPS
interoperability evidence before it can be described as fidelity-complete.

## Delivery roadmap

Work is delivered as vertical slices. Each slice must preserve controlled
component behavior and import/export compatibility before the next slice starts.

### Stage 1: shared editing foundation

- Use the same ribbon, status bar, dialogs, fields, zoom controls, loading
  states, shortcuts, and responsive rules in every editor.
- Keep one command route per product and make commands operate on typed
  selections.
- Load Document, Markdown, Spreadsheet, Presentation, PDF, Workers, and WASM as
  independent chunks.
- Keep reproducible Chromium screenshots and shared-shell geometry checks for
  all five editors at 1280 × 800 and 768 × 800.

Exit criteria: no native prompt, confirm, or select UI; no editor-specific copy
of shared chrome; keyboard-only access to all primary commands; no initial
bundle regression.

### Stage 2: Word-compatible document layout

- Finish replacing browser-measured visual-line pagination with shaped line
  boxes produced from resolved fonts and paragraph properties. Text-flow
  paragraphs already support multiple exact registered runs.
- Extend the current horizontal and vertical page rulers with bidirectional and
  locale-complete tab behavior, plus complete style, numbering, and theme
  inheritance beyond the implemented paragraph slices.
- Complete row-internal splitting for single paragraphs and oversized
  multi-page rows, and add nested tables, full floating-object geometry beyond
  the supported square/top-and-bottom image anchors, footnote balancing,
  columns, and mixed page sections. Row-flow pagination, direct-cell block
  splitting, repeating headers, and the initial image-wrapping slice are
  already implemented.
- Keep pagination results as mapped ProseMirror decorations so reflow never
  corrupts selection or undo history.

Exit criteria: deterministic layout goldens for the supported feature set;
incremental reflow does not rebuild unaffected pages; DOCX fixtures round-trip
through Microsoft Word and WPS without losing unsupported package parts.

### Stage 3: spreadsheet calculation and virtualization

- Introduce an A3S-owned sparse workbook model and a row/column virtualized
  viewport.
- Mount TipTap only for the active rich-text cell or floating text object.
- Move formula parsing, dependency graph updates, dirty-cell recalculation,
  number formatting, sorting, filtering, and print pagination to a Worker/WASM
  kernel.
- Keep viewport rendering independent from workbook size and retain a typed
  value for every cell.

Exit criteria: scrolling and selection do not scale with total row count;
incremental recalculation touches only affected dependency subgraphs; XLSX
fixtures preserve formulas, styles, names, validation, conditional formatting,
and unknown package parts within the declared support boundary.

### Stage 4: presentation scene and geometry kernel

- Keep slides as a typed scene graph and mount one TipTap editor only for the
  selected text box.
- Move transforms, snapping, alignment guides, grouping, connector routing,
  text fitting, theme/master resolution, and thumbnail layout to Worker/WASM.
- Add stable object IDs so selection, history, comments, animation, and OOXML
  relationships survive import and export.

Exit criteria: object drag and resize stay interactive on complex slides;
partial rich-text formatting survives PPTX round trips; masters, layouts,
themes, tables, charts, links, and notes have compatibility fixtures for
PowerPoint and WPS.

### Stage 5: PDF product surface

- Keep PDFium as the parsing and rendering engine.
- Replace embedded product chrome with the shared A3S Office shell.
- Own annotation, form, redaction, page organization, search, history, and save
  commands through typed models.

The first PDF shell slice is implemented. It disables the embedded toolbar and
routes page navigation, zoom presets, search result navigation, basic
annotation tools, annotation deletion, history, and copy export through public
plugin capabilities. It contains no shadow-DOM queries, private viewer
selectors, label inference, or synthetic clicks. Existing PDF password
handling remains in the PDF document lifecycle rather than the removed toolbar.
Advanced annotation styling, form-authoring controls, redaction review, page
organization, and compatibility fixtures remain part of this stage.

Exit criteria: annotations and form changes survive reopen; destructive
redaction requires an explicit review step; large documents render a bounded
page window rather than the complete file.

### Stage 6: loss-preserving Office persistence

- Compile the bounded A3S package, relationship, style, numbering, theme, and
  OOXML layers for browser Workers.
- Preserve unsupported parts and attributes instead of regenerating an entire
  package from the visible view model.
- Use the same structured mutation contracts in the browser library, Office
  CLI, MCP server, and Skill.

Exit criteria: import-edit-export fixtures show semantic and package-level
round-trip evidence in Microsoft Office and WPS; unsupported content is
retained or rejected explicitly, never silently discarded.

## Performance gates

Performance claims are measured on a documented reference machine with
repeatable fixtures. The targets below are release gates, not current claims:

- Local typing, selection, and object manipulation have a p95 main-thread task
  below 16 ms; no normal interaction creates a task above 50 ms.
- Incremental work is revisioned and bounded. A new edit cancels stale layout,
  formula, geometry, or serialization jobs.
- Document reflow begins at the earliest affected line and does not remeasure
  stable earlier pages.
- Spreadsheet rendering cost depends on the visible viewport, not sheet
  dimensions; formula cost depends on the dirty dependency subgraph.
- Presentation renders the active slide at full fidelity and virtualizes
  off-screen thumbnails and inactive slide content.
- PDF keeps a bounded render cache and releases page bitmaps outside the active
  window.
- Every editor has fixture-based memory, interaction-latency, load-time, and
  output-size budgets in CI before it can be called production-ready.

## Performance and safety rules

- Editing and selection stay on the main thread; parsing and layout do not.
- The public React entry loads editor engines as independent asynchronous
  chunks. `preloadOfficeEditor` may warm one engine from hover or keyboard
  focus without mounting it.
- Spreadsheet formula-language metadata loads with Spreadsheet diagnostics and
  editing, never with the empty Office workspace.
- The Playground initial JavaScript budget is 220 KiB gzip and is enforced in
  CI against the scripts referenced by the generated `index.html`.
- One active layout request exists per editor. A newer revision cancels the
  previous request.
- Requests reject invalid dimensions and more than 10,000 layout blocks.
- Text requests are bounded to 1,024 paragraphs and 1 MiB of UTF-8 text.
  They accept at most 16,384 contiguous runs, with no more than 4,096 in one
  paragraph and no more than eight ordered faces per run. Font registration is
  bounded to 16 faces and 32 MiB per face.
- Glyphs missing from the complete registered stack, ambiguous font stacks,
  and unsupported inline structures retain DOM measurement instead of
  accepting approximate line boxes.
- WebAssembly performs no network or filesystem access.
- Worker failure is recoverable and does not block typing.
- Automatic pagination transactions are excluded from editor history.
- Multi-column and mixed-page-layout documents currently retain explicit-break
  behavior until their layout protocols are implemented.

## Verification

The browser kernel is covered at four boundaries:

1. Rust unit tests for deterministic pagination, Unicode line breaking,
   grapheme-safe emergency wrapping, whitespace modes, and validation.
2. JavaScript fallback tests for protocol parity, safe page-prefix reuse, and
   no-Worker operation.
3. A raw generated-WASM ABI smoke test that registers both shipped fonts,
   proves the Latin face lacks CJK glyphs, resolves them through the ordered
   fallback face, and verifies mixed-face line metrics.
4. Browser checks for real Worker/WASM/font loading, shaped-line parity with
   browser line boxes at non-100% zoom, real per-grapheme fallback diagnostics,
   explicit unresolved-glyph fallback, page-view reflow, web-view clearing,
   page counts, nested and RTL list flow, undo behavior, and slide-relative
   element alignment.

Run the focused kernel checks with:

```bash
bun run kernel:test
cargo test -p a3s-office-web-kernel
```

Build and verify the Playground performance boundary with:

```bash
bun run playground:build
bun run playground:bundle-check
bun run playground:visual
```
