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
   accessibility, file actions, anchored context menus, theming, and responsive
   behavior. Context menus share pointer and Shift+F10 entry, viewport-bounded
   placement, disabled-item-aware arrow navigation, Escape restoration, and an
   explicit Tab exit that follows document focus order. Window-level shortcut
   listeners are scoped to their mounted editor root and may capture before a
   vendor engine's document listeners, so an editor cannot consume commands
   from host-page controls while its own shortcuts remain deterministic.
   Successful object and grid commands restore focus to the latest controlled
   selection before the next keystroke.
   Writer keeps tab, group, label, and shortcut metadata in one internal
   command catalog. Undo and redo live in a keyboard-described quick-access
   toolbar. The ribbon measures its own toolbar instead of the browser viewport
   and compacts low-priority, then normal-priority groups before exposing group
   paging. It can remain collapsed; activating a tab temporarily overlays that
   tab's commands without moving the document, and an outside pointer action
   returns to the compact row. Writer routes its catalogued WPS font-size,
   alignment, line-spacing, heading, spelling, field-refresh, comment, and
   track-changes shortcuts only while the document surface owns the event. Its
   local formatting clipboard powers WPS copy-format and paste-format shortcuts
   plus a one-shot format painter without requiring browser clipboard access;
   paste replaces only direct character and paragraph formatting, preserving
   links, comments, and revision semantics in one document transaction.
   Header and footer TipTap surfaces register the same WPS alignment and
   copy-format shortcuts. Cross-surface paste retains compatible marks and
   paragraph attributes while falling back from unsupported heading nodes.
   Writer orders Insert commands as Pages, Table, Illustrations, Links, Header
   and Footer, and Text. Page Layout keeps direct margin, orientation, paper,
   and equal-column presets in the ribbon; custom margins and advanced columns
   open the corresponding controlled Page Setup tab instead of duplicating a
   second settings model.
2. Each editor owns its selection model and exposes typed commands to the
   shell. A command never searches rendered text or infers intent from labels.
3. Interactive editing stays on the main thread. Presentation keeps object
   selection separate from content editing and mounts TipTap only for the text
   box explicitly opened with double-click or Enter; inactive text boxes use
   lightweight scene previews.
4. A versioned Worker client schedules bounded jobs, cancels superseded jobs,
   and rejects stale responses.
5. Rust WebAssembly kernels own deterministic calculation: layout, formulas,
   geometry, package parsing, OOXML semantics, and serialization.
6. Versioned structured models are the persistence boundary. HTML, canvas, and
   DOM measurements are views or migration inputs, not the long-term source of
   truth for OOXML documents.

At phone widths, Document, Markdown, Spreadsheet, and Presentation place file
identity and actions on the first compact row and keep unshrunk ribbon tabs in
a second horizontally scrollable row. This preserves tab labels and keyboard
navigation instead of squeezing both command regions into one line. PDF keeps
its purpose-built single compact toolbar and is excluded from that stacking
rule.

This split keeps typing, selection, drag, and resize responsive while moving
CPU-heavy and memory-bounded work away from the UI event loop.

## Document interaction surfaces

Document commands use four interaction surfaces. The command owns the choice;
localized labels and viewport heuristics never decide behavior.

| Surface | Use | Placement and behavior |
| --- | --- | --- |
| Task pane | Find and replace, page setup, citation sources, and revision review | Mutually exclusive at the right edge of the document workspace. It shares width with the page above 900 px, overlays the page below 900 px, and becomes workspace-wide below 520 px. Escape and the visible close action follow the same guarded close path. |
| Review rail | Anchored comments and comment drafts | Sits beside the paper and connects each thread to its text range. Below 620 px it becomes a bounded review drawer, removes connector lines, and lays cards in document order instead of using absolute placement. |
| Anchored popover | Table size, paragraph spacing, paragraph pagination, colors, and select options | Portaled to the body but positioned from the invoking control. It flips vertically, clamps to a 16 px viewport margin, updates on nested scroll and resize, focuses its first field when it behaves like a small dialog, and returns focus to the trigger on Escape. A dirty numeric field consumes the first Escape to discard only its local draft; a clean field lets the next Escape close the surface without a blur commit. |
| Modal dialog | Captions, cross-references, links, image descriptions, notices, and confirmations | Centered in the viewport with bounded height and an independently scrolling body. The body-level portal makes every non-dialog body child inert. Focus stays inside, destructive confirmations initially focus the safe action, and closing returns to the actual invoker unless the completed command deliberately restores the editor. |

Task panes preserve editing context instead of behaving like navigation.
Closing a pane returns focus to the body editor. Switching away from a dirty
citation draft or an unsent comment reply requires an explicit discard
decision; cancelling that decision returns to the control that requested the
switch. Citation editing uses progressive disclosure: citation identity,
title, year, and authors remain visible while secondary publication metadata
stays collapsed until requested. An edited source cannot be inserted before it
is saved.

List galleries use a roving tab stop that follows Arrow, Home, and End focus.
Tab therefore exits from the visibly focused style rather than from a stale
previous option, and completed list commands return directly to document
editing.

Numeric ribbon and task-pane controls keep incomplete edits local. Escape
restores the last committed value, Enter produces exactly one command even
when that command synchronously returns focus to the editor, and IME
composition Enter is never treated as a commit. Table dimensions follow this
contract without changing the active cell selection. Spreadsheet and
presentation chart axes additionally reject minimum/maximum pairs that invert
the visible range and expose the invalid draft before it is committed.

Spreadsheet workbook task panes share one explicit saved-draft contract across
defined names, charts, conditional formats, calculation settings, pivot tables,
editable protection ranges, and print settings. Unrelated workbook updates do
not overwrite a dirty draft, switching to another managed object is blocked
until the current draft is saved or cancelled, and save actions stay disabled
when there is no change. The first Escape discards a dirty pane draft while
keeping the pane open; a second Escape closes the clean pane. Field-level
numeric drafts still consume Escape before the pane-level draft, so incomplete
input cannot trigger either a workbook mutation or an accidental close.

Shared editor context menus keep their pointer or keyboard-selection anchor on
desktop. Below 640 px they use a safe-area-aware bottom action sheet instead:
every action is at least 44 px tall, long command sets scroll inside a bounded
surface, and Escape retains the same exact editing-target focus restoration.
The responsive change is presentation-only; command ordering, shortcuts, and
host-provided actions remain identical.

Review actions distinguish reversible, local decisions from broad destructive
ones. Individual revision decisions remain direct commands. Accepting or
rejecting every revision requires confirmation, and an empty revision pane does
not retain disabled bulk controls. Deleting a comment confirms that its thread
and any unsent reply will also be removed.

Selected document text has a typed host-owned context-menu boundary.
`getSelectionMenuItems` receives an immutable snapshot containing the selected
plain text and structured fragment, bounded adjacent text, synchronized HTML,
the complete document text, and the current controlled `WorkDocumentContent`.
When a host item is selected, its callback receives conflict-aware copy,
replace, insert-before, and insert-after commands. The menu factory completely
replaces the built-in selection menu, so the editor does not infer host intent
from localized labels or keywords. While an async callback Promise remains
pending, its target range maps through unrelated transactions. An edit to the
selected text makes the target stale and rejects the later command rather than
changing an unrelated range. Successful edit commands honor tracked changes
and produce one controlled update and one undo record. Open-ended host actions
collect required user input before creating an agent request. The Playground's
question action therefore opens a focused composer, preserves the complete
selection context in controlled state, and dispatches only after an explicit
submit; prepared context is collapsed by default in the assistant UI.

Markdown exposes the same host-owned boundary for both source and visual
selections. Its snapshot identifies the active surface and includes the exact
selection, bounded adjacent text, complete Markdown, rendered plain text, and
the current `WorkMarkdownContent`. Source actions validate their exact string
range; visual actions map through unrelated ProseMirror transactions. Both
surfaces reject changed targets with `stale-selection`, and every successful
edit returns a new controlled content value through `onChange`. The ribbon
also routes formatting and insertion commands to the active source selection
instead of silently changing the visual editor.

The controlled Markdown textarea owns a bounded source-history stack rather
than relying on browser-native undo for a controlled value. Consecutive typing
is coalesced, explicit commands remain separate records, undo and redo restore
the source selection, and a host replacement rebases the stack. Source-to-
visual synchronization is marked as non-historical so one source intent does
not create a second visual-editor undo record.

These rules are implemented by `DocumentTaskPane`, `Popover`, `Dialog`, and
`useOfficeDialog`. Native browser prompts, confirms, and selects are not valid
fallbacks. All four surfaces must be verified at 1280 px, 768 px, 520 px, and
390 px widths for viewport containment, keyboard entry, Escape behavior, and
focus restoration.

## Current implementation status

| Product | Implemented browser surface | Implemented kernel boundary | Next fidelity gate |
| --- | --- | --- | --- |
| Document | One TipTap/ProseMirror body tree retained across editing and read-only preview, controlled TipTap header/footer surfaces with direct paper-margin editing and a contextual ribbon, one typography and page-chrome baseline across editing, preview, and PDF surfaces, a persistent typed heading-navigation pane with active-selection tracking, filtering and collapsible keyboard traversal, host-owned selected-text menus with full document context and range-mapped async commands, responsive and keyboard-operated paragraph-style and list galleries, typed bullet and numbering commands with restart/continue/start controls, typed physical-page and section-page descriptors, repeated first/default/even page chrome, a versioned structured model with an HTML compatibility representation, prefix-reused visual-line measurement and pages, page decorations, page-aware horizontal and vertical rulers for page margins, paragraph indents and typed tab stops, structured list-item pagination, explicit paragraph and list-item direction, compact spacing and pagination controls, typed inline/square/top-and-bottom image layout, imported style-inherited paragraph properties, structured inline tabs, and theme-aware run font/size/color/background import | Worker plus resumable Rust/WASM flow pagination and Rustybuzz shaping across CSS-matched registered text runs; the same live result remains mounted in read-only preview and is consumed by bounded browser PDF capture, including eligible list paragraphs, Unicode bidi level segmentation, ordered per-grapheme font fallback, packaged Latin/CJK/Arabic/Hebrew faces, and structured left-to-right tabs, with explicit DOM and JavaScript fallbacks for text affected by supported floats | Language-complete font substitution, complete Word style and numbering coverage, locale-complete and bidirectional tabs, arbitrary floating-object offsets and layering, complex table flow, searchable/vector PDF output, and loss-preserving OOXML package state |
| Markdown | TipTap visual editing with a resizable source-and-preview split view by default, source-aware ribbon formatting and shortcuts, bounded source-native history with coalesced typing and selection restoration, host-defined selection menus across both editing surfaces, GFM tables, strikethrough, autolinks and nested task lists, controlled source state, coalesced preview rebuilds, proportional pane scrolling, keyboard-adjustable pane sizing, a flat measure-bounded reading surface, optional visual or source-only views, and a full-workspace Source/Preview switch on phones | No kernel required for normal editing | CommonMark differential fixtures, multi-megabyte profiling, and an off-main-thread parser boundary when measurements justify it |
| Spreadsheet | Fortune Sheet grid integrated with the shared Office shell, shared grid/formula/footer visual tokens and one spreadsheet accent across selection, sheet tabs, menus, and zoom, an A3S-owned accessible single-row workbook footer for creation, activation, rename, duplicate, color, hide, reorder, delete, selection status, and zoom, a resize-observed worksheet viewport that keeps the active tab and its touch action visible across desktop, compact, and phone-width changes, one cell/worksheet context-menu surface whose displayed accelerators remain executable while the menu owns focus, Arrow/Home/End tab navigation plus Shift+F10 and native context-menu access, typed editing and calculation command ports, A3S-owned deterministic Arrow, Enter, Tab, Home, PageUp/PageDown, extended-selection, row, column, and all-cells keyboard commands, selection-preserving focus restoration across controlled workbook remounts and F2/Escape editing transitions, Cmd/Ctrl formatting, undo/redo, clear, a permission-resilient editor-local clipboard fallback, Shift+F11 sheet creation, and Ctrl/Cmd+PageUp/PageDown sheet-navigation shortcuts, direct font-family, horizontal/vertical-alignment, text-wrap, general/number/percent, and decimal-place controls backed by native cell-style keys, live count/sum/average selection summaries, operation-driven sparse-workbook projection, guarded controlled-value remounts that reject stale engine callbacks, and no-history result patches with cell-scoped Fortune fallback | Versioned, cancellable Worker/Rust-WASM calculation sessions using the shared bounded Rust formula parser, retained formula ASTs, incremental forward/reverse dependency graphs, dirty-subgraph recalculation, cross-sheet references, and a dynamically loaded JavaScript fallback | A3S-owned virtual grid, moving replacement projection off the main thread, broader Excel formula semantics, A3S-owned custom number-format evaluation, and print layout |
| Presentation | Scene canvas with ordered typed multi-selection, persistent nested browser groups, native PPTX group-node export, exact keyboard-accessible table-dimension insertion, native slide/object context actions with optional AI actions, a separate object/content editing state, one on-demand TipTap instance, collective move/scale/nudge/clipboard/delete/layer commands, selection-bound alignment and distribution, typed group/ungroup commands, one typed dispatcher for ribbon commands, editor-scoped shortcuts with post-command selection focus, direct beginning/current-slide playback with fullscreen fallback, frame-coalesced transactional move/resize previews that commit once on pointer release, two-level thumbnail node and scene windowing, and a phone stage that top-aligns the primary canvas while retaining one readable slide indicator beside view and zoom controls | Revisioned, cancellable Worker/Rust-WASM slide-relative alignment and object-set snapping with typed visual guides and a JavaScript fallback | Arbitrary rotated or reflected PPTX group transforms, connectors, theme resolution, text fitting, kernel-owned thumbnail layout, and slide serialization |
| PDF | PDFium-backed page rendering with an A3S-owned responsive toolbar, a scrollable page-thumbnail rail with active-page synchronization and bounded bitmap/DOM windowing, a focus-contained phone page drawer, and typed capability controllers for navigation, zoom, search, basic annotations, annotation color, opacity, compatible stroke-width defaults and selection updates, history, and save; page and search drafts cancel without accidental commands, shortcuts remain scoped to the PDF root, and compact widths retain page status while exposing search-result traversal, navigation, zoom, and history through the keyboard-operated overflow menu | PDFium WebAssembly | Forms, redaction review, page organization, and reopen fixtures |

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

Document editing, read-only preview, and browser PDF composition resolve their
base font, size, line height, paragraph spacing, headings, lists, quotations,
and image wrapping from the same CSS rendering variables. Preview and PDF page
chrome is absolutely positioned inside the physical page margins, matching the
editing surface instead of consuming body flow. An absent header produces no
header element, and PDF composition never invents a filename header. Imported
inline formatting and host typography variables remain higher-priority inputs.
Narrow preview containers retain the physical page width and margins and
scroll the page instead of silently shrinking the text area and changing line
breaks.

Editing and read-only preview mount the same TipTap document surface. Switching
mode changes editability, accessible role, command chrome, and contextual
interactions without replacing the ProseMirror DOM or clearing its measured
blocks, shaped runs, page descriptors, and automatic-break decorations. Page
headers and footers are visual overlays inside the configured top and bottom
margins. Kernel protocol 15 therefore computes body height from page height
minus physical margins only; page-chrome heights position repeated overlays
and do not subtract the same space again. Browser PDF export locates the ready
surface through the host's stable artifact ID, clones its ProseMirror DOM once,
removes editing-only state, and crops exact physical pages from bounded capture
batches. Explicit descriptor pages remain a fallback for surfaces without a
registered live document.

The Word navigation pane uses that same capture boundary for real physical-page
thumbnails instead of reconstructing an approximate text card. The current and
adjacent pages are admitted first; the remaining captures use an
`IntersectionObserver` window, one serialized queue, debounced source-mutation
refresh, and off-screen image release. Each capture clones the already measured
live page, strips editing state, and crops one exact page offset. It deliberately
does not wait for the document-wide font set or a browser animation frame:
pagination readiness is the layout authority, while unrelated fonts and
background agent tabs must not stall a preview. Text excerpts remain a bounded
loading or failure fallback. Above 48 pages, the navigation pane mounts a
contiguous window of at most 24 page buttons, plus the selected and roving pages
when either falls outside that window. Variable-height spacers preserve the
physical scroll range, while Home and End materialize and focus their sparse
destinations without relying on a browser animation frame. A deterministic
120-page DOCX A3S Test proves bounded mounting, first/last-page keyboard access,
selection, focus retention, and spacer geometry. Raster work and page-button
DOM are therefore bounded.

The page list, heading outline, and full-text results reuse one navigation
window model. Heading and result collections switch to a window above 48 rows,
mount at most 32 contiguous rows, and add only sparse active, selected, or
keyboard-roving rows outside that range. Physical before, between, and after
spacers retain native scroll geometry; global `aria-posinset` and `aria-setsize`
values remain truthful even when most rows are not mounted. Home and End mount
their destination without an animation-frame dependency. A real 120-page DOCX
with 120 built-in headings and 120 text matches proves bounded outline and
result mounting, first/last keyboard access, exact result selection, and an
error-free browser run. Programmatic long-distance selection temporarily uses
instant document scrolling through the next paint, then restores the surface's
normal smooth-scroll style so the selected result is visible in the same frame.

Revision review reuses this window model above 48 tracked changes. It mounts at
most 32 contiguous revision cards, retains the keyboard-roving card as a sparse
pin, and represents every omitted range with a physical spacer. Arrow,
PageUp/PageDown, Home, and End can therefore reach the complete queue without
mounting it. Accepting or rejecting a revision transfers focus to the same
decision on the adjacent surviving item. A deterministic A3S Test imports a
real DOCX with 120 native OOXML insertions and proves bounded mounting,
first/last focus, the 120-to-119 decision transition, spacer geometry, and zero
console or page errors.

Spreadsheet now uses a persistent browser Rust/WASM calculation session. The
editor initializes it with a sparse workbook replacement, sends bounded cell
patches from stable Fortune cell operations, and requests only the dirty
dependency subgraph during automatic calculation. One-cell edits read only
their changed coordinates; row, column, sheet, broad data, pivot, and
unrecognized operations take the checked replacement path. Manual mode keeps
an independent current projection while the last submitted session remains an
immutable diff baseline. No-history calculation result patches are consumed as
projection updates without returning formula caches as user input or creating
a separate undo step. Fortune does not apply later `data` props to a populated
workbook, so external, history, panel, and agent-controlled values use a
guarded remount; a value already emitted by the live workbook does not remount
or lose its interaction state.

Rust retains parsed formula ASTs plus forward and reverse dependency edges;
references to blank cells, dependency rewiring, unresolved formulas, partial
target calculation, cancellation, and stale patch revisions have explicit
tests. Successful scalar results are applied without adding undo history,
known grouped formulas refresh before their dependents, and unresolved
dependencies enter an ordered, cell-scoped Fortune Sheet compatibility pass.
The shared Rust parser handles the same bounded formula grammar in the native
core and browser kernel. This is not a complete Excel engine: Fortune Sheet
remains the canonical grid, initial and replacement sparse projection still
run on the main thread, and the kernel does not materialize whole-row or
whole-column ranges, calculate arrays, spills, structured references or
external workbooks, own number formatting, or own print layout.
Presentation sends alignment plus move and resize snapping to Rust/WASM. The
main thread treats an ordered selection as one bounded geometry frame, paints
at most one transient preview per animation frame, and ignores stale geometry
responses. The Worker returns snapped slide-relative geometry and at most one
typed guide per axis. Pointer movement never mutates the controlled document;
pointer release emits one host change for every moved object, while
cancellation emits none. Selection-only clicks mount no editor, Shift-click
toggles membership, and double-click or Enter opens content editing.

Presentation elements may carry an outermost-first `groupIds` path. The
top-level path segment defines the current logical selection unit: selecting or
dragging one member selects or moves every member, while alignment,
distribution, layer movement, clipboard operations, deletion, and history
operate on the complete unit. Group and ungroup commands add or remove one path
level in one controlled update. Copying objects, slides, or layouts remaps every
group ID without linking the copy to its source. PPTX import retains nested
group paths after mapping child geometry into slide coordinates and applies the
smaller cumulative group axis to typography, explicit rich-text run sizes, and
border weights. PPTX export resolves exact generated objects through temporary
names, removes those names, and builds nested native group nodes for slides,
layouts, and master-derived artwork. The generated group coordinate maps are
identity transforms based on emitted OOXML bounds, preserving hierarchy and
visual geometry after non-rotated, non-reflected transform round trips.
Arbitrary group rotation or reflection, connector routing, text fitting, theme
resolution, and kernel-owned thumbnail layout remain later fidelity gates.

Spreadsheet, Presentation, and PDF commands cross explicit typed boundaries.
The shell never searches visible labels, scrapes rendered text, or synthesizes
clicks to infer product intent. This keeps localized UI copy independent from
behavior and gives Worker/WASM operations a stable request contract.

## Markdown editing flow

```text
Controlled Markdown source
            |
 immediate source textarea + host onChange
            |
 trailing 160 ms coalescing boundary
            |
 TipTap Markdown parser and visual tree
            |
 GFM table / task / link / strike surface
```

Markdown source remains the persistence boundary. Source keystrokes update the
controlled value immediately, while the visual tree consumes only the latest
queued value after a short trailing delay. Moving focus into the visual pane or
switching from source-only mode flushes the newest source first, so the user can
never edit a stale preview. Visual edits serialize back to Markdown and cancel
any pending source parse.

The split panes synchronize by normalized scroll progress rather than raw
pixels because source lines and rendered blocks have different heights. Status
counts use React deferred work so code-point counting does not compete with
high-priority typing. GFM compatibility fixtures cover tables, strikethrough,
autolinks, task-state round trips, controlled host replacement, and coalesced
source updates. Parsing remains on the main thread for now; moving it to a
Worker is a measured follow-up for multi-megabyte inputs, not an assumed
abstraction.

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

Document section nodes are isolating editing boundaries. Normal paragraph
breaks, empty paragraphs, and multiline paste therefore remain inside the
active section instead of becoming unmeasured root blocks. Explicit section
commands still split or merge those boundaries transactionally.

Page view renders one non-editable physical paper sheet per measured kernel
page behind that single editor tree. The mapped boundary widgets reserve the
remaining printable height, page chrome, and page gap while the sheet stack
provides separate page backgrounds, borders, and shadows. The visible page
count therefore cannot degrade into one continuously stretched paper surface,
and changing the page color still applies to every physical sheet without
creating additional editable roots.

The current browser-kernel slice collects contiguous geometry-affecting text
runs from eligible paragraphs. Each run carries an ordered registered font
stack, size, line height, letter spacing, ligature, and kerning behavior. Font
families and normal/italic styles remain exact; numeric weights follow the CSS
Fonts matching order so WPS-style values such as 680 or 730 select a registered
700 face. When a family has only one normal-style weight, the same face
provides deterministic metrics for browser-synthesized bold text. The kernel
selects a face for each grapheme, joins adjacent selections that use the same
face, and includes every used face in line ascent and descent.
Rustybuzz shapes those segments and the kernel applies Unicode and
grapheme-safe line breaking across run boundaries. Unsupported paragraphs
retain the existing DOM range path, which maps each browser visual-line start
back to a ProseMirror position. This fallback also fragments paragraphs with
explicit hard breaks so gradual typing and multiline paste cannot overflow a
physical sheet as one atomic block. The Worker/WASM kernel chooses page
breaks, keeps a minimum number of line fragments on either side of a break,
and returns mapped decorations.
Tables, images, code blocks, and other complex content use explicit
format-specific measurement. Top-level tables are row flows; eligible rows can
additionally expose synchronized direct-cell block fragments. Ordered and bullet
lists recurse through semantic list items, expose each direct item block at its
stable ProseMirror position, and distribute list-container height without
double-counting nested content.

Document image nodes carry typed layout, alignment, wrap-distance, alternative
text, size, and per-image aspect-ratio-lock attributes. The contextual Picture
ribbon keeps fast layout actions and opens one responsive Picture Properties
dialog for centimeter width and height, aspect-ratio locking, wrapping,
alignment, text distance, and alternative text. Applying the complete draft
updates only changed attributes in one separated TipTap history entry; cancel
and Escape preserve the node and return focus to the exact invoker. Dimensions
that were only projected into rounded centimeter fields retain their exact
imported pixel values when untouched. Inline images stay in normal flow;
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
history. Empty outer-paper margins render no persistent placeholder, border,
or activation control. A double click anywhere in the header or footer margin,
or the keyboard-accessible Header and Footer commands in the Insert ribbon,
replaces the static HTML with one toolbar-free TipTap surface in place and adds
a contextual Header and Footer ribbon. Existing imported content remains
visible in the clean state without an edit label. The contextual ribbon uses
the same typed command functions for formatting, alignment, links, images,
page-number visibility, switching between header and footer, and closing the
mode. Escape and web-view transitions close the mode and restore body focus.
The body is visually de-emphasized without being copied or removed from the
canonical tree. Both surfaces preserve semantic paragraphs, lists, tables,
links, marks, alignment, color, and inline raster images while exposing a
smaller command set than the document body. Toolbar controls derive active
state from the current selection; they do not inspect localized labels or use
`document.execCommand`. The controlled HTML values still pass through the
page-chrome sanitizer and the default/first/even-page model, so direct paper
editing does not create a second persistence format or bypass DOCX
header/footer relationships.

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

The same safe prefix is also the cache boundary for page descriptors. Stable
descriptors retain their resolved first/default/even header and footer content,
page number, preview text, and navigation target; only the rewound boundary and
new suffix are derived again. Runtime diagnostics expose reused and derived
page-chrome counts so long-document browser tests can enforce the incremental
budget instead of inferring it from elapsed time.

It also supports explicit page and section breaks, keep-together blocks,
keep-with-next headings, cancellation, stale-revision rejection, and automatic
reflow after editing, resizing, font loading, or zoom changes. Page view uses
the result for visual paper gaps and status-bar page counts. Its horizontal
ruler is bound to the active section's paper width and left/right margins and
to the active paragraph's left, right, first-line, and hanging-indent
attributes. A vertical ruler controls the active section's top and bottom
margins. Pointer and keyboard changes update the same section node used by
pagination; compact layouts hide the vertical ruler while retaining Page
Layout controls. The Page Setup task pane separates Page, Columns and Sections,
and Header and Footer into keyboard-operated tabs. It opens on paper and margin
settings, keeps only one settings group mounted, and defers the two bounded
TipTap page-chrome editors until the user explicitly requests them. The
paragraph attributes map to and from `w:ind` formatting during DOCX import and
export. Paragraph spacing keeps before/after values plus Word's `auto`, `exact`,
and `atLeast` line rules. Direct `keepLines`,
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
import resolves `w:jc` alignment through the same defaults, style inheritance,
and direct overrides. Physical left/right values remain physical, logical
start/end values follow the resolved paragraph direction, and Word's supported
distributed variants map to editable justification. Aligned list paragraphs
remain semantic list items with an aligned paragraph child. Indents, spacing,
and pagination use the same style chain. One shared style index serves these
paragraph import passes. Marker application scans converted text once per
property family instead of once per paragraph. Web view removes both automatic
pagination and the paper ruler.

Run-format import uses the same style index. It merges document run defaults,
paragraph-style run properties, direct paragraph run properties, bounded
character-style `basedOn` chains, contextual table-style properties, and direct
run properties. Unique sentinels survive Mammoth conversion and become
ordinary inline semantics before the TipTap/ProseMirror model is created. This
currently preserves visible bold, italic, underline, strikethrough, font
families, half-point sizes, direct hexadecimal colors, Word highlights, and run
shading. The importer also reads the DOCX font and color schemes, resolves
major/minor theme font references, system or sRGB theme colors, and tint/shade
transforms. Superscript and subscript remain structured, mutually exclusive
TipTap marks; their ribbon and keyboard commands switch vertical position
without stacking both marks, and DOCX round trips retain the corresponding
`w:vertAlign` values. Export maps background colors back to OOXML run shading.

Lists remain semantic ProseMirror structures. Pagination measures each direct
list-item block independently and maps a page widget before the first block
inside the target `li`; eligible list paragraphs are collected for the same
Rustybuzz text-layout request as ordinary paragraphs. DOCX export emits one
paragraph with `w:numPr` per item, supports levels zero through eight, positive
ordered-list starts, decimal, letter, and Roman formats, three canonical bullet
shapes, nesting, and RTL paragraph/run properties. Table cells, notes, headers,
and footers use the same export path. DOCX import uses bounded numbering
metadata to restore starts, common ordered formats, and disc/circle/square
bullets after Mammoth conversion, while direction markers are applied to the
semantic list item. Typed commands apply an active style idempotently, restart
at one, set a positive start, or continue the preceding same-depth list by
adopting its format and calculated next value. The split-button galleries keep
quick toggles separate from keyboard-operated style selection and local
numbering settings. Reversed lists, per-item `value` overrides, native Word
list-identity continuation, arbitrary multilevel templates, and loss-preserving
custom numbering formats remain later fidelity gates.

The declared direct-formatting layout slice now has a real WPS Writer reference
gate. Its deterministic A4 fixture uses installed Arial runs, explicit OOXML
paragraph spacing, a fixed centered table, direct cell fills and borders, and
physical cell margins. Automatic `w:line` values retain their original OOXML
multiples for export while the browser applies the measured WPS single-line
font metric; imported tables no longer receive editor-only block margins. The
Windows parity workflow exports the same fixture through WPS, captures both
794 by 1123 CSS-pixel pages, and rejects page-size, landmark, or bounded pixel
differences. Language-complete font substitution, variable font axes, the
remaining character and table style properties, arbitrary floating-object
offsets, cropping, contour wrapping and layering, row-internal table splitting
inside a single long paragraph or complex merged-cell flow, nested tables,
footnote balancing, multi-column flow, and mixed-size sections require the
later layout stages below.

Top-level tables now enter pagination as row flows. DOM measurement supplies
row heights and stable ProseMirror positions; the Worker/Rust-WASM kernel keeps
leading header rows with the first body row and reserves the measured header
height on continuation pages. A table-aware widget creates the paper gap and
an `aria-hidden`, non-editable copy of the rendered header with the original
column group. The editable row remains the only canonical content. TipTap row
attributes preserve `cantSplit` and `tblHeader` through structured HTML and
DOCX import/export. Eligible rows, including rows taller than a complete body
page, split at synchronized boundaries between direct cell blocks. One
non-editable widget per cell aligns the page gap and paints a clipped slice of
the same measured header overlay. The leading cell alone extends the paper-gap
paint across the page margins. `cantSplit` rows and a single indivisible long
paragraph remain atomic and may overflow; nested tables and complex merged-cell
flows require a later fragmentation model.

Table creation is selection-safe: a non-empty text selection is preserved and
the chosen table is inserted after its containing block instead of replacing
the selected content. The insert ribbon exposes a keyboard-navigable 8 by 10
size picker. Entering a table opens separate Design and Layout contextual tabs.
Design owns keyboard-operated style presets, the header row, multi-cell
shading, and a reusable border pen with whole-selection, outside, inside, side,
inside-horizontal, and inside-vertical targets. Layout owns row and column
insertion or deletion, cell merge and split, horizontal and vertical alignment, repeated headers,
atomic rows, centimeter-based column width and row height fields, equal row and
column distribution, content/window autofit, table alignment, four-edge cell
margins, and table deletion. Physical
column widths remain coherent across merged cells through ProseMirror's table
map. The layout algorithm, auto/percentage/pixel preferred width, alignment,
indent, table-level cell margins, cell-level margin overrides, and explicit row
heights are independent typed attributes rather than transient DOM
measurements. Edit, preview, print, and DOCX import/export share that geometry.
Table, row, column, and cell properties are edited through one four-tab Table
Properties dialog instead of disconnected immediate commands. The dialog owns
preferred width, placement, indent, selected-row height and pagination,
current-column width, selected-cell vertical alignment, and cell margins. It
reads imported geometry without approximation, validates the complete draft,
and dispatches one TipTap transaction and one undo record. Draft fields are
display projections only: unchanged pixel values and partial per-cell margin
inheritance are copied from the typed source instead of being quantized through
the two-decimal centimeter UI. Cancel and Escape leave the model untouched and
restore the exact ribbon invoker. At compact widths the four tabs, choices,
steppers, and commit targets remain inside the viewport with touch-sized hit
areas.
DOCX tables independently round-trip `tblGrid`, `tcW`, `tblLayout`, `tblW`,
`jc`, `tblInd`, `tblCellMar`, `tcMar`, and `trHeight`; numeric fiftieth-percent
and string percentage widths keep their actual size instead of collapsing to a
full-width table. Window/content autofit remains responsive in the browser.
Independent top, right, bottom, and left cell-border attributes
render identically in edit and preview, export as independent `w:tcBorders`,
and reopen without being flattened. Explicit table-level outer and inside
borders resolve onto their owning cell edges during import. Direct
`themeColor`/`themeFill` values and their tint or shade transforms resolve from
the DOCX package theme into stable edit, preview, and export RGB values. A
default or referenced `w:style w:type="table"` resolves through a bounded,
cycle-safe `w:basedOn` chain. `w:tblLook` flags or bitmasks, row and column band
sizes, `gridSpan`, and `gridBefore`/`gridAfter` select `wholeTable`, row/column
bands, first/last rows or columns, and corner-cell conditions in Word precedence
order. Their cell fills, per-edge borders, run emphasis, fonts, and colors, plus
paragraph alignment, direction, indents, spacing and line rules, pagination
rules, and tab stops enter the marker pipeline after paragraph styles and before
direct paragraph, character-style, or run formatting. Direct table and cell
properties remain higher-priority layers for cell presentation. Theme-derived
conditional presentation is materialized as stable RGB for editing, preview,
and regenerated DOCX output.
The real styled-table fixture exercises a centered 62.5% table, table margins,
a first-cell margin override, and matching edit/preview geometry through A3S
Test.
Command availability comes from the ProseMirror table state, so actions that
cannot apply to the current selection remain disabled. Loss-preserving semantic
theme references, less-common conditional paragraph properties outside the
supported layout set, percentage-width column authoring, nested-table editing,
and complex merged-cell conditional flow remain explicit fidelity gaps.

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
- `office_kernel_spreadsheet_calculation`
- `office_kernel_spreadsheet_session_calculation`
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
coordinates, and a tagged operation for slide alignment or object snapping.
Snap operations identify the moving element, transform mode, and independent
horizontal and vertical thresholds derived from a six-pixel visual distance.
Results return the resolved geometry plus typed slide or element guides. The
Worker suppresses cancelled responses, and the React integration accepts only
the newest request in the active gesture. Matching JavaScript implementations
preserve editing if Worker or WebAssembly loading is unavailable. The
JavaScript text fallback explicitly reports unsupported paragraphs so the
editor uses DOM line measurement instead of estimated font metrics.

Spreadsheet requests carry populated cells only, zero-based coordinates,
cached scalar values, formulas, and optional calculation targets. The kernel
recursively resolves dependencies in deterministic order and returns successful
cells separately from cell-scoped issues. Its first scalar function set covers
common arithmetic, comparison, concatenation, aggregation, logical, and numeric
operations. Persistent requests begin with `replace`, then use revisioned
`patch` updates plus `dirty`, `targets`, or `workbook` calculation scopes.
Formula ASTs and dependency edges survive between requests, while patches
rewire only changed formula nodes and mark their transitive dependents dirty.
Cancelled requests still execute in the serialized Worker queue so the next
patch always observes its declared base revision; only their responses are
discarded. The JavaScript fallback loads its parser only when needed and
recalculates from the supplied full sparse snapshot. Unsupported formula
structures remain unchanged and conservatively dirty so a host never loses the
source formula or cached value.
Rust/WASM remains the canonical calculation path. If Worker or WebAssembly
loading fails, the Fortune-based fallback keeps the workbook editable but may
use Fortune coercion and eager-branch evaluation for formulas beyond the
shared parity fixtures; this is an explicit first-slice compatibility limit,
not a cross-engine equivalence claim.

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

The document font selector groups packaged layout faces, common Chinese,
Western, and monospace system families, host-provided layout faces, and an
otherwise unknown family retained from imported content. Each choice previews
its own CSS family while short section headings distinguish packaged, system,
host, and imported entries without repeating status text on every row. The same
grouped catalog is reused by the compact selection toolbar, Presentation text
controls, and Spreadsheet cell formatting.

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
controlled value. Document-level page color remains outside the ProseMirror
tree, is updated through the same controlled value, and is preserved by DOCX
import/export and the browser PDF render surface.

This is a transitional persistence boundary, not loss-preserving OOXML yet.
DOCX still converts through Mammoth, but DOCX, HTML, and text imports create the
structured model immediately through the same extension schema used by the
editor. DOCX export materializes synchronized HTML from the structured model
before OOXML generation, so a stale HTML cache cannot override model changes.
Preview retains the mounted structured editor tree, and live browser PDF export
captures that same paginated surface; compatibility HTML remains an
import/export boundary. Unsupported OOXML package parts are not yet represented
by the model.

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
  states, shortcuts, and responsive rules in every editor. Ctrl/Cmd + wheel is
  captured inside the active editor and routed to the same bounded zoom model
  as its status controls instead of changing browser page scale.
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
- Complete row-internal splitting for single indivisible paragraphs and
  complex merged-cell flows, and add nested tables, full floating-object
  geometry beyond the supported square/top-and-bottom image anchors, footnote
  balancing, columns, and mixed page sections. Row-flow pagination, direct-cell
  block splitting, repeating headers, and the initial image-wrapping slice are
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

The first calculation slice is implemented: a sparse calculation projection,
shared bounded Rust formula grammar, Worker/WASM scalar dependency evaluation,
revision cancellation, stale-result rejection, target-only recalculation, and
JavaScript and cell-scoped Fortune fallbacks. Persistent Worker/WASM sessions
now retain parsed formulas and bounded forward/reverse dependency graphs;
stable workbook edits use cell patches and automatic calculation evaluates only
the affected transitive formula subgraph. Stable Fortune cell operations now
produce those patches directly without a dense workbook scan. Structural and
external controlled changes use a replacement projection, and controlled
values that did not originate from the live grid remount the Fortune surface
so the visible workbook cannot diverge from the host value. The canonical grid
is still Fortune Sheet. The A3S-owned virtual viewport, moving initial and
replacement projection off the main thread, complete formula semantics, kernel
number formatting, and print pagination remain open work in this stage.

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

The first transform slice is implemented. Move and resize gestures keep their
starting object immutable, coalesce pointer events into a transient preview,
and submit only the latest candidate to the cancellable geometry Worker. The
Rust/WASM kernel resolves slide-edge, slide-center, and sibling-object anchors;
the canvas paints the returned horizontal or vertical guide without adding
text or persistent scene objects. Pointer release submits one controlled
content value, so the complete gesture occupies one undo step. Pointer cancel,
selection-only clicks, stale responses, and superseded gestures do not create
history. A short timer fallback preserves progress when a browser temporarily
suspends animation frames.

The object-selection slice is also implemented. Selection is an ordered set
owned by the Presentation editor rather than a DOM class or visible-label
lookup. Shift-click toggles membership; a plain canvas click clears it;
double-click or Enter opens content editing; and Escape returns focus to object
mode. Multi-object move, arrow-key nudge, copy, cut, paste, duplicate, delete,
layer movement, alignment, and distribution each publish one controlled value
and therefore one undo record.

The persistent-grouping slice is implemented. Nested, outermost-first group
paths survive controlled browser edits and PPTX group import. A top-level group
is one selection and arrangement unit; group and ungroup commands,
`Ctrl`/`Command` + `G`, and `Ctrl`/`Command` + `Shift` + `G` each publish one
controlled value. Object, slide, and layout copies receive independent group
IDs. A shared bottom-right selection handle scales every selected member
against the immutable selection frame. Geometry uses independent horizontal
and vertical factors; typography, explicit rich-text run sizes, and border
weights use the smaller factor so a non-uniform scale does not force text
outside its child bounds. Preview work remains frame-coalesced and pointer
release publishes one controlled value, so one undo restores the complete
group. PPTX export emits nested native groups without matching objects by
position, removes every temporary object marker, and keeps slide, layout, and
master-derived group scopes separate. Import normalizes supported group scale
into child geometry and visual metrics. Arbitrary group rotation and
reflection, source group names, and kernel-owned group geometry remain
separate fidelity gates.

The thumbnail-virtualization slices are implemented. Short decks retain every
thumbnail button. Decks above the bounded threshold render a contiguous row
window plus measured virtual spacers in both the normal strip and sorter, so
React and DOM node counts depend on the viewport rather than the slide count.
Keyboard commands use absolute slide indexes instead of querying mounted
siblings; `Home`, `End`, arrow navigation, and deletion can therefore cross an
unmounted range. Each mounted button exposes its absolute position and total in
its accessible name, and an externally selected slide is revealed without
stealing focus.

The mounted window shares one root-scoped intersection observer. Full inherited
layouts, images, tables, charts, and text scenes mount only for the selected
slide and slides inside a bounded viewport overscan, and unmount after leaving
it. A mutation observer reconnects scene observation as thumbnail windows
change. Environments without `IntersectionObserver` render complete scenes for
the bounded mounted window. Kernel-owned thumbnail layout remains a separate
performance gate.

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
annotation tools, annotation color, opacity, and compatible stroke-width
defaults and selected-annotation updates, annotation deletion, history, and
copy export through public plugin capabilities. It contains no shadow-DOM
queries, private viewer selectors, label inference, or synthetic clicks.
Page-number Escape cancels without blur submission, pending searches cannot
replay stale results, and responsive overflow actions remain reachable whenever
their direct toolbar controls are hidden. The product-owned thumbnail rail uses
PDFium thumbnail capabilities without enabling the embedded sidebar, keeps page
numbers and active-page state synchronized with the main scroll controller,
automatically reveals the current page, and mounts only a bounded page window
for long files. At phone widths the same navigation becomes a focus-contained,
dismissible drawer rather than reducing the document canvas.
Existing PDF password handling remains in the PDF document lifecycle rather
than the removed toolbar. Form-authoring controls, redaction review, page
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
- Presentation pointer movement schedules at most one preview request per
  animation frame, never calls the host `onChange`, cancels superseded geometry
  work, and commits once when the pointer is released.
- The Playground initial JavaScript budget is 220 KiB gzip and is enforced in
  CI against the scripts referenced by the generated `index.html`.
- One active layout request exists per editor. A newer revision cancels the
  previous request.
- Requests reject invalid dimensions and more than 10,000 layout blocks.
- Text requests are bounded to 1,024 paragraphs and 1 MiB of UTF-8 text.
  They accept at most 16,384 contiguous runs, with no more than 4,096 in one
  paragraph and no more than eight ordered faces per run. Font registration is
  bounded to 16 faces and 32 MiB per face.
- Spreadsheet calculation requests are bounded to 1,024 sheets, 100,000
  populated cells, 100,000 targets, 8,192 Unicode characters per formula,
  100,000 materialized range cells cumulatively per formula, 64 nested formula
  dependencies, 100,000 patch cells, 1,000,000 dependency edges, Excel's
  XFD1048576 coordinate boundary, and 32,767 UTF-8 bytes per text value. The
  editor replaces rather than diffs a session when one controlled update
  changes more than 10,000 populated cells.
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
   grapheme-safe emergency wrapping, whitespace modes, shared formula parsing,
   scalar dependency calculation, cycles, targets, and validation.
2. JavaScript fallback tests for protocol parity, safe page-prefix reuse,
   no-Worker operation, cancellation, sparse Spreadsheet calculation,
   revisioned session fallback, and the shared Spreadsheet parity fixtures.
3. A raw generated-WASM ABI smoke test that registers both shipped fonts,
   proves the Latin face lacks CJK glyphs, resolves them through the ordered
   fallback face, verifies mixed-face line metrics, initializes a Spreadsheet
   session, and recalculates a dirty formula chain from a cell patch.
4. Browser checks for real Worker/WASM/font loading, shaped-line parity with
   browser line boxes at non-100% zoom, real per-grapheme fallback diagnostics,
   explicit unresolved-glyph fallback, page-view reflow, web-view clearing,
   page counts, one physical sheet per measured page, visible page gaps, nested
   and RTL list flow, undo behavior, and slide-relative element alignment.

Presentation group serialization tests inspect generated slide and layout
OOXML, nested group order, identity child-coordinate transforms, unique
non-visual IDs, marker removal, master/layout scope isolation, grouped
placeholder materialization, cumulative visual scaling, and unsupported
rotation/reflection diagnostics.

Presentation thumbnail tests use a controlled intersection observer to prove
that long decks retain absolute keyboard reachability while bounding mounted
buttons and full scenes, releasing scenes that leave the window, preserving
focus after deletion, and reconnecting the observer when switching between
normal and sorter views.

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
