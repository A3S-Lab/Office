# Browser editor architecture

A3S Office uses an editor-specific interaction surface above a shared,
browser-safe Office kernel. TipTap is the right document engine for flowing
rich text, but it is not the canonical interaction model for a spreadsheet
grid, slide canvas, or PDF page.

## Product architecture

| Product | Interaction surface | TipTap responsibility | Kernel responsibility |
| --- | --- | --- | --- |
| Document | TipTap and ProseMirror | Canonical structured model plus an equal-position lazy live tree for eligible giant DOCX files; selection, commands, history, comments, and collaboration boundary | Page layout, style and font resolution, OOXML semantics, and serialization |
| Spreadsheet | Virtualized grid and canvas | Rich text inside a cell or floating text object | Formula calculation, workbook semantics, print layout, OOXML, and serialization |
| Presentation | Scene graph and slide canvas | Text inside an individual text box | Masters, layouts, themes, object geometry, OOXML, and serialization |
| PDF | PDFium page surface | None | PDF parsing, rendering, annotation serialization, and document save |

This separation prevents an editor framework from becoming a false abstraction
over products with different selection, layout, and performance requirements.

TipTap and ProseMirror are not competing choices in this stack. TipTap owns the
extension, command, schema, React integration, and public customization layer;
ProseMirror is the model and view engine underneath it. Ordinary document
features stay at the TipTap layer. Position-sensitive infrastructure such as
pagination decorations, transaction mapping, collaboration bindings, and
large-document NodeViews uses ProseMirror primitives through TipTap extensions
because those primitives expose the exact model positions and view lifecycle
that the higher-level API intentionally abstracts.

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
   returns to the compact row. Writer routes its catalogued Traditional Office font-size,
   alignment, line-spacing, heading, spelling, field-refresh, comment, and
   track-changes shortcuts only while the document surface owns the event. The
   status route uses the same catalog for Traditional Office word count, opening live page,
   word, character, and paragraph details from either `Ctrl+Shift+G` or the
   visible count. Its local formatting clipboard powers Traditional Office copy-format and
   paste-format shortcuts
   plus a one-shot format painter without requiring browser clipboard access;
   paste replaces only direct character and paragraph formatting, preserving
   links, comments, and revision semantics in one document transaction.
   Header and footer TipTap surfaces register the same Traditional Office alignment and
   copy-format shortcuts. Cross-surface paste retains compatible marks and
   paragraph attributes while falling back from unsupported heading nodes.
   Writer orders Insert commands as Pages, Table, Illustrations, Links, Header
   and Footer, and Text. The Links group creates paired body-bookmark
   boundaries around a selection or caret, keeps copy identity separate, and
   represents internal destinations as `#name` marks with an explicit
   missing-target class. The References target picker also exposes those body
   bookmarks alongside captions; bookmark cross-reference atoms retain target
   identity, source instruction, live display text, and a missing-target state.
   DOCX import/export maps the boundaries to native
   `w:bookmarkStart`/`w:bookmarkEnd` ranges, bookmark cross-references to `REF`
   fields, and internal links to `w:anchor` rather than relationship-backed
   external hyperlinks. Footnote and endnote reference atoms share a
   transaction-normalized graph with their editable definitions. References
   retain identity through moves, copied references clone definition content
   under a new identity, deletions remove the complete pair, and one appended
   transaction keeps undo and redo atomic. Each note kind numbers independently
   in reference order; new or reconstructed footnote definitions use the
   reference section and newly inserted endnotes use the final section. The
   structured HTML parser
   gives note markers precedence over generic superscript so native DOCX notes
   survive repeated import and export. The Insert Text group creates atomic
   `PAGE`, `NUMPAGES`, `SECTION`,
   `SECTIONPAGES`, `DATE`, and `TIME` fields. A position resolver maps every
   field to its measured Worker/WASM layout block, physical page descriptor,
   owning section, and the set of physical pages occupied by that section, so
   continuous sections sharing one page retain truthful section-page counts.
   Pagination-driven numeric refreshes do not add undo history and deliberately
   leave clock fields unchanged; the scoped F9 command refreshes date and time
   too, as one history action. A transaction-mapped identity plugin lets moved
   fields keep their IDs while copied atoms receive fresh, redo-stable IDs.
   DOCX import atomizes only complete, non-nested inline `w:fldSimple` or
   complex fields. Nested, incomplete, cross-paragraph, deleted, and
   instructionless structures keep their rendered result as text and emit a
   structural compatibility warning instead of being flattened into a false
   editable field. Native export writes live simple fields and requests a Word
   field update on open. Page Layout keeps direct margin, orientation, paper,
   and equal-column presets in the ribbon; custom margins and advanced columns
   open the corresponding controlled Page Setup tab instead of duplicating a
   second settings model. References separates footnotes, captions, citations
   and bibliography, and field updates; Review derives previous, next, accept,
   and reject availability from the live revision ranges and moves the
   selection to the next resolvable change. View keeps document modes before
   display controls and calculates One Page and Page Width zoom from the live
   page metrics and scroll viewport. Picture, Table, and Header and Footer tabs
   carry explicit contextual semantics and disappear with their selection. The
   Spreadsheet shell now uses the same responsive ribbon contract with its own
   command catalog and workbook semantics. It keeps the Traditional Office tab order, places
   Conditional Formatting under Home and Styles, exposes real sorting under
   Data, and owns F9 workbook recalculation in the root-scoped keyboard
   controller. Paste, Cut, and Copy live in the first Home group; both button
   clicks and Traditional Office shortcuts cross one clipboard command port before using the
   permission-resilient browser/local clipboard implementation. Paste Special
   extends that port with a versioned rich snapshot and a pure bounded planner
   instead of delegating mutation to rendered vendor controls. Quick All,
   Values, Formulas, and Formatting commands and the `Cmd/Ctrl+Alt+V` dialog
   converge on the same request model. The planner tiles or transposes the
   source, translates relative formula references, validates protection,
   pivots, merges, bounds, unsupported formula state, and arithmetic, then
   publishes one controlled workbook value and one Undo record for at most
   50,000 cells. External text enters through a rectangular TSV snapshot and
   never gains rich-formatting authority. Dialog Apply resolves the current
   grid again on the next frame so controlled remounts cannot strand keyboard
   focus. Format Painter
   uses a dedicated format-pattern model and React session hook behind another
   typed command port. It captures owned native style attributes only, expands
   or tiles the pattern across same-sheet or cross-sheet targets, guards both
   sides at 50,000 cells, suppresses repeated engine callbacks for the same
   target, and emits one `batchCallApis` update per application. One-shot and
   double-click locked state therefore stay outside the ribbon, while the
   ribbon exposes pressed status and Escape cancellation without copying cell
   values or workbook structure. AutoFilter follows the same boundary: a pure
   workbook model derives a finite current region or accepts an explicit
   multi-row range, a React hook owns controlled enable/disable state, and a
   typed command port connects the Data ribbon plus `Cmd/Ctrl+Shift+L` and
   `Alt+ArrowDown`. The Fortune adapter only exposes the header trigger and
   value menu, adding dialog, checkbox, focus-restoration, and keyboard
   semantics without making vendor DOM text the command contract. Freeze Panes
   also stays controlled: a pure model translates the current Traditional Office cell into
   rows above and columns left or applies top-row and first-column presets, a
   typed View command performs one immutable workbook update, and the native
   Fortune/XLSX frozen-pane model remains the render and serialization
   boundary. Its menu owns Arrow, Home, End, Enter, Escape, pressed state, and
   focus restoration without introducing a parallel vendor toolbar. Home and
   Cells now exposes the existing typed structure commands through one Rows and
   Columns menu: insert and delete availability comes from the live selection,
   execution stays in the workbook command port, and successful actions return
   focus to the grid instead of introducing ribbon-local mutations. Home and
   Alignment owns cell merging through a split control whose primary action and
   `Ctrl+M` execute Merge and Center. Merge Cells, Merge Across, Unmerge Cells,
   and Unmerge and Fill use the same typed command path, derive availability
   from the selection and native Fortune merge ranges, and emit one
   `batchCallApis` transaction per user intent. Home and Editing owns one Clear
   disclosure for All, Formats, Contents, Comments, and Hyperlinks. Delete and
   Backspace share the Contents command; all five modes transform the live
   sheet immutably, retain merge geometry, preserve cell state outside the
   selected mode, and emit one controlled workbook batch. The
   catalog is the source of truth for labels, locations, and shortcut metadata;
   typed command ports remain the source of truth for execution and
   availability. The labelled status surface provides arrow-key traversal
   across view and zoom
   controls; compact Web widths remove secondary section, proofing, citation,
   and save items in stages while retaining physical-page and zoom feedback.
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
| Modal dialog | Bookmarks, captions, cross-references, links, image descriptions, notices, and confirmations | Centered in the viewport with bounded height and an independently scrolling body. The body-level portal makes every non-dialog body child inert. Focus stays inside, destructive confirmations initially focus the safe action, and closing returns to the actual invoker unless the completed command deliberately restores the editor. |

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

Paragraph formatting is a node-level review concern rather than an inline
mark. When tracking is active, one formatting transaction compares the complete
canonical paragraph-property set before and after each affected paragraph or
heading. A changed set receives one shared `paragraph-formatting` identity and
the first complete prior snapshot; later edits retain that original review
baseline. Accept clears only the revision attributes. Reject restores the
validated snapshot without changing child content. Both decisions are atomic
across every node with the same identity and have their own undo boundary.
DOCX import/export maps that model to strict or transitional `w:pPrChange`,
while Yjs stores the same node attributes and immutable decision audit without a
parallel transport model.

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
| Document | One TipTap/ProseMirror body tree retained across editing and read-only preview, controlled TipTap header/footer surfaces with direct paper-margin editing and a contextual ribbon, one typography and page-chrome baseline across editing, preview, and PDF surfaces, a persistent typed heading-navigation pane with active-selection tracking, filtering and collapsible keyboard traversal, host-owned selected-text menus with full document context and range-mapped async commands, responsive and keyboard-operated paragraph-style and list galleries, typed bullet and numbering commands with restart/continue/start controls, typed physical-page and section-page descriptors, repeated first/default/even page chrome, a versioned structured model with an HTML compatibility representation, prefix-reused visual-line measurement and pages, page decorations, page-aware horizontal and vertical rulers for page margins, paragraph indents and typed tab stops, structured list-item pagination, explicit paragraph and list-item direction, compact spacing and pagination controls, typed inline/square/tight/through/top-and-bottom/no-wrap image layout and floating-image drawing layers, imported style-inherited paragraph properties, structured inline tabs, and theme-aware run font/size/color/background import | Worker plus resumable Rust/WASM flow pagination and Rustybuzz shaping across CSS-matched registered text runs; the same live result remains mounted in read-only preview and is consumed by bounded browser PDF capture, including eligible list paragraphs, Unicode bidi level segmentation, ordered per-grapheme font fallback, packaged Latin/CJK/Arabic/Hebrew faces, and structured left-to-right tabs, with explicit DOM and JavaScript fallbacks for text affected by supported floats | Language-complete font substitution, complete Word style and numbering coverage, locale-complete and bidirectional tabs, arbitrary floating-object geometry and non-image drawing layers, complex table flow, searchable/vector PDF output, and loss-preserving OOXML package state |
| Markdown | TipTap visual editing with a resizable source-and-preview split view by default, source-aware ribbon formatting and shortcuts, bounded source-native history with coalesced typing and selection restoration, host-defined selection menus across both editing surfaces, GFM tables, strikethrough, autolinks and nested task lists, controlled source state, coalesced preview rebuilds, proportional pane scrolling, keyboard-adjustable pane sizing, a flat measure-bounded reading surface, optional visual or source-only views, and a full-workspace Source/Preview switch on phones | No kernel required for normal editing | CommonMark differential fixtures, multi-megabyte profiling, and an off-main-thread parser boundary when measurements justify it |
| Spreadsheet | Fortune Sheet grid integrated with the shared Office shell, an Office-oriented command catalog and tab order, quick-access Undo and Redo, a priority-aware adaptive and collapsible ribbon with temporary tab expansion, Home and Styles Conditional Formatting, Home and Cells row/column insertion and deletion through the shared typed structure commands, Home and Alignment Merge and Center plus Merge Cells, Merge Across, Unmerge Cells, and Unmerge and Fill through one native-model-backed controlled batch per intent and the Traditional Office `Ctrl+M` shortcut, Home and Font native strikethrough through the same typed formatting command and Traditional Office `Cmd/Ctrl+5`, Home and Editing Clear All, Formats, Contents, Comments, and Hyperlinks with Delete/Backspace mapped to content-only clearing and merge geometry retained, Data sorting and AutoFilter with bounded current-region discovery, controlled filter state, Traditional Office toggle and header-menu shortcuts, and accessible vendor menu semantics, View and Window Freeze Panes with Traditional Office current-cell, top-row, first-column, and unfreeze patterns backed by controlled state and native XLSX frozen panes, Formulas and F9 workbook recalculation, an executable Home clipboard group with permission-resilient Paste, Cut, and Copy plus an Office-style one-shot and locked Format Painter for bounded cross-sheet native-style patterns, shared grid/formula/footer visual tokens and one spreadsheet accent across selection, sheet tabs, menus, and zoom, an A3S-owned accessible single-row workbook footer for creation, activation, rename, duplicate, color, hide, reorder, delete, selection status, and zoom, a resize-observed worksheet viewport that keeps the active tab and its touch action visible across desktop, compact, and phone-width changes, one cell/worksheet context-menu surface whose displayed accelerators remain executable while the menu owns focus, Arrow/Home/End tab navigation plus Shift+F10 and native context-menu access, typed editing and calculation command ports, A3S-owned deterministic Arrow, Enter, Tab, Home, PageUp/PageDown, extended-selection, row, column, and all-cells keyboard commands, selection-preserving focus restoration across controlled workbook remounts and F2/Escape editing transitions, Cmd/Ctrl formatting, undo/redo, clear, a permission-resilient editor-local clipboard fallback, Shift+F11 sheet creation, and Ctrl/Cmd+PageUp/PageDown sheet-navigation shortcuts, direct font-family, horizontal/vertical-alignment, text-wrap, general/number/percent, and decimal-place controls backed by native cell-style keys, live count/sum/average selection summaries, operation-driven sparse-workbook projection, guarded controlled-value remounts that reject stale engine callbacks, and no-history result patches with cell-scoped Fortune fallback | Versioned, cancellable Worker/Rust-WASM calculation sessions using the shared bounded Rust formula parser, retained formula ASTs, incremental forward/reverse dependency graphs, dirty-subgraph recalculation, cross-sheet references, and a dynamically loaded JavaScript fallback | A3S-owned virtual grid, moving replacement projection off the main thread, broader Excel formula semantics, A3S-owned custom number-format evaluation, and print layout |
| Presentation | Scene canvas with ordered typed multi-selection, persistent nested browser groups, native PPTX group-node export, exact keyboard-accessible table-dimension insertion, native slide/object context actions with optional AI actions, a separate object/content editing state, one on-demand TipTap instance, collective move/scale/nudge/clipboard/delete/layer commands, selection-bound alignment and distribution, typed group/ungroup commands, one typed dispatcher for ribbon commands, editor-scoped shortcuts with post-command selection focus, direct beginning/current-slide playback with fullscreen fallback, frame-coalesced transactional move/resize previews that commit once on pointer release, two-level thumbnail node and scene windowing, and a phone stage that top-aligns the primary canvas while retaining one readable slide indicator beside view and zoom controls | Revisioned, cancellable Worker/Rust-WASM slide-relative alignment and object-set snapping with typed visual guides and a JavaScript fallback | Arbitrary rotated or reflected PPTX group transforms, connectors, theme resolution, text fitting, kernel-owned thumbnail layout, and slide serialization |
| PDF | PDFium-backed page rendering with an A3S-owned responsive toolbar, a scrollable page-thumbnail rail with active-page synchronization and bounded bitmap/DOM windowing, a focus-contained phone page drawer, and typed capability controllers for navigation, zoom, search, basic annotations, annotation color, opacity, compatible stroke-width defaults and selection updates, history, and save; page and search drafts cancel without accidental commands, shortcuts remain scoped to the PDF root, and compact widths retain page status while exposing search-result traversal, navigation, zoom, and history through the keyboard-operated overflow menu | PDFium WebAssembly | Forms, redaction review, page organization, and reopen fixtures |

Spreadsheet Tables/ListObjects remain part of the controlled workbook model,
not a formatting sidecar. The active cell selects the contextual Table Design
ribbon; creation, design, structure reconciliation, conversion, collaboration,
and XLSX import/export all read the same stable ID, range, column, filter, and
style record. Table presentation is resolved inside Fortune's existing
visible-cell Canvas callbacks. A row-cached resolver inspects only tables that
can intersect the cell being painted, and conditional formatting remains the
higher-priority visual layer. Ordinary style changes therefore do not stamp or
allocate cells. Convert to Range is the deliberate boundary where the selected
appearance becomes native cell fills, emphasis, and borders; a bounded
footprint check switches large matrix projections to sparse `celldata` before
materialization. Fortune row/column operations are reconciled back into table
ranges, column identities, filter offsets, and canonical header cells, while
merge and worksheet AutoFilter commands reject intersecting ownership.

Native XLSX rich-text editing is reconciled at the controlled Fortune boundary,
not inside the rendering engine. A pure model normalizes at most 512 runs and
32,767 UTF-16 code units, derives one contiguous replacement, uses binary run
lookup for inherited styles, and coalesces only semantically identical adjacent
runs. Authenticated `data[row][column]` operations permit text changes at exact
coordinates. In the full-projection compatibility path, a non-structural batch
without a cell coordinate may restore prior runs only when the visible text is
exactly unchanged; structural operations disable coordinate inference. This
keeps formula-bar and F2 commits to one host revision and one Undo record while
discarding Fortune focus callbacks that merely flatten or strip run metadata.

Formatted paste adds a separate one-shot authority instead of trusting live
DOM markup. A paste event in the formula bar or F2 editor records the exact
sheet object, coordinate, controlled source text, UTF-16 selection, sanitized
clipboard runs, and plain text before the browser mutates the editor. The
incremental and full projection paths may consume that authority only when an
authenticated cell operation emits the exact predicted replacement. Prefix
and suffix runs always come from the controlled source; only the inserted
range may use clipboard font family, point size, RGB color, bold, italic,
strikethrough, or underline. This also converts eligible plain or empty cells
to native inline strings without granting structural callbacks a coordinate.
The parser never retains clipboard HTML, ignores active/non-content elements,
and rejects mismatched text, invalid UTF-16, more than 256,000 HTML characters,
32,767 cell characters, or 512 resulting runs. Imported semantic color
identities remain attached to untouched runs; pasted colors are explicit RGB
unless a future package-level source can prove a native palette identity.

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
margins. Kernel protocol 16 therefore computes body height from page height
minus physical margins only; page-chrome heights position repeated overlays
and do not subtract the same space again. Browser PDF export locates the ready
surface through the host's stable artifact ID, clones its ProseMirror DOM once,
removes editing-only state, and crops exact physical pages from bounded capture
batches when their geometry is uniform or from exact per-page viewports when
it differs. Explicit descriptor pages remain a fallback for surfaces without a
registered live document.

Eligible structurally plain large DOCX files take a narrow import path in a
dedicated Worker. The main thread inflates `word/document.xml` once and
transfers its exact `ArrayBuffer`; the Worker decodes it, performs the strict
WordprocessingML tag and envelope checks, and streams paragraphs or table rows
in batches of 2,048. Table messages carry one text vector plus transferable
`Uint32Array` bands for row-to-cell and cell-to-paragraph counts, so neither
side clones a 100,000-row nested object graph. The main thread constructs the
complete, schema-validated `WorkDocumentModel` and canonical HTML from those
bounded messages, but it does not create the complete live ProseMirror tree.
The TipTap source initially materializes
only the first two leaf chunks. Every later leaf contains an internal
`documentLazyBlock` whose position tape has exactly the same `nodeSize`, text
offsets, and block boundaries as the canonical payload. Ordinary text uses
128-block leaves above 2,048 blocks, while tables use 16-row leaves above 512
rows. Runs longer than 32 table leaves are nested under geometry containers, so
a 100,000-row table exposes 6,250 logical leaves through only 196 top-level
containers without changing any canonical position.

The Worker boundary is deliberately fail-closed. An explicit `ineligible`
result continues to the complete Mammoth and OOXML-marker importer. Worker
creation, execution, message-validation, or timeout failure terminates that
exact Worker and retries the same narrow parser synchronously; cancellation
terminates it and propagates `AbortError`. The production import Worker is
23.9 kB uncompressed, about 97.4 percent smaller than the earlier
dependency-heavy prototype, and contains neither JSZip nor the rich
section-normalization pipeline. Envelope counting uses one scan, and the row
parser emits the same columnar shape that the Worker protocol consumes.

Hosts may reserve an artifact ID before starting import. The Playground mounts
a blank editor shell under that ID while the Worker parses, then applies the
final controlled model to the already mounted editor. Parsing and editor
initialization are therefore parallel critical paths rather than serial work;
source-backed export is rekeyed atomically to the reserved identity. A failed
or cancelled import removes only its own placeholder and restores the previous
workspace.

The chunk NodeView omits descendant DOM for an off-screen leaf. A visible lazy
leaf receives a read-only semantic preview from a structure-keyed DOM pool;
paragraph and table nodes are updated in place and returned to the pool as the
window moves. Pointer selection, search, model-boundary navigation, or a real
edit hydrates only the selected chunk from its canonical payload before the
selection is applied. Hydration is history-free, does not publish a controlled
edit, and transfers cached document statistics because its logical content is
unchanged. The placeholder retains estimated content height plus indexed
pagination spacer height, and nested leaf heights aggregate into the collapsed
container. Expanding, collapsing, and pooling therefore do not change the
scroll range. A cumulative geometry index and binary search map the viewport
to its exact leaves. `Ctrl+Home`, `Ctrl+End`, and the macOS equivalents operate
on model positions first and never depend on mounted DOM.

CSS `content-visibility: auto` is not the production long-document boundary.
It can skip browser layout and paint for an off-screen subtree, but the
ProseMirror nodes, descendant DOM, JavaScript projection work, mutation and
selection bookkeeping, and accessibility tree still exist. The production
NodeView window instead omits descendant DOM and page-sheet DOM while retaining
model-authoritative positions, indexed geometry, and semantic previews only
for the visible leaves. `content-visibility` remains useful as a measured
benchmark mode where a surface owns ordinary DOM, but it cannot replace model
windowing or visible-range Canvas painting.

Pagination reads the canonical JSON payloads of lazy chunks without hydrating
them into ProseMirror. A text edit reuses the measured block and page prefix
before the first changed position, locates the current page with a binary
search, and shapes only real paragraph or heading nodes. This prevents a local
tail edit from filtering, comparing, or recreating all 100,000 prior blocks.
Selection, IME, undo, comments, revisions, and exported model positions remain
model-authoritative throughout this process.

Controlled ownership remains intact. Text-only transactions in an eligible
lazy document are coalesced to one animation-frame publication. The editor
finds the changed persistent chunk, updates document statistics by subtree
delta, materializes the complete JSON model from the payload registry, serializes
only that chunk, and patches its indexed range in the canonical HTML. The host
still receives complete `html` and `model` values. A composable polynomial
fingerprint hashes the changed chunk and combines its cached prefix, chunk, and
suffix segments, so consecutive publications do not rescan the complete HTML
string. Process-local schema trust admits that precomputed fingerprint only for
the exact live model snapshot; cloned or persisted models still rescan and
verify the complete HTML, including fingerprints written by earlier releases.
Formatting, structure, or unsupported edits deliberately fall back to complete
schema serialization. The lazy projection is process-local and
parser-authenticated; rich compatibility imports, cloned untrusted models, and
collaboration bindings retain the complete TipTap path.

Physical page chrome has a separate React window above 24 pages and mounts only
the nearby sheets. A WeakMap-backed page-surface registry retains all page
frames for PDF geometry without recreating thousands of page-sheet elements.
Pagination widgets are indexed by leaf chunk and appear only with a mounted
leaf; logical spacer height remains in the chunk geometry when the widget DOM
is absent. Single-column sections deliberately use ordinary block flow rather
than CSS Multi-column, while true multi-column sections retain the column
algorithm. Word content-autofit tables ignore TipTap's inline grid widths at
render time so the browser sizes them from content; fixed and percentage table
geometry remains unchanged.

This path is not a claim that every rich 100,000-block DOCX is already bounded.
Unsupported tags, relationships, tracked structures, drawings, or other rich
features route to the complete compatibility importer. Extending the same
model-preserving window to those paths remains a performance gate.

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

The public formatting fixture also exercises the non-windowed review boundary
with independent character- and paragraph-formatting cards. Focused A3S Test
suites reject each kind separately, verify that text and the other revision
remain intact, and prove full paragraph alignment, indentation, spacing, and
line-height restoration with clean browser diagnostics.

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
Fonts matching order so Office-style values such as 680 or 730 select a registered
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

Native hidden text is resolved before layout eligibility is chosen. The
canonical TipTap mark keeps inherited, hidden, and explicit-visible states;
DOM paint hides `data-office-hidden-text="true"` by default, while the
editing-only View mode reveals it with a dotted underline. Preview and detached
PDF snapshots never inherit that view class. A paragraph containing hidden text
uses browser-authoritative measurement because its invisible run must not
contribute Worker/WASM advances. This is an explicit per-paragraph fallback,
not a second document tree or a second pagination model.

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
alignment, text distance, signed horizontal and vertical offsets, their column,
paragraph, margin, or page reference frames, four-edge percentage cropping,
relative drawing order, behind-text placement, overlap and cell-layout policy,
anchor locking, and alternative text. Applying the complete draft
updates only changed attributes in one separated TipTap history entry; cancel
and Escape preserve the node and return focus to the exact invoker. Dimensions
that were only projected into rounded centimeter fields retain their exact
imported pixel values when untouched. Inline images stay in normal flow;
supported left- or right-aligned square images use browser floats, tight and
through images additionally project their DrawingML polygon through CSS
`shape-outside`, top-and-bottom images clear surrounding text, and no-wrap
images use absolute placement so body text does not reserve their rendered
height. A one-pixel pagination anchor keeps each free-floating image attached to
its document position, while `behindDoc` determines whether it paints behind or
in front of text. Paragraphs
following any supported side-wrapped float remain on the DOM visual-line path
because the available line width changes while the float is active. The
paginator reserves and observes the rendered height of flow-affecting images,
uses only the anchor height for no-wrap images, then sends the measured text
fragments to the same Worker/Rust-WASM page-layout protocol. Square,
tight, through, top-and-bottom, and no-wrap metadata round-trips through DOCX
`wp:anchor`; tight and through anchors preserve their wrap side, edited flag,
and ordered `wrapPolygon` vertices. Aligned anchors remain aligned,
while precise offsets preserve `positionH` and `positionV` reference frames
through edit, preview, and regenerated DOCX. Four-edge crop geometry remains
typed in the controlled model and round-trips as DrawingML `a:srcRect` for
inline and floating pictures. Supported floating pictures also preserve
drawing-layer relationships through unsigned `relativeHeight`,
`behindDoc`, `allowOverlap`, `layoutInCell`, and `locked` anchor attributes.
The browser maps the unsigned drawing order into bounded foreground and
behind-text stacking ranges while retaining the exact OOXML value for export.
Each image node also owns one normalized eight-hex-digit object identity plus
its DOCX drawing-property, anchor, and edit identifiers. A ProseMirror
transaction normalizer gives a pasted copy a fresh complete identity while a
move, cut-and-paste, deletion followed by undo, and redo retain the surviving
object's identity. Existing nodes and newly inserted images are normalized at
the editor boundary, and duplicates prioritize the mapped pre-transaction
object even when a copy is inserted before it. DOCX import retains
`wp:docPr@id`, `wp14:anchorId`, and a conforming `wp14:editId`. Export keeps the
drawing-property and anchor identity stable independently of regenerated media
relationship IDs, retains an untouched edit identifier, and advances the edit
identifier when the image itself is changed or moved.

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

The declared direct-formatting layout slice now has a real Traditional Office Writer reference
gate. Its deterministic A4 fixture uses installed Arial runs, explicit OOXML
paragraph spacing, a fixed centered table, direct cell fills and borders, and
physical cell margins. Automatic `w:line` values retain their original OOXML
multiples for export while the browser applies the measured Traditional Office single-line
font metric; imported tables no longer receive editor-only block margins. The
Windows parity workflow exports the same fixture through Traditional Office, captures both
794 by 1123 CSS-pixel pages, and rejects page-size, landmark, or bounded pixel
differences. Language-complete font substitution, variable font axes, the
remaining character and table style properties, non-image floating-object
layering, row-internal table splitting
inside a single long paragraph, footnote balancing, multi-column flow, and
mixed-size sections require the
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
paragraph remain atomic and may overflow. Nested tables contribute their own
row boundaries to the containing cell's synchronized fragment plan. A
`rowspan` cell is projected through every covered physical row with clipped,
contiguous document ranges; page breaks before continuation rows render inside
the spanning cell instead of inserting a conflicting synthetic table row.

Table creation is selection-safe: a non-empty text selection is preserved and
the chosen table is inserted after its containing block instead of replacing
the selected content. The insert ribbon exposes a keyboard-navigable 8 by 10
size picker. Entering a table opens separate Design and Layout contextual tabs.
Design owns keyboard-operated style presets, the header row, multi-cell
shading, and a reusable border pen with whole-selection, outside, inside, side,
inside-horizontal, and inside-vertical targets. Layout owns row and column
insertion or deletion, cell merge and split, horizontal and vertical alignment, repeated headers,
atomic rows, percentage- or centimeter-based column width and centimeter-based
row height fields, equal row and
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
full-width table. Per-column percentage preferences are editable and export as
`tcW`, including merged cells, while `tblGrid` pixels remain browser layout
fallbacks. Window/content autofit remains responsive in the browser.
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
properties remain higher-priority layers for cell presentation. Nested tables
retain independent geometry and formatting through the same marker and export
paths, and target the innermost active table during editing. Theme-derived
presentation resolves to stable RGB for editing and preview, while untouched
run colors, run shading, cell fills, and per-edge borders retain semantic theme
references on regenerated DOCX output. Explicit color edits export direct RGB.
The real styled-table fixture exercises a centered 62.5% table, table margins,
a first-cell margin override, and matching edit/preview geometry through A3S
Test.
Command availability comes from the ProseMirror table state, so actions that
cannot apply to the current selection remain disabled. Less-common conditional
paragraph properties outside the supported layout set remain explicit fidelity
gaps.

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

XLSX file import starts two bounded, one-shot Workers before the calculation
session. The general Worker runs SheetJS off the main thread and streams dense
worksheets in 512-row chunks or sparse worksheets in 8,192-cell chunks. In
parallel, the package Worker decompresses each worksheet XML part once and
builds the feature and style gates used by the rest of the importer.

The package Worker also owns a deliberately narrow, fail-closed path for plain
OOXML workbooks. It first authenticates an exact package envelope: workbook,
content types, relationships, and worksheet ownership must form a closed set
with no styles, shared strings, external targets, or unknown parts. It then
parses only ordered primitive numeric, Boolean, error, and simple inline-string
cells. A worksheet-start message authenticates its declared range before any
cell is accepted. Each later transferable columnar chunk covers at most 256
rows and 4,096 cells and carries packed coordinates, value kinds, numeric
values, and the remaining text vector instead of a million nested cell
objects. The hot cursor validates row and cell coordinates in place with
bounded character-code arithmetic; it creates neither an address substring nor
a regular-expression match per cell. The main thread treats those chunks as
provisional until the complete worksheet XML has been consumed.
A formula, style index, rich/shared string, unknown element or attribute,
malformed entity, invalid address, relationship, or trailing payload rejects
the candidate instead of guessing. Package authentication cancels the
speculative SheetJS Worker early; a later cell-level rejection discards every
provisional row and restarts SheetJS with the original private bytes. Worker
creation, protocol, timeout, and cancellation failures retain the synchronous
parser as the final compatibility fallback. This removes the second complete
worksheet decompression and SheetJS object-materialization pass for eligible
million-cell files without weakening the full XLSX path.

Primitive imported cells and their sparse rows are frozen once, and an
identity-keyed profile authenticates the resulting matrix as directly
consumable by Fortune. A host may reserve worksheet IDs while the Playground
mounts a blank workbook shell. If the final workbook is editable, preserves
the same sheet count and identities, and contains no stateful structure such
as charts, protection, merged geometry, filters, validation, images, or custom
view state, the editor replaces that certified matrix through the existing
Fortune instance. Fortune adopts the frozen matrix and rebuilds only its
formula index; it does not clone every populated cell or remount the workbook.
Any failed predicate or update exception returns to the complete controlled
remount path. This narrow eligibility gate preserves correctness for rich and
multi-sheet imports while removing the simple million-cell first-mount task.

The browser pagination implementation is split by capability instead of
accumulating in one editor module. `work-document-pagination.ts` owns only the
TipTap extension and public facade; dedicated modules own block measurement,
table fragmentation, text-layout request collection, visual-line measurement,
DOM geometry, page-break decorations, and shared contracts. The Rust kernel
uses the same separation: the crate root owns the ABI and flow layout, while
text validation, UTF-16 offset mapping, shaping, and their tests live in
focused modules. This keeps the fallback and WASM paths independently
testable without duplicating their public protocol.

Live DOM measurement is cooperatively scheduled in bounded 32 ms slices. It
yields only between top-level ProseMirror blocks, so one block is always
measured atomically against its rendered element. Document invalidation aborts
the stale pass before it can replace the reusable snapshot, while the existing
single-flight coordinator coalesces one fresh run. The synchronous measurement
path remains available for deterministic callers and parity tests.

Layout requests carry a bounded, deduplicated page-style table and let each
measured block reference its section geometry. A geometry transition opens a
new physical page in both the JavaScript fallback and Rust/WASM kernel, and
each result page owns its exact metrics. The live page stack, navigation
thumbnail capture, and PDF capture all consume those result metrics, including
mixed custom sizes and orientations, instead of projecting the active
section's dimensions across the document.

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

`encodeDocumentSnapshot` and `decodeDocumentSnapshot` expose that complete
controlled value through the public Core entry point. The v1
`a3s.office.document.snapshot` envelope is deterministic bounded JSON and uses
the
`application/vnd.a3s.office.document-snapshot+json;version=1` media type. It
requires a synchronized structured model and retains every JSON-defined field,
including layout, review, comment, and bibliography state. A different schema
or version, malformed or oversized JSON, cyclic runtime input, and stale model
fingerprints fail closed. This is the persistence and transport boundary for
hosts that need exact browser/agent revision snapshots; they must not substitute
a lossy HTML-to-Markdown conversion.

This is a transitional persistence boundary, not fully loss-preserving OOXML
yet. DOCX still converts through Mammoth, but DOCX, HTML, and text imports
create the structured model immediately through the same extension schema used
by the editor. DOCX export materializes synchronized HTML from the structured
model before OOXML generation, so a stale HTML cache cannot override model
changes. Preview retains the mounted structured editor tree, and live browser
PDF export captures that same paginated surface; compatibility HTML remains an
import/export boundary.

The first package-state slice retains the original DOCX Blob by artifact ID.
After the generated core package is complete, safe source-only OPC parts are
copied byte-for-byte and their content-type declarations and relationships are
reconnected with collision-free IDs. The generated core parts remain
authoritative. Digital signatures, which are invalid after editing, plus VBA,
ActiveX, and custom-ribbon content are not propagated into the macro-free DOCX.
If a persisted artifact has lost its registered source Blob, export fails
closed. A persisted SHA-256 fingerprint also prevents a different DOCX from
being re-registered under the same artifact ID. Unknown attributes, elements,
and `mc:AlternateContent` branches are now selectively merged inside
`word/settings.xml`: only passive ignorable markup that is relationship-free,
structurally valid, and non-conflicting with generated Word settings survives.
Strict and transitional UTF-8/UTF-16 sources share this path. The package graph
also retains source font-table metadata and source-only internal obfuscated-font
payloads, then rewrites their references to collision-free relationship IDs.
External fonts, mismatched relationship or content types, duplicate identities,
and source payload paths that collide with generated parts are disconnected.
This preserves embedded fonts for native DOCX consumers without loading those
binaries into the browser editor, preview, or PDF renderer. Passive,
relationship-free extension trees at the roots and uniquely matched nodes of
`word/styles.xml` and `word/numbering.xml` are also retained. Styles match by
type plus style ID. Imported abstract-numbering, concrete-numbering, and level
metadata follows regenerated IDs. Duplicate, source-only, malformed, and
ambiguous one-to-many mappings fail closed, and generated Word semantics remain
authoritative. Relationship-free passive extensions from non-OOXML ignorable
namespaces also follow uniquely matched picture drawings in regenerated
document, header, and footer parts. The identity is the normalized anchor plus
drawing-property ID; header and footer imports retain both through sanitized
editable HTML. Strict/transitional UTF-8/UTF-16 sources share the same path.
Source-only or duplicate drawings, relationship-bound content,
Microsoft/OOXML semantic namespaces, and ambiguous identities fail closed,
while generated geometry and media win. The paragraph-identity slice preserves
the same class of passive extension on uniquely matched, unchanged paragraphs
and their paragraph properties. Native `w14:paraId` plus `w14:textId` survives
sanitized body and page-chrome HTML. Text edits rotate `textId`, formatting-only
edits and moves retain it, and copies or splits receive new paragraph IDs.
Changed text versions, duplicate identities, relationship-bound content, and
Microsoft/OOXML semantic branches fail closed; generated paragraph semantics
remain authoritative. The table-identity slice applies the same policy to
`w:tbl`, `w:tr`, and `w:tc` plus their property nodes. Rows carry native
`w14:paraId` and `w14:textId` through body and page-chrome HTML. Tables match by
their ordered directly owned row IDs, while cells match by their owning row and
directly owned paragraph IDs. Text or structural edits rotate a row version;
formatting-only edits and moves retain it, and copies receive independent IDs.
Nested DOM row and cell collections are filtered to their closest owning table
or row. Missing, duplicate, cross-kind, relationship-bound, or semantic
identities fail closed, and generated table geometry remains authoritative.
Unknown run, note, comment, and other inline markup remains the next
loss-preservation boundary.

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
document performance budgets, and real Microsoft Office and Traditional Office
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
- Complete row-internal splitting for single indivisible paragraphs and full
  floating-object geometry beyond the supported wrap, precise-anchor, and
  image-layer subset, plus footnote balancing and columns. Row-flow
  pagination, direct-cell and nested-row splitting, merged-cell continuations,
  repeating headers, and the current image-wrapping slice are already
  implemented.
- Keep pagination results as mapped ProseMirror decorations so reflow never
  corrupts selection or undo history.

Exit criteria: deterministic layout goldens for the supported feature set;
incremental reflow does not rebuild unaffected pages; DOCX fixtures round-trip
through Microsoft Word and Traditional Office without losing unsupported package parts.

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

The native table slice is also implemented at this stage. Insert Table creates
one bounded semantic ListObject, the contextual design surface selects among
60 built-in OOXML styles, visible cells receive render-time Canvas styling,
row/column operations reconcile table structure, and Convert to Range
materializes appearance through a sparse-safe path. XLSX table parts,
relationships, content types, styles, and supported filters round-trip, while
Yjs uses ordered ID-keyed records and creation claims for two-client
convergence. Structured-reference calculation, calculated columns, complete
totals authoring, slicers, and external/query tables remain open gates.

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

Presentation design metadata is memoized per controlled content identity.
Thumbnail scenes consume that prepared master/layout catalog directly, and the
editor does not copy the complete slide array merely to resolve one visible
scene. The deterministic performance fixture contains 1,000 slides and 9,000
scene elements. In five fresh headless Chrome 149 processes on the recorded
10-logical-core host, normal mode mounted 18 thumbnail buttons, rendered 13
full thumbnail scenes, retained 1,006 DOM nodes and a median 10.9 MiB
JavaScript heap, reached the final slide in 6.1–13.4 ms, and recorded zero Long
Tasks during import, final-slide navigation, and ten object nudges. The median
keydown-to-object-commit time was 4.5 ms.

`content-visibility: auto` remains a benchmark mode rather than a production
rail rule. The interleaved five-process comparison reduced median editor-ready
time only from 1,637.6 ms to 1,600.3 ms, while increasing median object-commit
time from 4.5 ms to 5.9 ms and median nudge script time from 8.0 ms to 11.5 ms,
with no retained-heap improvement. The rail therefore uses actual DOM and
scene virtualization; CSS skipped rendering cannot replace the JavaScript and
data-processing bounds and can add invalidation cost to the already bounded
window.

Exit criteria: object drag and resize stay interactive on complex slides;
partial rich-text formatting survives PPTX round trips; masters, layouts,
themes, tables, charts, links, and notes have compatibility fixtures for
PowerPoint and Traditional Office.

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
for long files. Files above 48 pages retain virtual spacers and mount only the
viewport plus five items of overscan on either side, with a hard ceiling of 32
thumbnail buttons. Leaving the window aborts the exact pending PDFium thumbnail
task and revokes its object URL. Long-distance Home and End navigation uses an
instant rail jump so intermediate windows do not start disposable bitmap work
or unmount the keyboard destination after it receives focus. At phone widths
the same navigation becomes a focus-contained, dismissible drawer rather than
reducing the document canvas.

`preloadOfficeEditor('pdf', { preloadRuntimeAssets: true })` can move the PDF
editor chunk and PDFium response body ahead of a high-confidence navigation.
Runtime-asset loading is opt-in because the unpacked binary is about 4.4 MiB,
and the helper does not initialize the PDFium Worker. Runtime fetching is best
effort: a failure leaves editor opening unaffected and clears the request cache
so a later intent can retry. The reference benchmark therefore treats
module/asset availability and Worker-side initialization as separate
boundaries. Host CSS `content-visibility` is not applied to the internal PDF
pages: EmbedPDF already mounts only six to seven pages behind its Shadow DOM
boundary, while CSS containment cannot cancel PDFium tasks or release bitmap
memory.

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
round-trip evidence in Microsoft Office and Traditional Office; unsupported content is
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

### Current 100,000-unit evidence

The following cold local Playground measurements were captured on 2026-08-19
on an Apple M2 Pro with 16 GB of memory, Playwright 1.61.1, headless Chrome 149,
and a 1,440 × 1,000 viewport. Document values are the median and range of three
fresh browser processes; Spreadsheet values use five. The benchmark drains
diagnostic memory-counter work before starting the scroll profile so
DevTools-triggered V8 cleanup is not charged to the product. These measurements
are evidence of current limits, not universal release targets. The spreadsheet
readiness milestone is observed inside the page on animation frames; automation
runner polling backoff is therefore excluded from the load time.

| Fixture | Editor visible | Pagination ready | Longest task before visible |
| --- | ---: | ---: | ---: |
| DOCX, 100,000 text paragraphs, 4.50 million characters | 0.397 s (0.396–0.480) | 1.460 s (1.381–1.525), 3,125 pages | 131 ms (128–140) |
| DOCX table, 100,000 rows × 3 columns | 0.813 s (0.810–0.814) | 1.747 s (1.739–1.749), 2,381 pages | 96 ms (88–144) |
| XLSX table, 100,000 rows × 10 columns, 1 million populated cells | 0.480 s (0.474–0.494) | n/a | No task ≥50 ms in any of five runs |

The XLSX row above is a rectangular one-million-cell data fixture, not a
semantic `WorkSpreadsheetTable`/OOXML ListObject benchmark. It proves the
plain-workbook import and visible-range Canvas path; it must not be used to
claim that structured-reference calculation, calculated columns, totals,
filters, or table conversion have the same profile. A dedicated large
ListObject matrix remains required before publishing table-specific load,
mutation, conversion, and export budgets.

The median text phases were 121.6 ms for the complete import, 76.1 ms inside
the Worker, and 56.6 ms for Worker parsing. The table phases were 462.6 ms,
377.6 ms, and 311.6 ms respectively; the already mounted editor applied that
100,000-row controlled model in 32.3 ms. Editor-shell initialization overlaps
these import phases instead of extending the critical path. Against the prior
Worker/windowing baseline, median editor visibility improved from 0.815 to
0.397 seconds for text and from 1.316 to 0.813 seconds for the table;
pagination readiness improved from 1.749 to 1.460 seconds and from 2.366 to
1.747 seconds respectively.

For the XLSX fixture, the complete import measured 375.8 ms (372.3–386.4).
The authenticated package Worker fast path owned 345.5 ms (343.6–355.8), while
main-thread canonical conversion took 13.8 ms (12.7–15.4). The hot worksheet
cursor now authenticates row and cell coordinates directly in the 35.7 MB XML
string instead of allocating and matching one million address substrings. Its
isolated median fell from 178.3 to 124.3 ms (30.3%). Against the immediately
preceding corrected end-to-end baseline, editor visibility improved from 536.7
to 480.2 ms (10.5%), complete import from 432.0 to 375.8 ms (13.0%), and the
Worker fast path from 403.5 to 345.5 ms (14.4%). No import Long Task, browser
error, page error, or console error occurred in the retained five-run result.

These values supersede the earlier 830.5 ms visibility figure. Page-local
observation showed that Fortune had become visible before Playwright's next
progressively backed-off locator poll; the old figure therefore included test
runner latency rather than product work.

| Fixture | Continuous scroll | p95 frame interval | Scroll long tasks | Retained browser state | Bounded render state |
| --- | ---: | ---: | ---: | ---: | --- |
| 100,000 text paragraphs | 120.0 FPS in every run | 10.7 ms (10.2–10.8) | 0 in every run | 70.4 MiB; 2,867 median DOM nodes | 1 of 782 content chunks at readiness, 2 after end navigation, and 4 of 3,125 page sheets mounted |
| 100,000 table rows | 120.6 FPS (120.5–120.9) | 10.1 ms (9.9–10.4) | 0 in every run | 145.8 MiB; 3,106 DOM nodes | 4 of 6,250 leaves at readiness, 1 after end navigation; 196 outer containers; 4 of 2,381 page sheets mounted |
| 100,000 × 10 spreadsheet cells | 120.1 FPS (120.0–120.2) | 12.9 ms (12.9–13.0) | 0 in every run | 38.41 MiB; 903 DOM nodes | Fortune Canvas paints only the visible row and column range; Ctrl+End reached `J100000` in 77.3 ms (76.0–79.7) |

The controlled-edit benchmark sends one browser text-insertion event at the
exact final paragraph or cell, waits for the complete host publication, then
does it again against the new controlled revision. A page-local `beforeinput`
listener starts `Wall latency`, and a `MutationObserver` stops it when the
publication counter changes. This excludes Playwright input transport and
polling backoff while retaining the input transaction, React and pagination
scheduling, and publication. `Publish CPU` measures only construction of the
complete `html` and `model` values. Document values below are the median and
range of five fresh browser processes.

| Fixture | First edit wall / publish CPU | Second edit wall / publish CPU | Publication path |
| --- | ---: | ---: | --- |
| 100,000 text paragraphs | 70.3 ms (69.0–115.4) / 21.8 ms (20.0–54.7) | 53.9 ms (22.5–56.4) / 10.0 ms (9.4–16.8) | `lazy-chunk` in all ten publications |
| 100,000 table rows × 3 columns | 102.3 ms (93.8–121.0) / 57.8 ms (53.5–70.1) | 85.1 ms (78.2–93.3) / 41.2 ms (36.9–45.7) | `lazy-chunk` in all ten publications |
| 100,000 × 10 spreadsheet cells | 52.5 ms (49.6–56.4) / 2.5 ms (2.5–2.6) | 45.9 ms (45.9–47.0) / 2.3 ms (2.2–2.3) | Incremental 3-operation projection of one changed cell; 0 Long Tasks |

Every run reached the exact final paragraph or row, both consecutive edit
markers survived in the final controlled value, and no browser, page, or
console error was recorded. Pooled previews reduced detached scroll DOM by
about 92 percent during profiling. Cached statistics removed the previous
233–246 ms scroll task, while early non-paragraph rejection, binary page
lookup, and measured-prefix reuse removed document-sized work from tail edits.

The earlier document wall figures used a Node-side clock around Playwright
input and polling. They therefore measured runner scheduling as product work
and are superseded by the page-local figures above. Against the same page-local
five-run baseline, cached segment fingerprints reduced the median second text
edit from 101.4 to 53.9 ms and the second table edit from 170.8 to 85.1 ms;
their publication CPU fell from 56.0 to 10.0 ms and from 123.2 to 41.2 ms.

The remaining cold-open boundary is the 100,000-row Worker's approximately
222 ms content scan plus main-thread construction and retention of the
canonical model. The live ProseMirror tree and strict OOXML inspection are no
longer document-sized main-thread constructors. A WASM parser is justified
only if it consumes the transferred bytes directly and emits the same columnar
protocol without adding a second complete copy or JSON serialization step.
During editing, controlled ownership intentionally requires a complete HTML
string; the indexed chunk patch and segment fingerprint avoid complete
semantic serialization and repeat hashing, but allocating that host value
remains a measurable lower bound. A future optional patch callback may remove
that copy only if the complete `onChange` contract remains available.

The million-cell XLSX path now retains about 38.41 MiB after collection and has
neither import nor continuous-scroll Long Tasks in any run. The remaining
cold-open boundary is the package Worker's approximately 344 ms XML scan and
columnar cell construction; canonical main-thread conversion is already about
13.8 ms. A WASM parser is justified only if it consumes the transferred bytes
directly, emits the same bounded columnar protocol, and improves that Worker
phase without introducing another complete copy. Fortune Canvas still owns
painting, but imported matrix adoption is now outside the rich-workbook remount
path for the strictly authenticated simple case.

### Current 1,000-page PDF evidence

The deterministic PDF fixture is 412,852 bytes with SHA-256
`eb86216416a633ce86dd4b31c1c0d9e42792e7bb802946cc0bb6e2b395ae79cb`.
It contains 1,000 A4 pages with stable page markers. The measurements below
combine two counterbalanced three-run batches on the same 2026-08-19 reference
machine and browser described above. Every value is the median and complete
range of six fresh browser processes. `Runtime preloaded` waits for the editor
module and PDFium response body before starting the file-open clock.

| Mode | Editor shell | Viewer ready | First page bitmap | End to page 1,000 | Retained heap |
| --- | ---: | ---: | ---: | ---: | ---: |
| Import-time warm-up | 410.7 ms (399.4–433.1) | 7.430 s (4.403–13.742) | 7.626 s (4.611–13.944) | 21.7 ms (19.6–28.7) | 10.04 MiB (9.48–10.52) |
| Runtime preloaded | 369.1 ms (322.8–428.6) | 9.464 s (5.286–12.390) | 9.655 s (5.478–12.580) | 23.5 ms (18.0–33.9) | 9.72 MiB (9.61–10.47) |

Runtime preloading moved the shell median forward by 41.6 ms (10.1%) but did
not improve viewer-ready or first-bitmap medians on the local server. The
dominant and highly variable boundary is PDFium Worker initialization, not
main-thread layout or local transfer. Runtime preloading remains an explicit
network-latency tool rather than a claimed local-startup acceleration.

All twelve runs mounted exactly seven main-view pages and 15 thumbnail buttons
at readiness. The final window contained pages 986–1,000, page-local End
navigation produced no Long Task, and no browser, page, or console error was
recorded. Cold import produced no task at or above 50 ms. Three runtime-preload
runs recorded one import task each, between 59 and 89 ms; this remains a gate
for later Worker-initialization work rather than being hidden by the preload
API. The deterministic A3S Test suite independently verifies the 15-item first
and final windows, retained keyboard focus on page 1,000, the page counter,
accessibility output, and empty console/page-error evidence.

## Performance and safety rules

- Editing and selection stay on the main thread. XLSX parsing, eligible
  large-DOCX inspection/parsing, and document layout use dedicated Workers;
  canonical large-document model assembly remains a measured main-thread
  boundary.
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
- Requests reject invalid dimensions and more than 200,000 layout blocks.
  Page-style tables remain bounded to 10,000 entries.
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
- Multi-column documents currently retain explicit-break behavior until their
  layout protocol is implemented.

## Verification

The browser kernel is covered at four boundaries:

1. Rust unit tests for deterministic pagination, Unicode line breaking,
   grapheme-safe emergency wrapping, whitespace modes, shared formula parsing,
   scalar dependency calculation, cycles, targets, and validation.
2. JavaScript fallback tests for protocol parity, 100,000-block linear
   pagination, safe page-prefix reuse, no-Worker operation, cancellation,
   sparse Spreadsheet calculation, revisioned session fallback, dense and
   sparse worksheet iteration, and the shared Spreadsheet parity fixtures.
3. A raw generated-WASM ABI smoke test that registers both shipped fonts,
   proves the Latin face lacks CJK glyphs, resolves them through the ordered
   fallback face, verifies mixed-face line metrics, initializes a Spreadsheet
   session, and recalculates a dirty formula chain from a cell patch.
4. Browser checks for real layout and XLSX-import Worker/WASM/font loading,
   shaped-line parity with browser line boxes at non-100% zoom, real
   per-grapheme fallback diagnostics, explicit unresolved-glyph fallback,
   page-view reflow, web-view clearing, page counts, one physical sheet per
   measured page, visible page gaps, nested and RTL list flow, undo behavior,
   and slide-relative element alignment.

Presentation group serialization tests inspect generated slide and layout
OOXML, nested group order, identity child-coordinate transforms, unique
non-visual IDs, marker removal, master/layout scope isolation, grouped
placeholder materialization, cumulative visual scaling, and unsupported
rotation/reflection diagnostics.

Presentation thumbnail tests use a controlled intersection observer to prove
that long decks retain absolute keyboard reachability while bounding mounted
buttons and full scenes, releasing scenes that leave the window, preserving
focus after deletion, and reconnecting the observer when switching between
normal and sorter views. A deterministic A3S Test suite imports the 1,000-slide
fixture, proves the 18-button first and final windows, reaches slide 1,000 with
`End`, performs two object nudges, and captures final focus, accessibility,
console, and page-error evidence.

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
bun run test:e2e:large-documents:check
bun run test:e2e:large-documents
bun run performance:large-documents
bun run performance:large-document-edits
bun run test:e2e:spreadsheet-large:check
bun run test:e2e:spreadsheet-large
bun run performance:large-spreadsheets
bun run performance:large-spreadsheet-edits
bun run test:e2e:large-pdf:check
bun run test:e2e:large-pdf
bun run performance:pdf
bun run test:e2e:large-presentation:check
bun run test:e2e:large-presentation
bun run performance:presentation
```
