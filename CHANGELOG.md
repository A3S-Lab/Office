# Changelog

All notable changes to A3S Office will be documented in this file.

## Unreleased

- Added a Writer command catalog for stable ribbon grouping and WPS-compatible
  shortcut metadata, moved undo and redo into a compact quick-access toolbar,
  added persistent plus temporary ribbon collapse behavior, and made lower
  priority groups compact before the ribbon falls back to horizontal paging.
- Aligned Writer superscript and subscript with the WPS `Ctrl+Shift+=` and
  `Ctrl+=` shortcuts and added deterministic desktop browser coverage for the
  expanded, collapsed, and temporary ribbon states.
- Made the displayed WPS Writer shortcuts executable inside the document for
  font sizing, paragraph alignment and line spacing, heading styles, spelling,
  field refresh, comments, and track changes without capturing host inputs.
- Added a permission-free Writer formatting clipboard with WPS
  `Ctrl+Shift+C` / `Ctrl+Shift+V`, a one-shot format painter, semantic-mark
  preservation, and single-transaction formatting paste.
- Extended WPS alignment and format-copy shortcuts into page headers and
  footers, corrected their superscript and subscript shortcut descriptions,
  and added schema-safe body-format projection for page-chrome editors.
- Reordered the Writer Insert ribbon into WPS-familiar Pages, Table,
  Illustrations, Links, Header and Footer, and Text groups, with page-number
  visibility beside header and footer commands.
- Added direct Writer Page Layout presets for margins, orientation, paper size,
  and one-to-three-column layouts, with custom margins and advanced columns
  routed to the matching Page Setup tab. Deterministic browser coverage proves
  live landscape and two-column rendering, Escape close, accessibility, and
  empty console and page-error diagnostics.

## 0.2.2 - 2026-08-07

- Selected the Word `ascii`, `hAnsi`, `eastAsia`, or complex-script font slot
  from each run's actual text while preserving `bCs`, `iCs`, `szCs`, `cs`,
  `rtl`, and font-hint behavior for multilingual DOCX content.
- Added a deterministic 30-row Latin, Chinese, Arabic, Hebrew, and mixed-format
  fixture with A3S Test coverage and a real WPS Writer PDF layout gate.
- Added a calibrated Chromium native-PDF fallback for WPS reference captures
  when the embedded PDF renderer cannot initialize the exported document.

## 0.2.1 - 2026-08-06

- Matched WPS Writer automatic line layout across common Latin and Chinese
  system fonts with measured per-font advances while retaining the original
  OOXML line-spacing multiple as the DOCX round-trip authority.
- Preserved section-level Word document-grid type and line pitch plus run-level
  `snapToGrid` overrides across DOCX import and export, and stopped exporting a
  generated document grid when the source document does not define one.
- Added deterministic 30-row common-font, 36-row CJK-font, and 18-row document-
  grid fixtures with A3S Test browser coverage and real WPS PDF layout gates.

## 0.2.0 - 2026-08-06

- Added a real WPS Writer page-layout gate that exports a deterministic A4 DOCX
  through WPS, captures normalized A3S and WPS pages, and rejects page-size,
  semantic-landmark, browser-error, or bounded pixel regressions.
- Matched WPS automatic Word line spacing without changing the original OOXML
  multiple used for DOCX export, removed editor-only spacing around imported
  tables, and removed the transparent paginated-page border from content
  geometry.
- Reworked the default Markdown split view into a flat writing-and-reading
  workspace with a bounded text measure, clearer typography for headings,
  quotations, code, tables, and task lists, and no nested preview card chrome.
- Replaced the unusable stacked Markdown phone split with a full-workspace
  Source/Preview switch whose controls remain touch-sized while preserving the
  controlled Markdown value and synchronized visual tree.
- Promoted shared editor context menus to a viewport-bound phone action sheet
  with 44 px rows, bounded internal scrolling, safe-area spacing, and the same
  keyboard dismissal and focus restoration as the desktop menu.
- Separated persistent desktop sidebar state from the temporary phone drawer,
  so resizing across the compact breakpoint closes the overlay immediately,
  preserves workspace focus, and restores the prior desktop preference.
- Kept the documentation language and version selectors visible at ordinary
  laptop widths; the site opens in Simplified Chinese on `latest` and retains
  the frozen `0.1.0` documentation for version switching.
- Separated Word Table Properties updates from preceding edits in TipTap
  history, so undoing a property change no longer removes a newly inserted
  table.
- Replaced the isolated Word alternative-text prompt with one responsive
  Picture Properties workflow for centimeter width and height, a per-image
  aspect-ratio lock, wrapping, alignment, text distance, and alternative text.
- Applied each Picture Properties draft as one separated TipTap history entry,
  while cancel and Escape leave the document unchanged, retain the selected
  image, and restore the exact ribbon invoker.
- Preserved untouched imported image dimensions instead of materializing their
  rounded centimeter display values, and added unit, Playwright, and phone A3S
  Test coverage for size coupling, compact controls, focus, and diagnostics.
- Expanded Word Table Properties into Table, Row, Column, and Cell tabs with
  one validated TipTap transaction for preferred table geometry, selected-row
  sizing and pagination, current-column width, and selected-cell alignment and
  margins.
- Preserved untouched imported table dimensions and partial cell-margin
  inheritance at their exact source values instead of quantizing them through
  the centimeter display fields.
- Kept custom select menus inside the active modal focus scope and routed
  document undo and redo shortcuts from non-text ribbon controls without
  intercepting native input history.
- Extended the deterministic styled-DOCX A3S Test workflow across the complete
  responsive Table Properties dialog, preview fidelity, accessibility, and
  empty console and page-error diagnostics.

## 0.1.0 - 2026-08-02

- Extracted the complete A3S Web document, spreadsheet, presentation, and PDF
  editor engine into an independent package.
- Added React, Vue 3, Web Component, and framework-free Core entry points.
- Added DOCX, XLSX, PPTX, PDF, HTML, Markdown, text, CSV, XLS, and ODS file
  workflows.
- Added a colocated PDFium WebAssembly asset with an overridable URL.
- Added a browser-neutral exact-unit Office layout-renderer contract and a
  source-bound, no-resampling implementation for opaque PNGs that completely
  cover one PPTX slide, including deterministic profile receipts, source
  mutation checks, bounded no-clobber output, and typed unsupported outcomes
  for every richer layout.
- Added constant-scope inspection of pages from a previously validated native
  PDF inventory, allowing large-document consumers to reuse one complete page
  scan while render-time source and profile checks remain authoritative.
- Added an interactive React playground, type checks, integration tests, and
  Rslib packaging.
- Added GFM task lists, coalesced source-to-visual Markdown updates,
  synchronized split-pane scrolling, and dedicated compatibility tests.
- Added source-aware Markdown ribbon commands and shortcuts plus host-defined
  selected-text menus for both source and visual editing surfaces, with typed
  React, Vue, and Web Component APIs and stale-selection protection.
- Added controlled Markdown source history with typing coalescing, toolbar and
  keyboard undo/redo, selection restoration, and safe rebasing after host
  content replacement without polluting visual-editor history.
- Added a shared bounded Rust Spreadsheet formula parser plus a cancellable
  Worker/WebAssembly scalar calculation kernel with sparse requests,
  deterministic dependency order, cross-sheet references, target-only
  recalculation, bounded dependency depth, JavaScript parity fallback,
  dependency-failure propagation, and ordered cell-scoped Fortune fallback.
- Added `kernelWasmUrl` support for Spreadsheet in React and Vue and the
  matching `kernel-wasm-url` Web Component attribute.
- Added persistent nested Presentation groups with atomic selection and
  collective transforms, plus native PPTX group-node export for slides,
  layouts, and master-derived artwork. Supported group scale is normalized
  across geometry, typography, rich-text runs, and border weights on import;
  rotated or reflected source groups remain an explicit compatibility warning.
- Virtualized Presentation thumbnail scenes in both the slide strip and sorter.
  Every slide keeps a stable keyboard target and scroll footprint, while only
  the selected slide and the viewport overscan mount full scene content.
- Added primary-ribbon Spreadsheet number and percent formats, decimal-place
  controls, and readable percentage defaults in the quarterly-plan template.
- Added direct Presentation playback from the beginning or current slide with
  F5/Shift+F5 shortcuts, automatic fullscreen, and an in-page fallback with an
  explicit exit path.
- Shared the keyboard-accessible table-dimension picker across Document and
  Presentation, with exact row and column creation in one controlled update.
  Desktop keeps the fast 8 × 10 matrix without duplicate table-cell semantics;
  phones use focused row and column controls with 44 px targets and an
  editor-accented, white-text insertion action.
- Kept Word caption numbers and inline cross-references synchronized in the
  live TipTap transaction graph. Deleting or reordering a caption now
  renumbers surviving targets, updates their references, and exposes a visible
  `Missing reference` state for dangling fields instead of leaving a valid-
  looking stale number.
- Kept PDF page navigation, zoom, and history reachable through the compact
  overflow menu while preserving page status in the primary toolbar.
- Moved the phone PDF page-drawer trigger into the toolbar page controls so it
  no longer overlays document content, while retaining modal focus isolation,
  current-page synchronization, and focus restoration after selection.
- Removed internal implementation terminology from the Spreadsheet
  conditional-format manager.
- Added a visible prompt to the empty Markdown source pane without changing its
  controlled content.
- Added Spreadsheet font-family, vertical-alignment, and text-wrap commands to
  the primary ribbon through the native Fortune cell-format model.
- Added an editor-owned Spreadsheet Find bar with Cmd/Ctrl+F interception,
  displayed-value, raw-value, formula, and sparse-cell matching, deterministic
  cell navigation, repeated-shortcut refocus, and grid-focus restoration.
- Kept the phone Spreadsheet Find bar inside the viewport with a 40 px input
  and 40 px previous, next, and close actions, while preserving exact result
  navigation and grid-focus restoration after Escape.
- Made Spreadsheet workbook task panes modal at phone widths, isolated the
  ribbon, grid, and worksheet footer while open, contained forward and reverse
  Tab navigation, and restored the exact ribbon invoker after Escape.
- Made the Presentation chart inspector a modal surface whenever its responsive
  layout overlays the canvas. The close action receives initial focus, the
  ribbon, slide workspace, and status bar remain inert, Tab stays contained,
  dirty fields consume the first Escape, and closing restores the selected
  chart without changing the desktop docked inspector.
- Made Presentation comment review a full-editor modal on phones while keeping
  the desktop review strip docked. The phone surface uses readable review
  typography and touch-sized actions, isolates the ribbon, slide workspace,
  and status bar, contains forward and reverse Tab navigation, lets a dirty
  comment consume Escape before close, and restores the exact New Comment,
  View Comments, or comment-marker invoker.
- Kept common Presentation arrangement commands visible at desktop width by
  compacting group, ungroup, and layer actions without removing their labels
  from accessible names and tooltips.
- Added PDF annotation opacity and compatible stroke-width editing through a
  compact keyboard-accessible style popover, typed capability commands, and
  native PDF annotation defaults and selection updates.
- Rebuilt Presentation transition controls as standard ribbon groups, paged
  compact Office ribbons by complete command groups, and reset stale ribbon
  scroll state when the available width grows.
- Reserved both compact-ribbon navigation edges while tools overflow, bounded
  Word list galleries to the phone viewport, enlarged their numbering controls
  for touch, and restored the TipTap selection after list commands so active
  style and start-value settings remain available when a gallery is reopened.
- Made open popovers explicit editor-shortcut boundaries so Escape closes the
  PDF annotation-style popover without cancelling the selected annotation
  tool.
- Split Word page setup into keyboard-accessible Page, Columns and Sections,
  and Header and Footer tabs so paper controls stay focused and heavyweight
  header/footer editors mount only when requested.
- Stacked the file command bar above non-PDF ribbon tabs at phone widths so
  filenames and actions no longer compress or overlap the keyboard-accessible,
  horizontally scrollable tab row; PDF retains its single compact toolbar.
- Replaced the phone Presentation thumbnail rail with a dismissible,
  focus-managed slide drawer so the editing canvas keeps the primary width.
- Unified modal focus boundaries across the compact Office sidebar,
  Presentation slide drawer, and AI assistant. Focus now enters the visible
  surface, wraps on Tab and Shift+Tab, keeps the background inert, lets only
  the top overlay consume Escape, and returns to the invoking control.
- Kept the phone PDF toolbar clear of host file actions and moved secondary
  annotation tools, opacity, stroke width, and deletion into its scrollable,
  keyboard-operated overflow menu.
- Matched registered document font weights using the CSS Fonts search order so
  common 680/730 heading weights and browser-synthesized bold metrics stay on
  the deterministic Rust/WASM text-layout path instead of falling back to DOM
  line measurement.
- Unified Word body typography, headings, paragraphs, lists, quotations, and
  image wrapping across editing, read-only preview, and PDF composition.
- Positioned preview and PDF headers and footers inside the configured page
  margins without shifting body content; empty headers are no longer rendered
  and PDF composition no longer inserts the filename as an implicit header.
- Preserved physical Word page width and margins in compact preview so a
  narrow viewport scrolls the page instead of changing line wrapping.
- Kept the canonical TipTap document and its Worker/WASM page decorations
  mounted when switching between Word editing and read-only preview, preserving
  page count, automatic breaks, font shaping, and table pagination instead of
  rebuilding a separate HTML preview.
- Corrected page-layout geometry so headers and footers overlay their physical
  top and bottom margins without reducing the body height a second time; the
  kernel protocol is now version 15.
- Made browser Word PDF export consume the mounted TipTap and Worker/WASM
  pagination surface through a stable `artifactId` across React, Vue, and Web
  Components. Export now preserves automatic breaks, shaped runs, table
  continuations, page geometry, and page chrome while capturing long documents
  in bounded batches; the Playground exposes both DOCX and PDF from one compact
  header export menu.
- Kept Word task-pane headers stretched to the pane edges at compact and phone
  widths so titles and close actions retain the shared Office alignment.
- Kept a Word picture selected while its alternative text is edited and
  returned keyboard focus to the exact Picture ribbon command after save or
  cancel, preventing the next key press from changing document content.
- Positioned the Word comment composer against its selected text before the
  browser's first paint, removing the initial jump from the top of the review
  rail.
- Kept Word revision review keyboard focus on the matching action for the next
  change, then returned it to the document after the final individual decision.
- Made the empty Word revision pane reflect whether new changes are actually
  being recorded and added in-pane start/stop controls, keeping the phone modal
  workflow actionable and focus-stable after the last revision is resolved.
- Highlighted all Word Find matches persistently, made initial forward and
  backward navigation select the expected match, and kept Find and Replace
  controls focused across repeated keyboard actions.
- Extended the Word navigation pane from heading-only filtering to contextual
  full-text results with persistent match highlights and exact selection
  jumps. Compact navigation now closes before restoring body focus and the
  selected range.
- Kept new Word comment composers inside the visible review rail for long
  selections and limited discard prompts to comment drafts or replies with
  written content.
- Returned keyboard focus to the unfinished Word comment, reply, or citation
  field when a user cancels closing or switching its task pane, preserving the
  draft and the user's editing position across desktop and compact layouts.
- Preserved unfinished Word replies and edited citation fields when users
  cancel comment deletion, citation deletion, or an internal citation switch,
  and returned focus to the exact field instead of the destructive action.
- Moved Word citation validation beside the invalid tag or title field, added
  accessible error relationships and invalid styling, and focused the field so
  phone users can repair and complete the citation workflow immediately.
- Guarded worksheet deletion with the shared safe-default Office confirmation
  dialog, and kept invalid inline renames open with concise visible and
  accessible validation instead of silently discarding the entered name.
- Enlarged the shared Office color palette into an eight-column phone layout,
  kept the panel inside the viewport, and made vertical keyboard navigation
  follow the rendered grid across theme and standard colors.
- Stopped delayed Spreadsheet grid-focus recovery from stealing focus after a
  letter starts cell editing, preserving direct text entry after F2 and Escape.
- Kept multi-cell paste selection mutable across Fortune Sheet and React state
  replays, preventing a frozen-range crash while preserving the pasted range
  for subsequent copy, cut, and undo commands.
- Released the Presentation slide drawer's modal focus isolation as soon as a
  responsive viewport returns to desktop width, while preserving the open rail
  and moving focus to the active slide.
- Stabilized Presentation focus across React commits after cutting the focused
  object, keeping the active slide inside the editor shortcut scope so an
  immediate paste restores and focuses the clipboard object.
- Moved phone PDF search and page navigation to a dedicated second toolbar row
  so clearing a query cannot collide with the host download action.
- Replaced Playground AI implementation snippets with concise, file-specific
  guidance for documents, Markdown, spreadsheets, presentations, and PDFs.
