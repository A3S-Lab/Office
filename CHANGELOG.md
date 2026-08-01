# Changelog

All notable changes to A3S Office will be documented in this file.

## 0.1.0

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
- Kept PDF page navigation, zoom, and history reachable through the compact
  overflow menu while preserving page status in the primary toolbar.
- Removed internal implementation terminology from the Spreadsheet
  conditional-format manager.
- Added a visible prompt to the empty Markdown source pane without changing its
  controlled content.
- Added Spreadsheet font-family, vertical-alignment, and text-wrap commands to
  the primary ribbon through the native Fortune cell-format model.
- Added an editor-owned Spreadsheet Find bar with Cmd/Ctrl+F interception,
  displayed-value, raw-value, formula, and sparse-cell matching, deterministic
  cell navigation, repeated-shortcut refocus, and grid-focus restoration.
- Kept common Presentation arrangement commands visible at desktop width by
  compacting group, ungroup, and layer actions without removing their labels
  from accessible names and tooltips.
- Added PDF annotation opacity and compatible stroke-width editing through a
  compact keyboard-accessible style popover, typed capability commands, and
  native PDF annotation defaults and selection updates.
- Rebuilt Presentation transition controls as standard ribbon groups, paged
  compact Office ribbons by complete command groups, and reset stale ribbon
  scroll state when the available width grows.
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
- Positioned the Word comment composer against its selected text before the
  browser's first paint, removing the initial jump from the top of the review
  rail.
- Kept Word revision review keyboard focus on the matching action for the next
  change, then returned it to the document after the final individual decision.
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
