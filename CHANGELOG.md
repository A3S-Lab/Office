# Changelog

All notable changes to A3S Office will be documented in this file.

## 0.1.0

- Extracted the complete A3S Web document, spreadsheet, presentation, and PDF
  editor engine into an independent package.
- Added React, Vue 3, Web Component, and framework-free Core entry points.
- Added DOCX, XLSX, PPTX, PDF, HTML, Markdown, text, CSV, XLS, and ODS file
  workflows.
- Added a colocated PDFium WebAssembly asset with an overridable URL.
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
- Kept common Presentation arrangement commands visible at desktop width by
  compacting group, ungroup, and layer actions without removing their labels
  from accessible names and tooltips.
- Added PDF annotation opacity and compatible stroke-width editing through a
  compact keyboard-accessible style popover, typed capability commands, and
  native PDF annotation defaults and selection updates.
- Split Word page setup into keyboard-accessible Page, Columns and Sections,
  and Header and Footer tabs so paper controls stay focused and heavyweight
  header/footer editors mount only when requested.
- Stacked the file command bar above non-PDF ribbon tabs at phone widths so
  filenames and actions no longer compress or overlap the keyboard-accessible,
  horizontally scrollable tab row; PDF retains its single compact toolbar.
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
