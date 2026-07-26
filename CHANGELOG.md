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
