# Changelog

All notable changes to A3S Office will be documented in this file.

## Unreleased

- Added native Writer pair kerning as an exact `w:kern` threshold from 0
  through 3,277 half-points. Zero enables kerning for every font size, positive
  values apply when effective `w:sz` meets or exceeds the threshold, missing
  values inherit, and an entirely absent hierarchy remains disabled. Body,
  headers, footers, footnotes, endnotes, styles, Format Painter, formatting
  revisions, DOCX export/reopen, and prior `w:rPrChange` values retain it.
- Extended the shared `Cmd/Ctrl+D` advanced font dialog with an independently
  mixed Kerning setting, exact 0.5-point validation, a 12-point authoring
  default, effective live preview, explicit direct-format clearing, captured
  selection restoration, and one-step Undo. No dedicated shortcut is claimed.
- Added collision-safe cross-story run-format import, strict namespace and
  malformed-input diagnostics, exact explicit-zero package patching, effective
  Worker/WASM layout state, responsive Playwright coverage, and a local-only
  A3S Test 1.0.0 ACL. GitHub Actions and Pages do not install or invoke A3S
  Test.
- Added native Writer horizontal character scaling as an exact `w:w` integer
  from 1% through 600%. Missing values still inherit, explicit 100% remains a
  direct reset, and an empty `<w:w/>` imports as the native 100% default.
  Body text, headers, footers, styles, Format Painter, tracked formatting,
  reject restoration, DOCX export/reopen, and `w:rPrChange` retain the value.
- Extended the shared `Cmd/Ctrl+D` advanced font dialog with an exact Scale
  field, independent mixed-selection state, validation and live `font-stretch`
  preview. Scale, spacing, and baseline position commit through one TipTap
  transaction and one Undo record; no dedicated character-scale shortcut is
  claimed.
- Added fail-closed namespace, duplicate, child-content, malformed, fractional,
  and range coverage, browser-measurement fallback evidence, responsive
  Playwright coverage, and a local-only A3S Test 1.0.0 ACL. GitHub Actions and
  Pages do not install or invoke A3S Test.
- Added native Writer character baseline position as a signed `w:position`
  value with explicit zero, exact strict measures, inherited style resolution,
  and body/header/footer DOCX import, export, and reopen. Numeric CSS
  `vertical-align` projects the effect while subscript or superscript keeps
  native precedence.
- Extended the shared `Cmd/Ctrl+D` advanced font dialog with Normal, Raised,
  and Lowered position modes, independent mixed-selection handling, live
  preview, half-point validation, captured-selection restoration, and one
  transaction when spacing and position change together. Format Painter,
  tracked formatting, reject restoration, and native `w:rPrChange` retain the
  exact value.
- Added fail-closed namespace, duplicate, malformed, and range coverage plus
  focused Rstest, desktop/compact Playwright, and a local-only A3S Test 1.0.0
  ACL. GitHub Actions and Pages do not install or invoke A3S Test.

## 0.26.0 - 2026-08-22

- Added bounded Spreadsheet partial-run formatting while editing a cell or the
  formula bar. A non-collapsed text selection accepts font family, size, color,
  bold, italic, strikethrough, and underline commands from the Home ribbon,
  converts plain strings to native inline-string runs when needed, and restores
  the text selection after direct ribbon actions.
- Preserved native XLSX rich-text runs through direct formula-bar and F2
  insertion or deletion. Exact authenticated cell operations reconcile one
  contiguous replacement against the controlled source, retain untouched run
  boundaries and semantic colors, and keep one host revision and one Undo
  record.
- Added authenticated formatted HTML paste for formula-bar and F2 selections.
  The one-shot authority binds the exact worksheet, coordinate, controlled
  source, UTF-16 selection, plain clipboard text, and sanitized supported font
  runs. Plain and empty cells may become native inline strings only after the
  emitted replacement matches that proof.
- Kept every authoring path fail closed at 32,767 characters and 512 result
  runs; formatted clipboard HTML is additionally capped at 256,000 characters.
  Focused Rstest, desktop and compact real-clipboard Playwright, and four
  local-only A3S Test 1.0.0 scenarios cover formatting, direct editing, paste,
  one-step Undo, accessibility, and empty console/page-error diagnostics.
  GitHub Actions and Pages do not install or invoke A3S Test.

## 0.25.0 - 2026-08-22

- Added bounded native XLSX rich-text cell import and export for shared strings
  and inline strings. Ordered runs preserve font family, size, RGB color,
  bold, italic, strikethrough, single/double/accounting underline variants,
  leading or trailing whitespace, and matching theme, indexed, automatic, or
  tint color identities across export and reopen.
- Kept Fortune's existing `ct.s` inline-string model as the single browser
  representation. Cell-wide Format Cells font changes immutably update every
  run through one controlled workbook revision and one Undo record, while
  number-format changes retain `ct.t='inlineStr'`. Edited semantic colors fall
  back to explicit RGB rather than exporting a stale source identity.
- Added fail-closed limits of 32,767 characters and 512 runs per cell, 10,000
  rich-text cells, and 100,000 materialized runs. Focused Rstest,
  import-export-reopen fixtures, desktop/compact Playwright, and a local-only
  A3S Test 1.0.0 ACL cover rendering, formatting, focus restoration,
  accessibility, and clean browser diagnostics. GitHub Actions and Pages do
  not install or invoke A3S Test.

## 0.24.0 - 2026-08-22

- Added editor-scoped Spreadsheet `Cmd/Ctrl+Shift+F` and
  `Cmd/Ctrl+Shift+P` aliases. Both reuse the existing Format Cells command and
  captured selection, open the Font tab, and focus font family or font size
  respectively. The controls expose catalog-owned `aria-keyshortcuts`, and
  Escape restores the exact worksheet grid node.
- Kept host inputs, the formula bar, active cell editing, composing or repeated
  events, and existing dialogs outside shortcut ownership. Runtime validation
  rejects malformed or contradictory tab/focus routes before creating a
  dialog, without introducing another formatting model or Apply path.
- Made the native-style font preview readable for white and other light text by
  preserving the requested text color while choosing a dark or light neutral
  canvas from relative luminance. Focused Rstest, desktop/compact Playwright,
  and a local-only A3S Test 1.0.0 ACL cover routing, initial focus, contrast,
  focus restoration, accessibility, and empty browser diagnostics. GitHub
  Actions and Pages do not install or invoke A3S Test.

## 0.23.0 - 2026-08-22

- Added grid-scoped Spreadsheet `Ctrl+'` to copy the exact formula text from
  the active cell's neighbor above without translating relative references.
  When the source is a constant, the same command copies its scalar value.
  Added `Ctrl+Shift+'` to copy only the source's calculated or displayed value
  and remove an existing target formula.
- Kept each gesture limited to the active cell while preserving the broader
  selection and every target style, including emphasis, fill, and number
  format. One native `setCellValuesByRange` batch produces one controlled
  update and one Undo record, and the formula bar synchronizes on a best-effort
  path after the mutation succeeds.
- Added fail-closed guards for top-row, bounds, read-only, inactive, protected,
  merged, pivot, array, dynamic-array, data-table, external, and malformed
  formula cases. Focused Rstest, desktop/compact Playwright, and a local-only
  A3S Test 1.0.0 ACL cover exact formula/value behavior, target-style
  preservation, revision counts, one-step Undo, accessibility, and empty
  browser diagnostics. GitHub Actions and Pages do not install or invoke A3S
  Test.

## 0.22.0 - 2026-08-22

- Added static Traditional Office/Excel Spreadsheet current-date (`Ctrl+;`) and
  current-time (`Ctrl+Shift+;`) entry through a discoverable Home and Number
  menu. Date uses
  the local calendar day and normalized Excel 1900 serial; time uses local
  hour and minute and discards seconds and milliseconds.
- Kept each gesture scoped to the active cell while preserving a broader
  selection. One Fortune batch writes the scalar through the native value API,
  applies only the `yyyy-MM-dd` or `hh:mm` number format, clears an old formula,
  preserves unrelated styles, and creates one Undo record. Read-only,
  inactive, protected, locked, merged, pivot, and out-of-bounds targets fail
  before mutation.
- Added a dedicated Playground workbook with revision evidence, focused Rstest
  and Ribbon coverage, desktop/compact Playwright with a fixed local clock, and
  a local-only A3S Test 1.0.0 ACL with accessibility and clean browser
  diagnostics. GitHub Actions and Pages do not install or invoke A3S Test.

## 0.21.0 - 2026-08-22

- Preserved OOXML Spreadsheet `diagonalDown` and `diagonalUp` borders as
  independent directions across import, Ribbon and Format Cells editing,
  Paste Special, built-in style resolution, Yjs collaboration, XLSX export,
  and reopen. Crossed borders render both directions through the visible-cell
  Canvas hook, legacy Fortune `border-slash` records remain diagonal-down, and
  later native slash or no-border ranges retain their ordered override
  semantics after external workbook ingestion.
- Kept diagonal-border writes bounded to 4,096 cells and replaced repeated
  maximum-range border scans with one indexed pass. Focused model, Canvas,
  XLSX, collaboration, desktop/compact Playwright, and local-only A3S Test
  1.0.0 regressions cover editing, Undo/Redo, row insertion, rendering, clean
  browser diagnostics, and exact round trips; GitHub Actions and Pages do not
  install or invoke A3S Test.

## 0.20.0 - 2026-08-22

- Added Traditional Office/Excel `Ctrl+2`, `Ctrl+3`, and `Ctrl+4` aliases for
  Spreadsheet Bold, Italic, and Underline. The shared command catalog now owns visible
  shortcut copy, `aria-keyshortcuts`, editor keymaps, and the typed formatting
  path across desktop and compact Ribbon layouts.
- Added Automatic Color and No Fill actions to the Spreadsheet font and fill
  palettes. They remove direct Fortune `fc` and `bg` properties instead of
  writing replacement RGB values, preserve one-step Undo, and return the live
  command path to the worksheet grid.
- Preserved imported XLSX theme, indexed, automatic, and tint identity for
  direct font, solid-fill, and border colors across unrelated style edits,
  Yjs collaboration transport, export, and reopen. Export reconstructs the
  referenced palette slots only while rendered RGB and source palettes agree;
  changed values or conflicting slots degrade safely to explicit RGB.
- Added focused Rstest and real XLSX import/edit/export/reopen coverage,
  desktop and compact Playwright regression, and a local A3S Test 1.0.0 suite
  with empty console and page-error evidence. GitHub Actions and Pages continue
  to exclude A3S Test.
- Kept the documentation version picker inside the viewport with bounded
  scrolling as the frozen release list grows.

## 0.19.0 - 2026-08-22

- Added a keyboard-accessible Traditional Office Spreadsheet Text Orientation menu with
  horizontal, counterclockwise, clockwise, stacked vertical, rotate-up, and
  rotate-down choices. Fortune `rt` and `tr='3'` now round-trip the complete
  native XLSX 0–180 and 255 mapping without reversing negative angles.
- Expanded Home and Cells Rows and Columns with row/column Hide and Unhide plus
  grid-scoped `Cmd/Ctrl+9`, `Cmd/Ctrl+0`, `Cmd/Ctrl+Shift+9`, and
  `Cmd/Ctrl+Shift+0`, with exact focus ownership and read-only exclusions.
- Bounded text-orientation mutation to 10,000 cells and visibility mutation to
  10,000 rows or 1,000 columns before native calls or index allocation. One
  orientation intent remains one Fortune batch and one Undo record.
- Split the Spreadsheet controller and Ribbon into focused selection,
  navigation, structure, keyboard, alignment, clipboard, editing, view, and
  row/column modules without changing the public editor API.
- Added focused Rstest, XLSX, desktop/compact Playwright Canvas, accessibility,
  console, page-error, and local A3S Test 1.0.0 regression coverage. GitHub
  Actions and Pages continue to exclude A3S Test.

## 0.18.0 - 2026-08-22

- Added Office-style Spreadsheet advanced underline formatting with a split Home
  and Font control for none, single, double, single-accounting, and
  double-accounting styles. Format Cells preserves mixed selections and exposes
  the same five states; `Cmd/Ctrl+U` turns any active underline variant off and
  restores single underline from the off state through one controlled update
  and one Undo record. Native XLSX import and export now retain the exact OOXML
  underline value instead of collapsing every variant to single, and SheetJS
  fallback imports normalize the corresponding Boolean, numeric, and named
  values. Focused Rstest, desktop and compact Playwright, and a local-only A3S
  Test 1.0.0 regression cover the workflow with empty console and page-error
  diagnostics; GitHub Actions does not install or invoke A3S Test.

## 0.17.0 - 2026-08-22

- Added Office-style Spreadsheet Grow Font and Shrink Font commands to Home and
  Font with `Cmd/Ctrl+Shift+.` / `Cmd/Ctrl+]` and
  `Cmd/Ctrl+Shift+,` / `Cmd/Ctrl+[` aliases. Mixed selections step each cell
  independently through the shared 9–72 point scale, compact equal results
  into native rectangles, and commit one bounded Fortune batch and one Undo
  record. Added grid-scoped `Cmd/Ctrl+Shift+&` Outside Borders and
  `Cmd/Ctrl+Shift+_` Clear Borders commands with visible, accessible shortcut
  metadata. Host inputs, formula or cell editing, and modal controls keep their
  native keys. Focused Rstest, desktop and compact Playwright, and a local-only
  A3S Test 1.0.0 regression cover the workflow; GitHub Actions does not install
  or invoke A3S Test.

## 0.16.0 - 2026-08-22

- Added native Spreadsheet Tables/ListObjects through Insert and the
  grid-scoped `Cmd/Ctrl+T` shortcut. The bounded creation dialog captures a
  current region or explicit range, canonicalizes unique headers, and rejects
  merges, protection, overlapping tables or worksheet AutoFilters, pivots,
  invalid bounds, and selections above 100,000 cells before one controlled
  update and one-step Undo. A contextual Table Design ribbon exposes table
  names, 60 OOXML Light/Medium/Dark styles, first/last-column emphasis,
  row/column stripes, exact style-trigger focus restoration, and sparse-safe
  Convert to Range. Visible-range Canvas painting retains cell ownership while
  row/column reconciliation keeps ranges, columns, filters, and headers
  coherent. Native XLSX table parts, relationships, content types, built-in
  styles, and supported filters round-trip, and ordered ID-keyed Yjs records
  converge independent two-client design edits. Focused Rstest, desktop and
  compact Playwright, and a local-only A3S Test 1.0.0 regression cover the
  workflow; GitHub Actions does not install or invoke A3S Test. Structured
  references, calculated columns, complete totals authoring, slicers, and
  external/query tables remain explicit gaps.

## 0.15.0 - 2026-08-21

- Added Office-style Spreadsheet Data Validation under Data and Data Tools for
  lists, whole numbers, decimals, dates, and text length across one or more
  selected ranges. The accessible dialog supports input messages, explicit
  invalid-input blocking, compact Apply and Remove, exact focus restoration,
  and one-step Undo without materializing blank cells. Validation rejects
  protected, merged, pivot, read-only, out-of-bounds, malformed, and
  over-10,000-cell targets before mutation. XLSX decimal rules now accept
  integers correctly, and date boundaries normalize across the 1900 and 1904
  date systems before stable `DATE(...)` export. Focused Rstest, desktop and
  compact Playwright, and a local-only A3S Test 1.0.0 regression cover the
  workflow; GitHub Actions does not install or invoke A3S Test.
- Added Office-style Spreadsheet Insert/Edit Hyperlink under Insert and Links,
  including the grid-scoped `Cmd/Ctrl+K` shortcut and accessible Web page,
  cell-range, and worksheet targets. One immutable controlled update preserves
  dense or sparse storage, cell values, formulas, formatting, comments, and
  unrelated vendor records; Remove and one-step Undo use the same typed path.
  Unsafe URLs, missing or hidden sheets, invalid or out-of-bounds ranges,
  protected cells, pivots, read-only views, host inputs, and modal surfaces
  fail before mutation. Desktop and compact Playwright coverage is paired with
  a local-only A3S Test 1.0.0 gate using agent-browser 0.26.0 and protocol 15;
  GitHub Actions does not install or invoke A3S Test.

## 0.14.0 - 2026-08-21

- Added Office-style Spreadsheet Paste Special with a split Paste command,
  `Cmd/Ctrl+Alt+V`, ten content modes, four arithmetic operations, Skip
  blanks, and Transpose. Same-editor copies retain formulas, native styles,
  comments, validation, protection, hyperlinks, borders, merges, and column
  widths; external tabular text remains a bounded TSV fallback. Formula
  references translate at the destination, one paste creates one controlled
  update and Undo record, and unsafe merge, pivot, protection, bounds,
  unsupported-formula, and divide-by-zero cases fail before mutation.
- Replaced Spreadsheet's standalone Home Find action with a keyboard-accessible
  Traditional Office Find and Select menu. Find retains `Cmd/Ctrl+F`; Go To
  adds `Ctrl+G` and
  `F5` for bounded direct, quoted cross-sheet, and named ranges while keeping
  host inputs and editing surfaces untouched. Navigation changes only the live
  worksheet view and selection, so sparse data, controlled content, and Undo
  history remain unchanged.
- Added the Traditional Office Spreadsheet AutoSum split command before Fill
  with Sum, Average, Count, Maximum, and Minimum plus the editor-scoped `Alt+=` Sum
  shortcut. It infers contiguous sources above or left, supports multi-formula
  totals rows and columns, preserves target styles and sparse worksheets,
  rejects unsafe targets, and commits one bounded native batch and Undo record.
- Added the complete Office-style Format Cells dialog under Home and Number with
  Number, Alignment, Font, Border, Fill, and Protection tabs and the
  editor-scoped `Cmd/Ctrl+1` shortcut. Mixed values, dense/sparse representation,
  one controlled update, one-step Undo, and exact focus restoration are
  preserved across the bounded workflow.
- Made Spreadsheet Increase Decimal and Decrease Decimal first-class
  Traditional Office Home and Number commands. Mixed selections now retain each
  cell's currency,
  accounting, percentage, number, or scientific format family, compact equal
  results into one bounded native batch, and leave date, time, fraction, text,
  and unknown custom formats unchanged.

## 0.13.1 - 2026-08-20

- Added explicit Traditional Office Spreadsheet presets for General, Number,
  CNY Currency, Accounting, Percentage, Short Date, Time, Scientific, Fraction,
  and Text,
  plus the seven standard `Cmd/Ctrl+Shift` formatting shortcuts. Each command
  preserves values and formulas, creates one Undo entry, and retains its exact
  XLSX number-format code.
- Added 17 grouped Traditional Office built-in Spreadsheet cell styles with
  native previews, two-dimensional keyboard navigation, current-style
  recognition, bounded
  blank-cell application, per-cell border semantics, and one-step Undo. Direct
  XLSX font, fill, alignment, wrapping, rotation, border, and number-format XF
  state now round-trips, including RGB normalization for theme, indexed, and
  tint colors.
- Expanded focused unit, component, desktop, compact, accessibility, XLSX, and
  local ACL coverage. A3S Test remains a local release gate; GitHub CI and
  Pages continue to use Rstest and Playwright without installing or invoking
  A3S Test.
- Made native collaboration-session JSONL input cancellation-safe so periodic
  polling cannot discard a partially received host record or parse its suffix
  as a new command.

## 0.13.0 - 2026-08-20

- Reworked large Document and Spreadsheet editing around bounded windowing,
  incremental projections, cancellable Workers, and persistent Rust/WASM
  sessions. Cold import, first interactive paint, `Ctrl+End`, repeated edits,
  and controlled publication now avoid rebuilding complete 100,000-unit models.
- Bounded PDF runtime and thumbnail work and Presentation layout, snapping,
  thumbnail, and scene rendering. Production keeps explicit windowing while
  `content-visibility` remains a measured opt-in benchmark mode where browser
  behavior does not provide a reliable product contract.
- Published reproducible 100,000-paragraph, 100,000-table-row, and one-million-
  cell measurements for import, readiness, scrolling, memory, DOM size, and
  repeated edits. The documentation shell now follows the A3S navigation and
  code-rendering system, and every editor uses a centered accessible loading
  state with reduced-motion support.
- Added Office-aligned Spreadsheet single strikethrough, native cell-border
  presets and styles, and Fill Down, Right, Up, and Left. Fill preserves native
  relative-formula, series, style, Undo/Redo, sparse-matrix, protection,
  merge, pivot, read-only, and browser-shortcut semantics, with a
  50,000-target-cell limit.
- Expanded unit, component, Rust/WASM, production-build, bundle-budget,
  desktop/compact Playwright, and local ACL coverage. A repository policy now
  prevents GitHub Actions from installing or invoking A3S Test; CI and Pages
  retain Playwright as their browser regression gate.

## 0.12.0 - 2026-08-18

- Added cancellable browser file imports with monotonic `reading`, `parsing`,
  `analyzing`, and `finalizing` progress. Bounded reads and parser checkpoints
  yield to the host UI, `AbortSignal` cancellation fails with `AbortError`, and
  the Playground now exposes live progress plus an explicit Cancel action.
- Added maximum-dimension XLSX worksheets with 1,048,576 rows and 16,384
  columns while keeping empty ranges sparse. Navigation, search, selection,
  formulas, filters, formatting, statistics, host projection, import, and
  export visit only materialized cells; editing a far blank row creates only
  that row and preserves the complete logical scroll range.
- Preserved large data-validation, protection, passwordless editable, and
  conditional-formatting regions as compact ranges through browser editing and
  native XLSX round trips instead of expanding them into per-cell records.
- Hardened Spreadsheet collaboration and Fortune Sheet projection so derived
  formula metadata and visible-row caches do not create false controlled
  changes, sheet activation remains local view state, system calculations stay
  outside user undo history, and cross-sheet results update their owning sheet.
- Added focused sparse-workbook, import-performance, collaboration, controlled
  projection, vendor-patch, and cross-sheet calculation regressions. The
  deterministic A3S Test workflow proves maximum-cell navigation, one-row
  materialization, persisted editing, accessibility, and clean browser
  diagnostics; bilingual documentation and the Playground expose the complete
  workflow.

## 0.11.0 - 2026-08-18

- Added node-level `paragraph-formatting` revisions for alignment, direction,
  indentation, spacing and line rules, pagination controls, contextual
  spacing, outline level, tab stops, borders, shading, and collapsed state.
  Multi-paragraph commands share one stable identity and original canonical
  snapshot; accept keeps current properties, reject restores the complete old
  properties, and both preserve text with independent undo boundaries.
- Imported the supported strict and transitional DOCX `w:pPrChange` subset as
  Paragraph Formatting cards and exported it back to native OOXML. Validation
  rejects duplicate, malformed, namespace-spoofed, non-canonical, or unsupported
  revisions, keeps them on the structural diagnostic path, and removes private
  transport markers from final DOCX packages.
- Extended browser/Yrs collaboration with immutable
  `changeKind: "paragraph-formatting"` decisions, browser-client convergence,
  restart-safe Rust projection, and suggestion-mode protection that permits
  normal text proposals while rejecting paragraph-revision metadata tampering.
- Expanded the public Playground to show independent character and paragraph
  formatting revisions. Added deterministic A3S Test gates for both rejection
  paths, including text and sibling-revision preservation, paragraph-property
  restoration, screenshots, accessibility, console, and page-error evidence.
  Updated the bilingual Document, collaboration, architecture, roadmap, and
  complete A3S Boot backend guides.

## 0.10.0 - 2026-08-17

- Added character-formatting revisions for bold, italic, underline, strike,
  subscript, superscript, font family, font size, text color, highlight, and
  Word grid state. Accepting keeps the new formatting, rejecting restores the
  exact prior direct marks, and either decision preserves text and remains one
  undoable transaction.
- Imported the supported native DOCX `w:rPrChange` subset as reviewable
  Formatting cards and exported it back as native run-property changes,
  including strict namespaces, missing dates, malformed-input diagnostics, and
  source-backed round trips without private marker leakage.
- Extended browser/Yrs collaboration validation and immutable decision audit
  records with the separate `formatting` change kind. Formatting revisions
  survive restart plus duplicate or reordered delivery; authenticated
  `suggest` sessions preserve existing formatting revisions but cannot forge or
  rewrite them before A3S Boot persistence and broadcast.
- Added a public Playground formatting-revision demo and deterministic A3S Test
  regression covering the Review pane, Formatting label, rejection semantics,
  accessibility, screenshots, console, and page errors. Updated the bilingual
  Document, collaboration, and complete A3S Boot backend guides and froze the
  0.10.0 documentation set.

## 0.9.2 - 2026-08-17

- Added an ephemeral native Yrs Awareness controller for CLI and Rust
  collaboration peers. It publishes active, idle, or away state plus typed
  Document, Markdown, Spreadsheet, Presentation, and PDF locations, projects
  validated remote participants, handles peer departure and reconnect cleanup,
  and emits an orderly disconnect tombstone.
- Separated each Presence-enabled process connection ID from the durable
  replica's stable Yrs author ID. Host envelopes and Awareness use the fresh
  connection identity while canonical CRDT updates retain the replica identity;
  Presence state, clocks, snapshots, and delivery never enter checkpoints,
  update logs, or operation receipts.
- Extended `collab session` with bounded actor profile options and typed
  `set-presence`, `receive-awareness`, and `peer-left` JSONL input alongside
  `outbound-awareness` and sorted `presence` output. Core and real CLI process
  regressions cover browser-compatible Awareness, clock ordering, identity
  rejection, all five location schemas, and persistence isolation.
- Updated the README, roadmap, packaged Office Skill, CLI reference, and
  bilingual A3S Boot collaboration runbook; added deterministic A3S Test
  coverage for the published protocol and froze the 0.9.2 documentation set.

## 0.9.1 - 2026-08-17

- Made live Document pagination measure the canonical editor DOM cooperatively
  in bounded 32 ms slices. Measurement yields only between top-level blocks,
  preserves the synchronous block, section, page-style, and reuse contracts,
  and keeps input, media, and animation work schedulable while large documents
  are repaginated.
- Connected cooperative measurement to the existing single-flight pagination
  coordinator. A document edit aborts the stale pass before it can publish its
  snapshot, then coalesces follow-up work without replacing the controlled
  TipTap model or weakening Worker/WASM layout authority.
- Canonicalized sub-micro-pixel page metrics across the JavaScript and WASM
  boundary so an incrementally reused page prefix also retains its resolved
  headers, footers, page numbers, previews, and navigation descriptors.
- Added multi-section parity and abort regressions, retained the 120-page
  deterministic A3S Test navigation/edit flow, updated the bilingual Document
  and architecture guidance, and froze the 0.9.1 documentation set.

## 0.9.0 - 2026-08-17

- Added `spreadsheet-batch-cells` across Rust, `collab mutate`, standard MCP,
  A3S Code, and the packaged Office Skill. One mutation applies 1 to 4,096
  distinct guarded coordinates in one sheet, mixing recursive set/create and
  exact delete operations while retaining unrelated concurrent JSON leaves.
  Every guard is evaluated against one shared snapshot; a duplicate
  coordinate, invalid delete, stale leaf, or stale exact deletion rejects the
  complete gesture without a Yjs update or durable log entry.
- Refactored Spreadsheet cell mutation state into one validated sheet snapshot
  and one Yrs transaction. Existing `spreadsheet-set-cell` and
  `spreadsheet-delete-cell` now use the same batch kernel, dense dimensions are
  extended once from the final coordinate set, and sparse or initially empty
  sheets remain `celldata`. Core restart/idempotency, atomic-conflict, bounded
  contract, real CLI/MCP subprocess, and browser Yjs duplicate/reordered
  delivery regressions cover the shared path.
- Extended the Playground and deterministic A3S Test flow with a visible
  native atomic batch that updates, creates, and deletes cells in one action,
  followed by the existing single-cell lifecycle. Updated the bilingual
  collaboration, Spreadsheet, CLI/MCP, Skill, README, and roadmap guidance.

## 0.8.1 - 2026-08-17

- Added closed native Document suggestion mutations across Rust, `collab
  mutate`, standard MCP, A3S Code, and the packaged Office Skill.
  `document-suggestion-create` lets an actor-scoped `suggest` replica propose
  an insertion, deletion, or atomic replacement after matching a stable plain
  paragraph, its current text identity, exact UTF-16 range, and selected text.
  `document-suggestion-decide` lets an `edit` replica accept or reject a batch
  of complete projected identities atomically, apply the tracked changes, and
  append immutable browser-compatible actor-attributed decisions. Stale guards,
  overlapping proposals, reused IDs, forged attribution, incomplete replacement
  identities, and conflicting final decisions fail without a durable log entry.
- Promoted the native collaboration projection to schema v3 with live
  suggestions, exact placements, and immutable final decisions. Native restart,
  mode/identity enforcement, idempotent retry, rejection, atomic replacement,
  real CLI/MCP subprocess, and browser Yjs compatibility regressions cover the
  complete propose/accept workflow.
- Extended the Playground with deterministic native Agent proposal and native
  editor acceptance controls over the same browser-readable Yjs fixture, visible
  ready/proposed/accepted state, final text, and decision audit. Updated the
  A3S Test workflow, bilingual collaboration and A3S Boot backend guides, CLI
  reference, Skill, README, and roadmap; froze the 0.8.1 documentation set.
  The documentation shell now matches the current A3S Test navigation height
  and exact logo asset while retaining the A3S UI-aligned Shiki code rendering,
  responsive code scrolling, inline code, and table treatment.

## 0.8.0 - 2026-08-17

- Added authenticated Document suggestion mode across the Core API, React,
  Vue, and Web Components. Suggesters can type attributed insertions, mark
  deletions, and create replacements while canonical text, structure,
  formatting, options, comments, and other actors' proposals remain protected.
  Suggesters can withdraw their own insertion proposals but never accept,
  reject, rewrite, or remove another actor's suggestion.
- Added immutable shared accept/reject audit records under
  `document.change-decisions` and `document.change-decision-order`. Edit-mode
  reviewers apply each final decision atomically with the corresponding tracked
  change, retain proposer and decider actor attribution, deduplicate identical
  disconnected retries, reject stale conflicting decisions, clear unsafe local
  history after a decision, and continue to read protocol-v1 Documents created
  before the additive roots existed.
- Extended the native Yrs authorization boundary and runnable A3S Boot backend
  to admit only semantically valid authenticated Document `suggest` updates
  before persistence or room broadcast. Candidate-state validation permits
  actor-attributed insertion, deletion, and replacement marks while rejecting
  canonical content, structure, non-suggestion formatting, roots, sidecars,
  forged identity/timestamps, foreign proposal changes, and unresolved Yjs
  dependencies. Fixed the canonical projection to merge adjacent equal-format
  text after removing revision marks, so valid deletion and replacement
  proposals are no longer rejected.
- Reworked the Playground suggestion entry into two independent synchronized
  Yjs clients with separate editor and suggester identities, live proposal
  delivery, editor-only decisions, and a visible final audit trail. Added a
  deterministic A3S Test desktop/phone suite covering proposal creation,
  convergence, control separation, acceptance, audit output, responsive
  reachability, accessibility, screenshots, console, and page errors. Updated
  the bilingual collaboration, backend, Document, CLI/MCP, roadmap, and example
  documentation and froze the resulting 0.8.0 documentation set. The CLI/MCP
  reference now explicitly records that native typed suggestion creation and
  decision mutations are not yet exposed.

## 0.7.3 - 2026-08-17

- Added durable Feishu-style Document selection comments for authenticated
  `comment` sessions. Browser participants can select text, create attributable
  threads, reply, resolve or reopen a thread, and delete only comments or
  replies owned by their collaboration actor. Canonical text and every other
  Document option remain read-only; remote review updates stay outside local
  undo history, and detached threads remain available when their marked text is
  removed by an authorized editor.
- Added native Document comment mutations through Rust, `collab mutate`, MCP,
  and the packaged Office Skill. `document-comment-create`,
  `document-comment-reply`, `document-comment-set-resolved`, and
  `document-comment-delete` share browser-compatible stable records, immutable
  claims, ProseMirror selection marks, UTF-16 offsets, actor attribution,
  ownership checks, durable receipts, and idempotent retry behavior. Document
  projection schema v2 now returns comment threads, replies, resolution state,
  detached state, and exact paragraph/text identities plus anchor text and
  offsets.
- Hardened the runnable A3S Boot collaboration backend for review-only
  publishing. Version-2 signed tickets bind the actor display name as well as
  actor ID, kind, room, namespace, mode, and expiration. Under the durable room
  lock, the Yrs store semantically validates every Document `comment` update
  against the previous canonical state and rejects forged content, structure,
  options, roots, authorship, ordering, claims, anchors, or foreign deletion
  before persistence or broadcast. `edit` may publish canonical and review
  updates; Document `comment` may publish validated review updates;
  `view` and `suggest` remain receive-only.
- Reworked the Playground collaboration entry into a two-participant real-time
  review workflow and added A3S Test coverage for selection-comment creation,
  attribution, highlight, reply, resolve/reopen, blocked text editing,
  responsive UI, accessibility, console, and page errors. Updated the
  bilingual guides, CLI/MCP references, backend runbook, and frozen 0.7.3
  documentation while retaining the A3S Test-aligned navigation and the A3S UI
  content/code-highlighting treatment.

## 0.7.2 - 2026-08-15

- Added an Office-owned native projection for initialized Markdown and
  Document collaboration replicas. Rust, `collab read`, and
  `office_collaboration_read` now expose canonical Markdown source or bounded
  Document paragraphs with structure, stable Word identities, current option
  fields, agent-readable text, sequence, and state vector. The new
  `document-replace-paragraph` mutation matches `paragraphId`, current
  `textId`, and complete text before changing one plain paragraph, so browser
  edits invalidate stale agent decisions without replacing the document.
  Restart, reordered delivery, browser edits, CLI, and MCP regressions cover
  the read/mutate path.
- Added a complete runnable A3S Boot 0.2 collaboration backend with typed ACL
  configuration, HMAC room tickets, Origin and permission enforcement,
  standard Yjs sync messages, crash-safe Yrs persistence, Awareness fan-out,
  stale-peer cleanup, deterministic delivery receipts, and reconnect repair.
  The browser adapter, service runbook, and integration tests cover two-client
  broadcast, read-only synchronization, restart recovery, ticket tampering,
  and the repository API envelope.
- Made real-time collaboration visible on the Playground homepage with a
  two-participant Document demo, remote cursor, participant roster, responsive
  phone layout, and a canonical link to the bilingual backend guide. Removed
  the remaining route-dependent Playground rows from every frozen
  documentation homepage. A3S Test regressions now cover both the capability
  entry and the complete documentation navigation path.
- Patched the initial Yjs-to-ProseMirror selection restoration to resolve the
  nearest valid text position inside nested Document sections. Collaboration
  startup no longer emits a `TextSelection` warning in desktop or compact
  browsers.

## 0.7.1 - 2026-08-15

- Made real-time collaboration a first-class, bilingual documentation path.
  The default Simplified Chinese site now includes the complete host-owned
  Yjs/Yrs setup for two-browser editing, Awareness participants and remote
  locations, Document, Markdown, Spreadsheet, Presentation, and PDF bindings,
  CLI/MCP/A3S Code peers, reconnect and offline convergence behavior,
  authorization boundaries, and current limitations. Both documentation
  homepages and component indexes now link directly to collaboration and show
  the five-format capability matrix.
- Removed the route-dependent “在线体验” / “Playground” button injected into
  every documentation header because its generated relative destinations were
  not reliable across localized and versioned routes. Focused source and A3S
  Test navigation regressions cover the resulting documentation contract.

## 0.7.0 - 2026-08-15

- Added conflict-local Presentation scene-element z-order moves through Rust,
  `collab mutate`, standard MCP, and A3S Code. The closed
  `presentation-move-element` mutation addresses slides, masters, or layouts
  by stable element and predecessor IDs instead of array indexes. Callers
  provide the observed predecessor and requested predecessor, with `null`
  meaning the first element-order position. A move already at its destination
  is idempotent; otherwise a stale source position, missing or deleted target
  anchor, deleted element, or self-anchor fails without a durable update. The
  mutation removes and reinserts only the moved element's order entries,
  preserves every scene-object field and the surrounding container, and
  projects concurrent duplicate order entries by first active occurrence.
  Native restart, causal reordering, concurrent field edits, real CLI/MCP
  subprocesses, browser Yjs duplicate/reordered delivery, and the Playground
  `a3s-test` regression cover update/create/move/delete interoperability.

## 0.6.0 - 2026-08-15

- Added native Presentation scene-element collaboration through Rust,
  `collab mutate`, standard MCP, and A3S Code. The closed
  `presentation-create-element`, `presentation-update-element`, and
  `presentation-delete-element` mutations address slides, masters, or layouts
  by stable identity. Creation writes a canonical browser-compatible claim and
  supports deterministic placement after an existing element; top-level
  optimistic field guards merge unrelated concurrent edits while rejecting a
  stale same-field write; deletion requires the complete current element and
  writes a durable tombstone that reserves the ID. Element ID and type remain
  immutable, identical retries are idempotent, conflicting same-ID creation
  fails closed, and legacy browser sessions without complete claims remain
  readable. Browser/Rust interoperability, restart, duplicate and reordered
  delivery, real CLI/MCP subprocesses, atomic conflict failures, and a
  Playground `a3s-test` regression cover the full update/create/delete
  lifecycle.

## 0.5.0 - 2026-08-15

- Added native Spreadsheet collaboration cell mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. `spreadsheet-set-cell` creates
  or recursively patches one browser-compatible cell with optimistic leaf
  guards, while `spreadsheet-delete-cell` requires an exact complete-cell
  match. Unrelated concurrent formula, value, style, hyperlink, note, and
  metadata leaves merge; same-leaf conflicts fail without a durable update.
  Dense sheets retain and safely extend their matrix projection, sparse sheets
  remain sparse, and an empty sheet's first write uses `celldata`. Coordinates,
  JSON depth and size, unsafe object keys, malformed shared roots, and orphaned
  fields are bounded and validated before mutation. Native restart, real CLI
  and MCP subprocesses, browser Yjs duplicate/reordered delivery, and a
  Playground `a3s-test` regression cover set, create, and delete behavior.

## 0.4.0 - 2026-08-15

- Added a versioned, deterministic Document Snapshot codec for lossless
  controlled-value persistence and an agent-readable Markdown Source
  projection for single-section documents. Snapshot decoding validates the
  schema, version, size, structured model, and synchronized HTML fingerprint;
  Source revisions retain Office-owned section layout and reattach only
  unambiguous surviving comment anchors. Public Core exports, README guidance,
  and focused round-trip and fail-closed tests cover both contracts.
- Hardened controlled Document editing and pagination under concurrent host,
  agent, observer, and font updates. External snapshots now apply the smallest
  history-neutral ProseMirror transaction, while a single-flight pagination
  coordinator coalesces observer churn and aborts only invalidated work. The
  Playground now lazy-loads editors, file import, and PDF evidence support from
  lightweight shared contracts, reducing the initial entry bundle from 716.8
  KiB to 71.4 KiB gzip while keeping the full editor available on interaction.
  Adaptive ribbons now derive density from the stable outer viewport so
  overflow navigation cannot trigger a resize oscillation.
- Added native PDF annotation create, optimistic leaf update, and irreversible
  delete mutations through Rust, `collab mutate`, standard MCP, and A3S Code.
  The closed surface accepts portable FreeText, Highlight, Underline,
  StrikeOut, and Ink records; creation writes a browser-compatible `created`
  record plus immutable claim, concurrent updates merge unrelated JSON leaves,
  same-leaf conflicts fail without an update, and deletion writes a durable
  tombstone. Identical retries are no-ops, immutable ID/page/type changes are
  rejected, and created claims remain valid after later mutable edits. Native
  restart, duplicate/reordered-delivery, real CLI/MCP subprocess, and browser
  Yjs interoperability tests cover the lifecycle. Exact native updates now
  project through the EmbedPDF harness with real nested Highlight geometry.
  The stable logical document digest canonicalizes JSON object arrays across
  Yrs restart and delivery order while raw commit detection continues to retain
  causally pending structs. The Playground exposes a deterministic remote
  create/update/delete fixture, with an `a3s-test` ACL regression that captures
  screenshots, accessibility, console, and page-error evidence.
- Added native append-only PDF redaction and page-operation review mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. `pdf-propose-redaction` writes
  bounded page geometry with the replica actor and a canonical UTC timestamp;
  `pdf-propose-page-rotation`, `pdf-propose-page-deletion`, and
  `pdf-propose-page-reorder` write validated source-page subsets or a complete
  permutation without changing source bytes; and
  `pdf-decide-review` writes the single attributable final decision for an
  existing redaction or page operation. Both records and their canonical
  creation claims are committed atomically, identical stable-ID retries are
  no-ops, and conflicting ID reuse, missing targets, duplicate final
  decisions, invalid geometry, unsupported rotations, page-range violations,
  deleting every page, and incomplete reorders fail without a durable update.
  Native restart tests, real CLI/MCP subprocesses, and browser Yjs fixtures
  cover concurrent edits plus duplicate and reordered delivery. An exhaustive
  native convergence test now checks all 24 delivery orders for causally
  related rotation, deletion, reorder, and final-decision updates, including a
  duplicate delivery and durable restart. Native replay treats the immutable
  checkpoint and raw update log as authoritative and canonically rebuilds Yrs
  pending structures before reporting the committed state vector.
- Added the first typed native PDF collaboration mutation through Rust,
  `collab mutate`, standard MCP, and A3S Code. `pdf-set-form-value` updates an
  existing conflict-local form-value leaf or deterministically creates its
  typed presence/fields/order record without synchronizing source bytes.
  Browser-generated Yjs fixtures, native restart/idempotency checks, real CLI
  and MCP subprocesses, and concurrent browser/native replay prove the update
  remains readable and convergent across Yjs and Yrs. Invalid or oversized
  field identities fail before any durable state changes.
- Extended typed native collaboration mutations to Document. Coding agents can
  replace an exact, fail-closed match count inside ProseMirror `Y.XmlText`,
  rotate the affected Word `textId` plus every identified ancestor table row's
  `rowTextId`, and insert or guarded-delete plain paragraphs in bounded section,
  nested list-item, table-cell/header, and blockquote containers without
  replacing the shared XML tree. Page-color and track-changes remain
  conflict-local options. Rust, CLI, MCP, and browser Yjs fixtures cover
  restart, idempotency, stale identity/text rollback, emoji offsets, nested
  tables, concurrent structural edits, and cross-language replay.
- Added idempotent typed native collaboration mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. The initial Markdown
  replace/splice surface writes canonical `Y.Text` with browser-compatible
  UTF-16 offsets, rejects surrogate-splitting or stale ranges, requires edit
  mode, and publishes its incremental update through the existing durable live
  session path.
- Native collaboration receipts now preserve validated browser source
  actor/operation origins separately from host delivery IDs. Attribution
  survives restart, appears in resumable CLI/MCP events, participates in
  idempotency conflicts, and is re-emitted unchanged to other live peers.
- Added a transport-neutral native collaboration session and the machine-only
  `collab session` JSONL bridge. Coding agents can now join a host-owned live
  room with browser-compatible `SyncStep1`/`SyncStep2` reconnect handshakes,
  durable delivery receipts, external CLI/MCP update projection, compaction
  recovery, typed outbound origins, and remote-echo suppression while the host
  retains connectivity, room, authentication, and authorization ownership.
- Added a bounded, identity-bound host transport adapter for Yjs v1
  state-vector/update synchronization with explicit reconnect, echo
  suppression, and typed incremental origins. Added a provider-owned Awareness
  controller with validated participant state and format-specific locations for
  Document, Markdown, Spreadsheet, Presentation, and PDF sessions.
- Live Word pagination now preserves each section's exact page size,
  orientation, margins, and page gap across mixed-layout documents. Kernel
  protocol 16 carries deduplicated page styles into both JS and Rust/WASM,
  returns metrics on every physical page, and drives variable-size page
  sheets, borders, navigation thumbnails, and exact per-page PDF capture.
- Added bounded structured OMML equations for Word documents. Inline and
  display equations now round-trip as native `m:oMath` and `m:oMathPara`
  across the document body, headers, footers, footnotes, and endnotes, with an
  accessible MathML preview and atomic insert/update commands. The structured
  display subset preserves `left`, `right`, `center`, and `centerGroup`
  paragraph justification, canonicalizes absent or empty justification to the
  `centerGroup` default, and requires one optional `m:oMathParaPr` before one
  `m:oMath` child. The structured
  subset covers Unicode runs with literal/normal-text semantics, six math
  scripts, four styles, manual breaks, and alignment points,
  plus ordered `m:rPr -> w:rPr -> m:t/w:t` runs with bounded Word fonts and
  theme references, Latin/complex-script bold and italic flags, all-caps and
  small-caps presentation, single and double strike, outline, shadow, emboss,
  imprint, proofing/grid flags, hidden and web-hidden states, direct/theme
  colors, signed character spacing through
  31,680 twips, 1–600% horizontal scaling, half-point kerning thresholds and
  signed baseline positions, half-point sizes, colored underline styles, all
  seven legacy text-animation values, all 27 line-border styles with
  direct/theme colors, 2–96 eighth-point widths, 0–31 point spacing, and
  explicit shadow/frame flags, every named highlight color, complete patterned
  run shading with direct or theme foreground/background colors, manual run
  widths from 0 through 31,680 twips with optional signed 32-bit grouping IDs,
  explicit baseline/superscript/subscript run alignment, all five Word
  emphasis-mark values (`none`, `dot`, `comma`, `circle`, and `underDot`),
  RTL/complex-script flags, language tags, and East Asian typography metadata
  with optional signed 32-bit run IDs, two-lines-in-one flags, all five
  enclosing-bracket styles, horizontal-in-vertical rotation, and rotated-text
  compression, explicit paragraph-mark always-hidden/reset flags, and Office
  2010 text glow, shadow, reflection, text-outline, and text-fill effects. Glow
  retains an optional 0 through 2,147,483,647 EMU radius, one RGB or 17-slot
  theme color source, and up to 64 ordered, repeatable tint, shade, alpha,
  hue-modulation, saturation, and luminance transforms. The distinct Office 2010
  shadow effect retains the same
  color model plus optional 0 through 2,147,483,647 EMU blur and offset
  coordinates, a direction from 0 inclusive to 360 degrees exclusive, signed
  horizontal and vertical scales, skew angles strictly between -90 and 90
  degrees, and all ten rectangle alignments. Angles retain exact
  1/60,000-degree units and scales retain exact 1/1,000-percent units.
  The leaf Office 2010 reflection effect retains optional blur and distance
  coordinates, start/end opacity and position from 0 through 100 percent,
  direction and fade direction from 0 inclusive to 360 degrees exclusive,
  signed horizontal and vertical scales, skew angles strictly between -90 and
  90 degrees, and the same ten rectangle alignments. Angles retain exact
  1/60,000-degree units, while opacity, positions, and scales retain exact
  1/1,000-percent units.
  The structured Office 2010 text-outline effect retains an optional width from
  0 through 20,116,800 EMUs, three line caps, five compound-line styles, two pen
  alignments, distinct none/solid/gradient fills, optional lists of 2 through 10
  gradient stops with RGB or theme colors and ordered transforms, exact linear
  or path shading with optional signed 32-bit relative fill rectangles, all 11
  preset dashes, and round, bevel, or miter joins with optional exact
  nonnegative limits.
  The Office 2010 text-fill effect reuses the same strict no/solid/gradient fill
  grammar, bounded colors and transforms, 2 through 10-stop lists, linear/path
  shades, signed relative rectangles, and exact units without outline geometry.
  Explicit zero/default geometry values reset inherited formatting instead of
  canonicalizing away. Strict universal font-size and position measures enter
  the model only when they convert exactly to the bounded half-point form.
  Strict universal manual widths are accepted only when they convert exactly
  to bounded whole twips; omitted grouping IDs remain distinct from explicit
  zero. Explicit baseline alignment remains present so inherited superscript or
  subscript formatting can be reset, and explicit `none` emphasis remains
  present so inherited emphasis marks can be removed.
  Empty `w:eastAsianLayout` elements canonicalize away, while omitted flags
  stay distinct from explicit `false` resets and signed run IDs retain explicit
  zero. An empty `w:specVanish` canonicalizes to `true`; omission remains
  distinct from an explicit `false` inheritance reset.
  Omitted `w14:glow/@w14:rad` keeps its zero schema default while explicit zero
  remains present. Export declares `w14` and merges it into `mc:Ignorable`.
  Omitted `w14:shadow` geometry keeps its zero/`none` defaults while explicit
  zero and `none` remain present; the effect remains distinct from legacy
  `w:shadow` on/off formatting.
  A present empty `w14:reflection` remains distinct from omission. Its omitted
  geometry keeps zero/`none` defaults while explicit zero and `none` remain
  present.
  A present empty `w14:textOutline` remains distinct from omission and keeps
  the schema's bevel default. Omitted fill, dash, and join choices retain their
  defaults, while explicit zero/default attributes and empty child choices
  remain present.
  A missing `w14:textFill` continues to use `w:color`; a present empty wrapper,
  empty solid fill, or gradient without a stop list retains the distinct black
  schema default.
  All-caps and small-caps presentation, character spacing, width scaling,
  effective kerning, baseline shifts, baseline/superscript/subscript alignment,
  and exact transform-free Office 2010 RGB text fills or black fill defaults use
  safe MathML/CSS projections without changing source Unicode text. Word
  emphasis marks project through CSS as filled dots,
  a literal comma, or an open circle above the text, or a filled dot below it.
  Superscript and subscript also project the smaller rendered size required by
  Word. When `w:position` and `w:vertAlign` coexist, both remain in native
  schema order and the later explicit alignment controls the CSS vertical
  position. `w:em` remains after `w:rtl`/`w:cs` and before `w:lang`, while
  `w:eastAsianLayout` remains after `w:lang`, `w:specVanish` follows it,
  `w14:glow` follows `w:specVanish`, `w14:shadow` follows `w14:glow`, and
  `w14:reflection` follows `w14:shadow`, followed by `w14:textOutline` and then
  `w14:textFill`. Simple
  explicitly sized solid, double, dotted, dashed, inset, and outset line
  borders project through CSS with direct or automatic color and point
  padding; explicit `nil`/`none` resets also project. Relief effects, legacy
  text animations, complex
  multi-line, wavy, or 3D line borders, border shadow/frame, theme-only border
  colors, and hidden or web-hidden states remain native metadata because Word
  view and rendering settings govern them. Manual run widths also remain
  native-only because Word ignores `w:fitText` inside Office Math, so the
  MathML preview deliberately does not emulate them. East Asian
  two-lines-in-one, enclosing brackets, horizontal-in-vertical rotation, and
  rotated-text compression also remain native-only because CSS writing modes,
  text combination, and transforms cannot reproduce Word's inline line-box
  semantics without layout drift. `w:specVanish` also stays native-only and
  never hides equation previews because its display semantics apply only to
  paragraph marks; Word additionally ignores it unless `w:vanish` is set.
  Schema-valid values remain preserved without inventing that dependency.
  Office 2010 glow, shadow, reflection, and text-outline effects also remain
  native-only because CSS `text-shadow`, reflection, opacity, transform,
  text-stroke, paint-order, and border approximations cannot preserve theme
  colors, ordered transforms, exact blur and offset coordinates, reflection
  opacity/position/fade geometry, signed scale/skew, rectangle alignment,
  gradient or compound strokes, preset dashes, caps, joins, or pen alignment.
  Text `noFill`, theme or transformed text-fill colors, and nonempty gradients
  remain native-only; previews keep readable fallback color instead of using
  fragile transparent-text or background-clipped-gradient approximations.
  Highlight precedence over shading is
  retained. Named highlights, explicit highlight removal, clear direct fills,
  solid direct foregrounds, and nil shading project through MathML
  `mathbackground`; pattern masks and theme-only colors remain native metadata.
  Explicit on/off values survive native regeneration. Enabled mutually
  exclusive casing, strike, or relief combinations, invalid animation values,
  art-border styles, out-of-range border width/spacing, malformed border
  colors/flags, missing, malformed, fractional, or out-of-range manual widths
  and grouping IDs, missing or unknown vertical-alignment or emphasis-mark
  values, malformed or out-of-range East Asian layout IDs, flags, or bracket
  styles, malformed paragraph-mark visibility flags, malformed glow radii,
  shadow or reflection geometry, text-outline fill/gradient/dash/join
  structure, text-fill wrapper/fill/gradient structure, color choices, or
  transform chains, and
  unknown, reordered,
  duplicated, spoofed, or
  relationship-bound Word run properties fail closed.
  Supported object property containers also preserve one optional ordered
  `m:ctrlPr` control format through that bounded property model. The control
  may contain a direct `w:rPr` or tracked `w:ins`, `w:del`, `w:moveFrom`, or
  `w:moveTo` provenance with bounded IDs/authors, optional validated core dates
  and Microsoft 365 `w16du:dateUtc` values with a UTC `Z` suffix,
  Word-legal `moveFrom/moveTo -> ins/del` and `ins -> del` nesting, and an
  optional deepest `w:rPr`. Empty direct control properties canonicalize away;
  empty revisions remain native provenance. Every supported `deg`, `den`, `e`,
  `fName`, `lim`, `num`, `sub`, and `sup` argument slot now retains the same
  direct or revision-wrapped control format after its expressions. Fixed slots
  use named metadata; matrix cells, equation-array rows, and delimiter arguments
  use strictly dimension-aligned metadata. Argument formatting, revision
  provenance, and document-level move-range pairing remain native-only because
  they are absent from professional MathML. Unknown, malformed, illegally
  nested, spoofed, or relationship-bound control markup fails closed. Safe
  object-control values project only onto separable MathML control/operator
  nodes while all supported values remain in native metadata.
  Matrix properties now preserve the ordered
  `baseJc -> plcHide -> rSpRule -> cGpRule -> rSp -> cSp -> cGp -> mcs -> ctrlPr`
  grammar. The five spacing rules, unsigned-short row/gap values, and minimum
  column widths through 31,680 twips round-trip with attribute-free Word
  defaults. Row and column gaps project to safe MathML table spacing, while
  minimum width remains native-only for layout because MathML exposes fixed
  column width.
  N-ary `grow` and delimiter `grow`/`shp` properties now round-trip with their
  distinct object defaults and attribute-free enabled values. Growing n-ary
  operators project to MathML `stretchy`; fixed delimiters project with
  `stretchy=false`, and content-matched growing delimiters use
  `symmetric=false`. Native export retains schema order and canonicalizes
  default non-growing n-aries and growing centered delimiters.
  bar/no-bar/skewed/linear fractions, super- and subscripts, aligned right-side
  sub-superscripts, left-side pre-sub/superscripts with empty script slots,
  radicals with optional degrees and canonical hidden empty degree slots,
  functions, supported n-ary operators, combining accents, overbars and
  underbars, group characters with explicit character
  position and baseline justification, phantoms with visible or hidden bases,
  independently zeroed width, ascent, or descent, and transparent spacing,
  border boxes with independently visible edges and four strike directions,
  semantic boxes with operator-emulation,
  no-break, differential-spacing, manual-break, and alignment properties,
  bounded rectangular matrices with base and column alignment, ordered
  row-spacing and column-gap rules, and minimum column widths, equation arrays
  with 1–64 rows, vertical base alignment, maximum/object distribution,
  row-spacing rules, and `&` alignment/spacer markers, lower and upper limit
  objects, and custom delimiters in strict or transitional UTF-8/UTF-16
  packages. Fraction properties enforce `fPr -> num -> den` ordering and
  canonicalize an absent or attribute-free `type` to the `bar` default.
  N-ary operators enforce optional `naryPr` before required `sub`, `sup`, and
  `e` slots. An absent `chr` defaults to U+222B, an attribute-free `limLoc`
  defaults to `undOvr`, disabled growth normalizes away, enabled growth
  round-trips, and hidden limits use canonical empty script slots.
  Attribute-free operator characters and contradictory hidden nonempty limits
  fail closed. Delimiters
  enforce optional `dPr` before 1–32 `e` arguments and the ordered
  `begChr -> sepChr -> endChr -> grow -> shp -> ctrlPr` property grammar.
  Omitted characters normalize to `(`, U+2502, and `)`, while attribute-free
  character properties remain explicitly empty. Empty argument slots and the
  growing, centered defaults canonicalize away; non-growing and shape-matched
  delimiters round-trip. Functions enforce optional `funcPr` before required
  `fName` and `e` slots and preserve empty name or argument slots. Every
  supported `CT_OMathArg` slot now preserves an empty argument and enforces
  `argPr -> expressions -> ctrlPr` ordering. Its optional trailing `ctrlPr`
  retains one bounded direct or revision-wrapped Word control. Absent or empty argument/control
  properties and absent, empty, or zero `argSz` values normalize to the
  default. Relative sizes from -2 through 2 round-trip in every argument slot.
  The 13 Word-effective parent/child pairs project through inverse-sign relative
  MathML `scriptlevel`; valid sizes elsewhere remain native-only. Out-of-range
  or malformed sizes, duplicate or misplaced properties, malformed control
  revisions, and semantic properties fail closed. Depth, node, text, model-size,
  matrix-dimension, cell-count,
  equation-array row/alignment-marker, and equation-count budgets are enforced.
  Invalid or non-combining accent characters, malformed math-run or function
  structures, invalid or contradictory fraction, radical, n-ary, delimiter, bar,
  group-character, phantom, border-box, box, or equation-array properties,
  malformed
  lower/upper limit structures,
  malformed, duplicated, reordered, or out-of-range matrix spacing/gap
  properties, ragged or over-limit matrices, over-limit
  equation arrays, malformed script-property, pre-script, or math-paragraph
  structures, malformed placement, namespace spoofing, nested math, and
  relationship-bound properties
  fail closed to bounded text with explicit diagnostics.
- Added source-backed DOCX package preservation. Browser import now registers
  the original package, and export retains safe source-only OPC parts,
  content-type declarations, and relationships while generated core parts stay
  authoritative. Invalidated signatures, VBA, ActiveX, and custom-ribbon parts
  are deliberately omitted, and a missing registered source fails closed
  instead of silently producing a lossy export. Imported source metadata also
  carries a SHA-256 fingerprint so a different re-registered DOCX is rejected.
- Preserved passive OOXML extensions inside regenerated `word/settings.xml`.
  Ignorable extension attributes and elements plus structurally valid,
  non-conflicting `mc:AlternateContent` blocks now survive strict or
  transitional UTF-8/UTF-16 source packages. Generated Word settings remain
  authoritative; malformed, relationship-bound, protection, template,
  mail-merge, field-update, and duplicate-setting markup is not restored.
- Preserved passive ignorable extension trees on regenerated DOCX styles and
  numbering definitions. Styles match by type and style ID, while imported
  abstract-numbering, concrete-numbering, and level metadata follows rewritten
  numbering IDs. Generated Word semantics remain authoritative; source-only,
  duplicate, relationship-bound, malformed, and ambiguous one-to-many
  extension mappings are dropped instead of attaching to the wrong identity.
- Preserved passive non-OOXML vendor extensions on uniquely matched picture
  drawings in regenerated document, header, and footer parts. Drawing identity
  uses the normalized anchor plus drawing-property ID across strict/transitional
  UTF-8/UTF-16 sources. Header and footer image identities now survive editable
  page-chrome HTML. Relationship-bound, source-only, duplicate, Microsoft/OOXML
  semantic, and ambiguous drawing branches remain disconnected.
- Preserved passive non-OOXML vendor extensions on stable paragraphs and their
  paragraph properties in regenerated document, header, and footer parts.
  Native `w14:paraId` plus `w14:textId` identities survive sanitized body and
  page-chrome HTML; text edits rotate the version ID, while formatting-only
  edits and moves retain it and copies or splits receive independent paragraph
  IDs. Changed text versions, duplicate identities, relationship-bound content,
  and Microsoft/OOXML semantic branches fail closed.
- Preserved passive non-OOXML vendor extensions on stable `w:tbl`, `w:tr`, and
  `w:tc` scopes plus their property nodes. Native row `w14:paraId` and
  `w14:textId` identities now survive body and page-chrome HTML; table and cell
  identity is conservatively derived from directly owned row and paragraph
  IDs. Row text or structure edits rotate the row version, copied identities
  are repaired, nested rows and cells no longer leak into outer-table export,
  and duplicate, cross-kind, relationship-bound, or semantic branches fail
  closed while generated table geometry and formatting remain authoritative.
- Preserved stable native footnote, endnote, comment, and reply identities
  across reorderings while assigning fresh IDs to copies and collisions.
  Resolved comments now emit valid `commentsExtended.xml` even without replies;
  passive extensions on uniquely matched note, comment, and `commentEx` roots
  survive regeneration, and valid `commentsIds` durable IDs are rebound to the
  final comment paragraph IDs. Deleted records, duplicate or namespace-spoofed
  identities, relationship-bound branches, and unsupported modern
  reaction/people sidecars fail closed instead of reviving stale metadata.
- Added native editable DrawingML pictures inside footnotes and endnotes.
  Public import and artifact export retain picture identity, layout, wrapping,
  crop, and layer metadata across the body and both note parts. Export repairs
  missing note-part image relationships from the OOXML writer, assigns
  collision-free relationship IDs, and validates every embedded media target.
  Passive non-OOXML extensions follow only uniquely matched note drawings
  across strict/transitional UTF-8/UTF-16 sources. Changed, duplicate,
  namespace-spoofed, relationship-bound, or semantic branches fail closed;
  generated geometry and media remain authoritative, while legacy VML,
  shapes, SmartArt, and drawing-bearing content-control wrappers normalize.
- Preserved text-stable direct runs inside uniquely matched footnotes,
  endnotes, comments, and replies. Passive extensions on paragraph, run, and
  run-property scopes now follow exact text and structural ancestry; safe
  unmodeled note properties survive, and unchanged plain-text comments regain
  relationship-free source run segmentation and formatting. Supported
  regenerated semantics remain authoritative, while edits, duplicate
  paragraphs or properties, wrapped or mixed semantic content, relationship
  references, and ambiguous mappings fail closed.
- Preserved supported hyperlink wrappers and their stable runs inside
  text-stable footnotes, endnotes, comments, and replies. Generated note links
  remain authoritative, while unchanged plain-text comments recover safe
  HTTP(S), `mailto`, or internal-anchor links, eligible tooltips, passive
  wrapper metadata, and relationship-free formatting. External relationships
  are validated against the owning part, deduplicated or assigned a
  collision-free ID, and rewritten in final XML. Text edits, missing or
  duplicate relationships, wrong types or target modes, relative or unsafe
  targets, combined external-plus-anchor destinations, namespace spoofing,
  unsupported wrappers, and ambiguous spans fail closed.
- Preserved text-stable static rich-text and plain-text content controls inside
  footnotes, endnotes, comments, and replies. Eligible inline controls and
  contiguous block controls recover their wrappers, aliases, tags, locks,
  signed native IDs, Word 2013 appearance and color, end-character formatting,
  passive extensions, and stable runs. Footnote and endnote tables now export
  as native editable OOXML blocks instead of flattened row text; structurally
  stable and nested tables can participate in rich-text block controls while
  generated geometry remains authoritative. ID collisions are rewritten
  without disturbing unconflicted source IDs. Text or table-structure edits,
  duplicate or ambiguous mappings, active bindings or placeholder state, form
  or nested controls, relationship-bound content, hyperlinks, math, and
  drawings fail closed before any wrapper is emitted.
- Preserved source DOCX font-table metadata and eligible internal obfuscated
  font payloads through relationship-ID collisions. Font references are
  rewritten to their final package IDs, while external references, wrong
  relationship or content types, duplicate identities, and payload-path
  collisions fail closed. Strict and transitional UTF-8/UTF-16 package XML is
  decoded consistently. Embedded fonts remain available to native DOCX
  consumers; browser editing, preview, and PDF export still use registered A3S
  fonts or substitution.
- Defined authoritative controlled-update behavior for reviewed Word ranges.
  Comment and tracked-change mutations now produce typed React, Vue, and Web
  Component conflict events plus an accessible warning, while harmless range
  movement and document switches remain quiet and orphaned comment records are
  retained for host recovery.
- Added editable Word tight and through image wrapping, including wrap-side
  controls, browser `shape-outside` contour presentation, and exact DrawingML
  `wrapPolygon` import/export for supported floating pictures without leaking
  internal export markers.
- Added editable four-edge Word image cropping with percentage validation,
  matching edit/preview presentation, and exact DrawingML `a:srcRect`
  preservation for inline and floating pictures. Export patches the owning
  picture deterministically without leaking internal markers into the DOCX.
- Preserved and authored precise Word floating-image anchors with signed
  horizontal and vertical offsets relative to the column, paragraph, margin,
  or page. Picture Properties validates the complete placement atomically,
  edit/preview apply matching offsets, and DOCX `positionH`/`positionV`
  round-trip without converting aligned anchors into offset anchors.
- Continued row-spanning Word table cells across every covered physical row
  during pagination. Combined `rowspan`/`colspan` cells now receive contiguous
  selection ranges and in-cell page-break widgets, while DOCX `vMerge` and
  `gridSpan` round-trip together.
- Completed editable nested Word tables across insertion, targeted inner-table
  sizing, DOCX import/export, and pagination. Outer rows can now split at
  nested-row boundaries instead of forcing a tall inner table to overflow as
  one atomic block.
- Added percentage-based table-column authoring in the Layout ribbon and Table
  Properties dialog. Percentage preferences survive merged cells and DOCX
  `tcW` round-trips while pixel `tblGrid` widths remain browser fallbacks.
- Preserved semantic Word theme color references for run text, run shading,
  table-cell fills, and independent cell borders, including tint and shade.
  Untouched formatting now writes the original theme attributes with a correct
  RGB fallback, while explicit color edits discard stale theme semantics.
- Added bounded PDFium-native text runs with stable indices, character and
  UTF-8/UTF-16 ranges, exact PDF-space bounds, deterministic validation, and an
  independent hard run limit to the native PDF text-layer receipt. PDFium 7881
  segment look-ahead indices are normalized as exclusive ends so valid final
  text runs cannot be rejected as out of range.
- Recorded the completed native exact-unit source-layout contracts: bounded,
  content-addressed PPTX slide rasters and PDFium-backed PDF page inventory,
  geometry, rendering, typed failures, and deterministic receipts now satisfy
  the Office-side requirements tracked in #1 and #4.
- Preserved imported Word numbering identities, abstract-numbering identities,
  and levels in the controlled document model. Separated list runs that belong
  to one native Word list now reuse one DOCX numbering instance on export
  instead of silently restarting under unrelated generated identities.
- Preserved native multilevel `numFmt` and compound `lvlText` patterns across
  controlled edits and DOCX export, including non-Latin numbering families.
  Continue Numbering now adopts the preceding native identity, while an
  explicit style change clears stale imported formatting metadata.
- Preserved native numbering suffix, level alignment, physical and logical
  indentation, hanging or first-line offsets, and `lvlRestart` rules. RTL list
  definitions keep `start`/`end` semantics instead of being flattened to
  `left`/`right` during browser editing and DOCX export.
- Added conditional table-style support for paragraph contextual spacing and
  outline levels across style precedence, controlled editor attributes, format
  copy and clearing, and DOCX export.
- Added a Spreadsheet command catalog and adopted the shared Office-oriented
  quick-access, adaptive, and collapsible ribbon. Conditional Formatting now
  lives under Home and Styles, Data exposes executable ascending and descending
  sort commands, and workbook recalculation is visible and executable with F9.
- Added focused component and controller coverage, desktop and compact
  Playwright regression, and a schema-validated deterministic A3S Test workflow
  for the aligned Spreadsheet ribbon.
- Added executable Paste, Cut, and Copy commands to the Spreadsheet Home
  clipboard group. Ribbon clicks and Traditional Office `Cmd/Ctrl+V`,
  `Cmd/Ctrl+X`, and
  `Cmd/Ctrl+C` shortcuts now share one typed command port, permission-resilient
  browser/local clipboard fallback, and grid-focus restoration.
- Added an Office-style Spreadsheet Format Painter to the Home clipboard group.
  Single-click one-shot and double-click locked sessions copy native cell-style
  patterns across ranges and sheets without changing values, formulas,
  comments, links, or merges; another click or Escape exits cleanly.
- Added bounded format capture and target guards, duplicate-target suppression,
  one controlled workbook batch per application, accessible pressed/live
  state, and desktop plus compact Web regression coverage.
- Added Office-style Spreadsheet AutoFilter under Data and Sort and Filter.
  `Cmd/Ctrl+Shift+L` toggles filtering, `Alt+ArrowDown` opens the selected
  header menu, and arrows, Space, Enter, and Escape operate it without leaving
  the grid.
- Added finite current-region discovery for single-cell selections, exact
  explicit-range filtering, safe empty/merge/pivot rejection, controlled
  selection and hidden-row preservation, accessible vendor filter controls,
  XLSX round-trip coverage, and desktop plus compact Web regression.
- Added Office-style Spreadsheet Freeze Panes under View and Window. The current
  cell freezes the rows above and columns to its left, with separate top-row,
  first-column, and unfreeze commands behind one controlled workbook update.
- Added Arrow, Home, End, Enter, and Escape menu operation, pressed and live
  state, selection and grid-focus restoration, XLSX round-trip coverage, and
  desktop plus compact Web regression. Delayed grid-focus recovery now yields
  to deliberate pointer and Tab navigation so repeated ribbon actions remain
  usable immediately after a controlled workbook remount.
- Added an Office-familiar Rows and Columns menu to Spreadsheet Home and Cells.
  The existing typed workbook commands now expose row insertion above or below,
  column insertion left or right, and selected-row or selected-column deletion
  without duplicating the structure-editing model.
- Added independent command availability, Arrow/Home/End/Enter/Escape menu
  behavior, exact grid-focus restoration, desktop and compact Web workflow
  regression, and schema-validated deterministic A3S Test coverage.
- Added an Office-familiar Merge and Center split control to Spreadsheet Home and
  Alignment. Its menu executes Merge and Center, Merge Cells, Merge Across,
  Unmerge Cells, and Unmerge and Fill, while `Ctrl+M` shares the primary path.
- Kept every merge intent within one controlled Fortune workbook batch, used
  the native merge model for availability and unmerge ranges, restored grid or
  invoker focus exactly, and added focused, XLSX round-trip, desktop, and
  compact Web regression coverage.
- Added the Traditional Office Clear menu to Spreadsheet Home and Editing with
  Clear All, Clear Formats, Clear Contents, Clear Comments, and Clear
  Hyperlinks. Delete
  and Backspace now share the typed Clear Contents path.
- Preserved content, formats, comments, hyperlinks, and merge geometry according
  to each Clear mode, including bounded range subtraction for borders,
  conditional formats, and alternating formats; each intent stays within one
  controlled workbook batch and restores grid focus.

## 0.3.0 - 2026-08-07

- Added a Writer command catalog for stable ribbon grouping and Office-compatible
  shortcut metadata, moved undo and redo into a compact quick-access toolbar,
  added persistent plus temporary ribbon collapse behavior, and made lower
  priority groups compact before the ribbon falls back to horizontal paging.
- Aligned Writer superscript and subscript with the Traditional Office
  `Ctrl+Shift+=` and `Ctrl+=` shortcuts and added deterministic desktop browser
  coverage for the
  expanded, collapsed, and temporary ribbon states.
- Made the displayed Traditional Office Writer shortcuts executable inside the
  document for font sizing, paragraph alignment and line spacing, heading
  styles, spelling,
  field refresh, comments, and track changes without capturing host inputs.
- Added a permission-free Writer formatting clipboard with Traditional Office
  `Ctrl+Shift+C` / `Ctrl+Shift+V`, a one-shot format painter, semantic-mark
  preservation, and single-transaction formatting paste.
- Extended Traditional Office alignment and format-copy shortcuts into page
  headers and footers, corrected their superscript and subscript shortcut descriptions,
  and added schema-safe body-format projection for page-chrome editors.
- Reordered the Writer Insert ribbon into Office-familiar Pages, Table,
  Illustrations, Links, Header and Footer, and Text groups, with page-number
  visibility beside header and footer commands.
- Added direct Writer Page Layout presets for margins, orientation, paper size,
  and one-to-three-column layouts, with custom margins and advanced columns
  routed to the matching Page Setup tab. Deterministic browser coverage proves
  live landscape and two-column rendering, Escape close, accessibility, and
  empty console and page-error diagnostics.
- Aligned Writer References, Review, and View grouping with Traditional Office
  terminology and order, added direct previous/next plus accept/reject revision
  commands, and
  marked Picture, Table, and Header and Footer tabs as contextual tools.
- Replaced arbitrary Writer zoom presets with Office-style 100%, One Page, and Page
  Width commands. Fit zoom is calculated from the live page and editor viewport;
  deterministic browser coverage resolves two tracked changes, verifies both
  fit modes, and captures accessible, error-free evidence.
- Made Writer's status-bar word count actionable with live page, word,
  character, and paragraph details plus the Traditional Office `Ctrl+Shift+G`
  shortcut. The
  labelled view-and-zoom toolbar supports arrow-key traversal, while compact
  Web layouts retain page and zoom controls before lower-priority status items.

## 0.2.2 - 2026-08-07

- Selected the Word `ascii`, `hAnsi`, `eastAsia`, or complex-script font slot
  from each run's actual text while preserving `bCs`, `iCs`, `szCs`, `cs`,
  `rtl`, and font-hint behavior for multilingual DOCX content.
- Added a deterministic 30-row Latin, Chinese, Arabic, Hebrew, and mixed-format
  fixture with A3S Test coverage and a real Traditional Office Writer PDF layout
  gate.
- Added a calibrated Chromium native-PDF fallback for Traditional Office
  reference captures
  when the embedded PDF renderer cannot initialize the exported document.

## 0.2.1 - 2026-08-06

- Matched Traditional Office Writer automatic line layout across common Latin
  and Chinese system fonts with measured per-font advances while retaining the original
  OOXML line-spacing multiple as the DOCX round-trip authority.
- Preserved section-level Word document-grid type and line pitch plus run-level
  `snapToGrid` overrides across DOCX import and export, and stopped exporting a
  generated document grid when the source document does not define one.
- Added deterministic 30-row common-font, 36-row CJK-font, and 18-row document-
  grid fixtures with A3S Test browser coverage and real Traditional Office PDF
  layout gates.

## 0.2.0 - 2026-08-06

- Added a real Traditional Office Writer page-layout gate that exports a
  deterministic A4 DOCX through Traditional Office, captures normalized A3S and
  Traditional Office pages, and rejects page-size,
  semantic-landmark, browser-error, or bounded pixel regressions.
- Matched Traditional Office automatic Word line spacing without changing the
  original OOXML multiple used for DOCX export, removed editor-only spacing
  around imported
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
- Expanded shared Office select menus to 44 px option rows on phones, with a
  taller viewport-bounded scroll region and preserved End-key selection,
  editor-focus recovery, and Escape-to-trigger restoration.
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

- Extracted the complete document, spreadsheet, presentation, and PDF editor
  engine into an independent package.
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
