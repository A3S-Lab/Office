# Editor Quality Roadmap

This roadmap turns A3S Office from a broad editor demo into a dependable
component product. Delivery is depth-first: finish the workflows users repeat
every day, prove native-file compatibility, and only then expand the command
surface.

The technical ownership model, Worker/WASM boundaries, and long-term engine
stages remain defined in
[Browser editor architecture](browser-editor-architecture.md). This document
owns product priority, quality gates, and the order in which capabilities are
made release-ready.

## Product Principles

1. Depth before breadth. A complete editing workflow is more valuable than
   another toolbar control without reliable import, undo, save, and export.
2. Word first. Document editing is the first quality track because it exercises
   layout, styles, tables, media, review, and long-document behavior together.
3. Native files are the acceptance test. A feature is not complete until its
   supported semantics survive import, edit, export, and reopen in Microsoft
   Office and WPS.
4. Controlled state stays predictable. One user intent emits one host update
   and one undo record. External content replacement must not corrupt
   selection, history, or editor state.
5. Public APIs describe stable product contracts. Internal controllers are not
   exported merely to make an integration possible.
6. Documentation ships with the capability. Every public Prop, event, content
   field, and Extension boundary must be documented and tested.

## Release Gates Shared by Every Editor

These are foundations, not a separate feature-expansion phase:

- No browser-native prompt, confirm, context menu, or select control in a
  primary workflow.
- Shared dialog, popover, menu, ribbon, status, save, loading, empty, and error
  states use the same placement and interaction rules.
- Keyboard-only users can reach and leave every primary surface. Focus returns
  to the invoking control when a temporary surface closes.
- React, Vue, and Web Component adapters expose the same supported behavior,
  with framework-specific names documented where they differ.
- Controlled updates, undo, redo, import, export, and read-only preview have
  focused tests.
- Desktop and compact visual checks cover normal, empty, loading, error, menu,
  dialog, and long-content states.
- Performance budgets are fixture-based. A release cannot replace measured
  evidence with a qualitative claim.

Current shared-shell evidence includes one modal focus boundary for the compact
Office sidebar, Presentation slide drawer, AI assistant, and editor dialogs.
Each surface receives focus on entry, wraps forward and reverse Tab navigation,
isolates obscured content with `inert`, lets only the top overlay handle Escape,
and restores the current semantic invoker even when React remounts that control.
Compact Word navigation, find, layout, citations, changes, and comments task
panes now use the same modal boundary while retaining persistent side-pane
behavior on desktop.
Phone, 768 px, and desktop browser checks keep overlay and persistent-pane
behavior distinct. Shared Office color palettes now replace their dense
desktop grid with eight larger columns on phones; spatial keyboard movement
follows the rendered column count across theme and standard-color sections.

## Priority 1: Word-Compatible Document Editing

Document work is delivered as complete vertical slices in the order below.
New ribbon commands wait unless they are required by the current slice.

### 1. DOCX Round-Trip Foundation

- Move package relationships, styles, numbering, themes, comments, headers,
  footers, notes, fields, and media toward loss-preserving browser state.
- Preserve unsupported package parts and attributes instead of regenerating
  them from visible HTML.
- Build small, reviewable fixtures for every supported semantic and include
  documents produced by both Microsoft Word and WPS.
- Report unsupported or normalized content explicitly before export.

Exit evidence:

- Import-edit-export-reopen passes in Word and WPS for the declared fixture
  matrix.
- Unedited supported files retain package relationships and unsupported parts.
- A failed conversion does not overwrite the source file or emit a
  fidelity-success claim.

### 2. Deterministic Pagination and Text Layout

- Resolve style inheritance, numbering, paragraph spacing, indents, tabs,
  bidirectional text, columns, section breaks, widow/orphan rules, and exact
  registered fonts before line layout.
- Reflow from the earliest affected block and reuse stable earlier pages.
- Keep page breaks as mapped editor decorations so layout never becomes
  editable document content.
- Define font substitution and missing-glyph behavior that matches export.

Current product evidence includes a responsive Page Setup task pane with
keyboard-operated Page, Columns and Sections, and Header and Footer tabs. The
default view contains only paper and margin controls; page-chrome editors mount
on demand so compact and desktop workflows do not begin with one mixed,
unbounded settings stack. Deterministic text collection now follows CSS font
weight matching while retaining exact family and normal/italic style
boundaries, keeping intermediate and synthesized bold weights on the WASM
layout path when a compatible registered face exists. Editing, read-only
preview, and PDF composition now share the same base typography and structural
content styles. Header and footer content occupies the physical page margins
without shifting body flow, while empty headers remain absent and PDF output
does not synthesize a filename header. Compact previews keep the same physical
page geometry and use bounded scrolling instead of changing margins and text
flow to fit the viewport.

Editing and read-only preview now retain one canonical TipTap surface and the
same Worker/WASM pagination result, so switching mode preserves shaped runs,
automatic page breaks, table fragments, and the computed page count. Page
chrome is treated as an overlay inside the physical margins rather than an
additional body-height deduction. Browser PDF export finds this mounted surface
by stable artifact ID, clones it without editing state, and crops physical pages
from bounded batches. The three browser rendering paths therefore share one
page-layout result; searchable text and vector PDF output remain separate
fidelity work.

Exit evidence:

- Deterministic page goldens cover Latin, CJK, Arabic, Hebrew, lists, sections,
  headers, footers, and mixed formatting.
- Typing and selection do not create main-thread tasks above the documented
  budget on long-document fixtures.
- Pagination does not create extra undo entries or move a stable selection.

### 3. Styles, Lists, and Tables

- Make named styles, direct formatting, numbering, restart/continue behavior,
  nested lists, and theme values round-trip together.
- Complete table insertion, selection, row/column/cell operations, merged
  cells, repeating headers, nested tables, row splitting, borders, shading,
  alignment, and sizing.
- Ensure contextual commands operate on typed selections, including
  multi-cell and keyboard selection.

Current implementation evidence includes the responsive paragraph-style
gallery and a typed list slice with disc/circle/square bullets, five common
ordered formats, restart/continue/start commands, keyboard-operated galleries,
single-transaction undo behavior, and DOCX round trips for the declared list
styles. Tables now expose separate Design and Layout contextual tabs, five
keyboard-operated style presets, multi-cell fill and uniform border commands,
horizontal and vertical alignment, exact column-width and row-height fields,
equal row/column distribution, content/window autofit, row/column operations,
merge/split, repeat-header, and non-splitting row controls. Common cell
shading, vertical alignment, uniform borders, fixed grid widths, layout mode,
and explicit row heights round-trip through editable HTML and DOCX. Native
Word list identities, arbitrary multilevel numbering, per-edge and theme table
borders, percentage-width column authoring, nested tables, and all multi-page
row cases remain part of this priority rather than being treated as complete.

Exit evidence:

- Formatting and table operations have command, interaction, undo, DOCX, and
  visual tests.
- Large tables paginate without duplicate content, selection loss, or
  quadratic reflow.

### 4. Media, Fields, and Page Objects

- Complete inline and floating images, wrapping, anchors, captions, links,
  bookmarks, cross-references, page numbers, dates, fields, footnotes, and
  endnotes.
- Keep object identity stable across move, copy, delete, undo, and DOCX
  relationship updates.
- Make image and object dialogs validate size, placement, alternative text, and
  destructive replacement before commit.

Exit evidence:

- Relationship-bearing objects reopen with working targets in Word and WPS.
- Object layout has deterministic fixtures at page and section boundaries.

### 5. Review and Long-Document Quality

- Finish anchored comments, replies, resolve/reopen, tracked insert/delete,
  accept/reject, citations, navigation, find/replace, outline, and references.
- Define conflict behavior when a controlled host update changes a reviewed
  range.
- Virtualize or incrementally derive outlines, comments, revisions, and page
  chrome for large files.

Current implementation evidence includes a persistent Word-style navigation
pane with a typed heading hierarchy, active-heading tracking, collapsible
branches, keyboard traversal, and responsive left-side placement. Its search
returns full-text results with section context, bounded excerpts, match
highlights, and safe selection-based jumps that do not create history entries.
Compact result selection closes the modal pane before restoring the exact body
selection and focus. The same pane now offers a page view with measured textual
previews, physical-page and restarted-page-number labels, active-page tracking,
arrow/Home/End keyboard traversal, and selection-safe jumps to each page. A
page preview is derived from the measured text ranges assigned to that page, so
pages split from one long paragraph do not repeat the paragraph's opening text.
The dedicated Find/Replace task pane now has deterministic phone-width coverage
for query entry, match navigation, single replacement, disabled-action focus
recovery, content synchronization, and modal close-to-invoker focus restoration.
New comment
drafts stay inside the visible review rail even for document-wide selections,
and discard confirmation is limited to drafts or replies that contain written
content. Cancelling a task-pane switch now returns keyboard focus to the exact
unfinished comment, reply, or citation field on desktop and compact layouts,
so protected content can be edited immediately. The same editing context is
retained when users cancel comment deletion, citation deletion, or an internal
citation switch. Raster-quality page thumbnails and bounded page,
outline/search-result virtualization for 100-page fixtures remain part of this
priority.

Exit evidence:

- Review operations retain authorship, anchors, replies, and state through
  DOCX round trips within the declared compatibility boundary.
- Representative 100-page and review-heavy fixtures meet interaction, memory,
  reflow, and save budgets.

## Priority 2: Markdown

After the Document gates above are stable:

- Prove lossless switching between source, split, visual, and preview modes for
  the supported GFM model.
- Complete tables, task lists, links, images, fenced code, escaping, and
  clipboard behavior.
- Publish Markdown Extension parse/serialize requirements and fixtures.
- Keep source typing coalesced and visual updates responsive on large files.
- Keep source undo/redo independent from browser-native textarea history,
  restore selections, rebase on host replacements, and avoid duplicate visual
  history records during synchronization.
- Keep the default split view directly resizable without breaking synchronized
  scrolling or the stacked compact layout.

Exit evidence: canonical Markdown fixtures round-trip without semantic drift,
custom structural Extensions have serialization tests, and source/visual undo
remains one coherent history.

## Priority 3: Spreadsheet

- Replace remaining dense and main-thread workbook work with an A3S-owned
  sparse model and virtualized viewport.
- Expand Worker/WASM formula parity, dependency tracking, number formats,
  sorting, filtering, validation, charts, pivots, and print layout.
- Keep lightweight selection statistics responsive for dense and sparse sheets.
- Keep common number and percent presets plus decimal-place commands available
  from the primary ribbon while the native format engine expands.
- Keep font family, vertical alignment, and text wrapping available from the
  primary ribbon and backed by the native workbook cell-style model.
- Keep worksheet lifecycle actions in the shared Office interaction system:
  destructive deletion uses a safe default, invalid names remain editable with
  local accessible feedback, and temporary menus and dialogs restore focus.
- Keep phone worksheet renaming focused on the active task: unrelated footer
  tools yield the available width, invalid feedback expands into a readable
  second row, and successful or cancelled edits restore the normal workbook
  status controls without covering the grid.
- Keep Cmd/Ctrl+F inside the editor through the compact Spreadsheet Find bar,
  including displayed, raw, formula, and sparse-cell matching, deterministic
  navigation, repeated-shortcut refocus, and grid-focus restoration on close.
- Preserve direct type-to-edit after focus recovery and keep multi-cell paste,
  selection, copy, and undo coherent when React replays vendor state updates.
- Preserve XLSX relationships, names, styles, formulas, comments, charts,
  images, and unsupported package state.

Exit evidence: render cost depends on the viewport, recalculation depends on
the dirty dependency graph, and the compatibility fixture matrix passes in
Excel and WPS.

## Priority 4: Presentation

- Complete masters, layouts, themes, placeholders, groups, tables, charts,
  links, notes, comments, transitions, and presenter workflows.
- Keep exact row and column table insertion behind the shared, keyboard-
  accessible Office table picker and one controlled history record.
- Keep scene object identity stable through selection, transforms, clipboard,
  history, import, and export.
- Keep native slide and object context actions available independently of
  optional AI integrations.
- Keep F5 and Shift+F5 as direct slideshow paths with an in-page fallback when
  the browser denies fullscreen.
- Move bounded geometry, text fitting, connector routing, snapping, and
  thumbnail work behind Worker/WASM contracts.

Exit evidence: complex slides stay interactive, supported objects preserve
their semantics through PowerPoint and WPS round trips, and one gesture creates
one controlled update.

## Priority 5: PDF

- Keep color, opacity, and compatible stroke-width controls routed through
  typed annotation capabilities for both tool defaults and selected objects.
- Add form authoring, signatures, page organization, and explicit redaction
  review while preserving the implemented search, history, and save paths.
- Keep navigation, zoom, and history reachable from the responsive overflow
  menu whenever the primary toolbar hides those controls.
- Retain the implemented active-page thumbnail synchronization, bounded
  long-document thumbnail window, and focus-contained phone page drawer.
- Keep PDFium lifecycle, page bitmap cache, and source Blob ownership bounded.
- Verify saved annotations and forms by reopening the emitted Blob.

Exit evidence: large files render a bounded page window, edits survive reopen,
and destructive operations always show a review step before save.

## Public Extension Plan

The public Extension surface grows only when its context can remain stable
across minor releases.

### Available now

- `DocumentEditor.extensions`: additional TipTap Extensions.
- `MarkdownEditor.extensions`: additional TipTap Extensions.
- `DocumentEditor.getSelectionMenuItems`: a host-owned selected-text menu
  factory with full document snapshots and conflict-aware text-edit commands.
- `MarkdownEditor.getSelectionMenuItems`: the same host-owned boundary across
  source and visual selections, with complete Markdown context and safe edits.
- Both editors append host Extensions after built-ins and throw on duplicate
  names.
- Vue accepts `:extensions`; custom elements accept the `.extensions`
  property.

Structural TipTap Extensions own a persistence obligation. Document Nodes and
Marks need DOCX import/export handling. Markdown Nodes and Marks need Markdown
parse/serialize rules. An Extension that only adds shortcuts, storage, or a
ProseMirror Plugin does not change the persisted content schema and is the
recommended first integration.

### Stable host extension points

- Document and Spreadsheet: `fileActions` and `onAgentRequest`.
- Presentation: `fileActions`, `onAgentRequest`, and `onStartSlideshow`.
- PDF: `loadSource` and `onSave`.

Spreadsheet, Presentation, and PDF command contexts remain internal until each
has a typed, versioned Extension context, deterministic command ordering,
capability discovery, cleanup semantics, and compatibility tests. Consumers
must not import `internal` modules.

## Playground Documentation Definition of Done

The Playground integration guide is the public API reference for a release.
For each editor it must include:

- Every public Prop with its exact type, required state, default value,
  framework binding, and behavioral description.
- React callbacks, Vue events, and Web Component events or properties.
- The canonical content model or, for PDF, the Blob load/save lifecycle.
- Supported Extensions, collision behavior, lifecycle guidance, persistence
  requirements, and a copyable highlighted example.
- Document and Markdown selection-menu snapshots, async target lifetime,
  stale-selection behavior, editing commands, and React/Vue/Web Component
  bindings.
- Explicit alternatives when an internal Extension API is not public.
- Links or examples for React, Vue, and Web Components.

The guide is tested as product code. Navigation, all five editor tabs, Props
tables, content contracts, Extension examples, code highlighting, horizontal
table scrolling, and compact layout must pass before release. A public API
change is incomplete until the corresponding guide and tests change in the
same pull request.

## Execution Order

1. Keep the shared quality gates green while finishing the current Word
   vertical slice.
2. Close that slice with unit, interaction, visual, performance, and
   Word/WPS round-trip evidence.
3. Update public types and Playground documentation only after the contract is
   stable.
4. Move to the next Word slice. Do not start a lower-priority editor milestone
   merely because its UI is easier to demonstrate.
5. Revisit the priority order only with measured user demand, compatibility
   risk, and engineering evidence.
