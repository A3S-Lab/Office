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
The Office shell keeps persistent desktop navigation and the temporary compact
drawer as separate states. Crossing into the compact breakpoint therefore
closes the drawer in the same render without obscuring the workspace or moving
its active focus, while returning to desktop restores the user's persistent
sidebar preference.
Compact Word navigation, find, layout, citations, changes, and comments task
panes now use the same modal boundary while retaining persistent side-pane
behavior on desktop.
Phone, 768 px, and desktop browser checks keep overlay and persistent-pane
behavior distinct. Shared Office color palettes now replace their dense
desktop grid with eight larger columns on phones; spatial keyboard movement
follows the rendered column count across theme and standard-color sections.
Shared editor context menus remain pointer- or selection-anchored on desktop,
but become safe-area-aware bottom action sheets below 640 px. Their actions are
at least 44 px tall, long menus scroll inside a bounded surface, Escape returns
focus to the editing target, and deterministic Spreadsheet evidence captures
the phone layout, accessibility tree, console, and page errors.

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
unbounded settings stack. Header and footer formatting actions are divided
into balanced text-formatting and alignment/insertion rows instead of leaving
orphaned controls on an accidental wrap line; a deterministic phone workflow
now covers all three tabs, a live column update, modal isolation, and invoker
focus restoration. Paragraph-spacing and paragraph-pagination popovers now
expand to a viewport-safe 320 px surface on phones. Their numeric fields,
steppers, checkboxes, and reset actions provide 44 px touch targets while
preserving direct document commits and exact trigger-focus restoration. A
deterministic phone workflow proves both semantic changes and captures
accessibility, console, and page-error evidence.

The Writer Insert ribbon now follows the WPS Pages, Table, Illustrations,
Links, Header and Footer, and Text sequence, including page-number visibility
with the other page-chrome commands. Page Layout exposes direct normal, narrow,
moderate, wide, orientation, paper-size, and one-to-three-column presets. Custom
margins and advanced columns open the matching controlled Page Setup tab. A
deterministic desktop workflow proves live landscape and two-column rendering,
advanced-tab routing, Escape close, accessibility output, and empty browser
diagnostics.

References, Review, and View now follow the WPS-oriented information sequence
without advertising an unimplemented table-of-contents command. Review adds
direct previous/next and accept/reject commands backed by live revision ranges;
resolving a change selects the next available revision and collapses to a safe
caret when none remain. View orders document modes before display controls and
replaces arbitrary presets with 100%, One Page, and Page Width. The fit commands
use live page and viewport measurements, while contextual Picture, Table, and
Header and Footer tabs are explicitly distinguished. A deterministic desktop
workflow records two changes, navigates and resolves them, exercises both fit
modes, and saves accessibility plus empty console and page-error evidence.

Writer's status bar now exposes its word count as a real command instead of a
passive label. Clicking it or pressing the WPS `Ctrl+Shift+G` shortcut opens
live page, word, character-with-and-without-spaces, and paragraph statistics,
then restores the invoking status control or document caret. The labelled view
and zoom toolbar supports arrow-key traversal. At compact Web widths, status
items are removed by priority while page position and zoom remain available; a
deterministic desktop-to-phone workflow captures both states with accessibility
and browser-diagnostic evidence.

Deterministic text collection
now follows CSS font weight matching while retaining exact family and
normal/italic style boundaries, keeping intermediate and synthesized bold
weights on the WASM layout path when a compatible registered face exists.
Editing, read-only preview, and PDF composition now share the same base
typography and structural content styles. Header and footer content occupies
the physical page margins without shifting body flow, while empty headers
remain absent and PDF output does not synthesize a filename header. Compact
previews keep the same physical page geometry and use bounded scrolling instead
of changing margins and text flow to fit the viewport.

Editing and read-only preview now retain one canonical TipTap surface and the
same Worker/WASM pagination result, so switching mode preserves shaped runs,
automatic page breaks, table fragments, and the computed page count. Page
chrome is treated as an overlay inside the physical margins rather than an
additional body-height deduction. Browser PDF export finds this mounted surface
by stable artifact ID, clones it without editing state, and crops physical pages
from bounded batches. The three browser rendering paths therefore share one
page-layout result; searchable text and vector PDF output remain separate
fidelity work.

A deterministic WPS Writer layout gate now covers the first declared native
page-parity slice. The fixture uses an A4 page, installed Arial runs, explicit
automatic paragraph spacing, a fixed centered table, direct fills and borders,
and physical cell margins. WPS exports the reference PDF through its desktop
automation API; the browser captures both page surfaces at 794 by 1123 CSS
pixels. The comparison requires one physical page, matching title/body/table
landmarks within one pixel, no more than a 2% thresholded pixel difference,
and a mean absolute channel error no greater than 1.0. The current measured
result is a one-pixel maximum landmark delta, a 1.4612% thresholded difference,
and 0.7308 mean absolute error. Original OOXML line multiples remain the DOCX
round-trip authority instead of being replaced by the browser-only WPS metric.

Four focused matrices now extend that gate beyond Arial and table geometry.
The common-font matrix contains 30 rows across Arial, Times New Roman, Calibri,
Segoe UI, and Microsoft YaHei at two sizes and two automatic spacing multiples.
The CJK matrix contains 36 Latin and Chinese rows across Microsoft YaHei,
SimSun, SimHei, FangSong, KaiTi, and DengXian. The explicit document-grid matrix
contains 18 rows across three fonts and two sizes. Browser layout uses the
measured WPS per-font advance without changing the source OOXML spacing
multiple. DOCX import and export also retain each section's `docGrid` type and
line pitch and each run's `snapToGrid` override. The deterministic browser
suite requires one page, exact semantic metric markers, and empty console and
page-error diagnostics; the WPS comparison rejects a text-band top drift above
three pixels or a consecutive-band advance drift above four pixels. The current
maximum top/advance deltas are 3/1 pixels for common fonts, 2/2 pixels for CJK
fonts, and 1/1 pixels for the document grid.

The fourth matrix adds 30 Latin, CJK, Arabic, Hebrew, and mixed-formatting rows
at two sizes. DOCX import now chooses the Word `ascii`, `hAnsi`, `eastAsia`, or
`cs` font slot from the run text and honors `bCs`, `iCs`, `szCs`, `cs`, `rtl`,
and `rFonts` hints before calculating the WPS line-height factor. The same
fixture proves paragraph RTL direction, complex-script emphasis, mixed-run
bold and italic formatting, one-page pagination, and empty browser diagnostics.
Its real WPS 12.1 reference has 30 matching text bands with a three-pixel
maximum top delta and a three-pixel maximum consecutive-band advance delta.

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
styles. Writer now has one command catalog for stable WPS-oriented tab and
group placement plus shortcut metadata. Its displayed WPS font-size,
alignment, line-spacing, heading, spelling, field-refresh, comment, and
track-changes shortcuts execute inside the document while host text inputs and
modal surfaces retain their native keyboard behavior. WPS copy-format and
paste-format shortcuts share a permission-free local formatting clipboard with
the one-shot format painter; applying it keeps links and review marks intact
and commits direct character and paragraph formatting in one transaction.
Header and footer surfaces use the same WPS alignment and format-copy commands,
falling back to their compatible paragraph schema when body heading formats are
pasted. Undo and redo use the quick-access row, while double-clicking the active
tab persistently collapses the ribbon; tab activation temporarily overlays
commands and outside pointer input closes that overlay without moving document
layout. The toolbar's measured width compacts low-priority groups first,
normal-priority groups second, and retains high-priority groups before
horizontal paging. A deterministic desktop workflow checks comfortable,
adaptive, collapsed, and temporary states, accessibility output, and empty
console and page-error diagnostics. Compact ribbon overflow now reserves both
navigation edges so tool width does not change after paging. Phone list
galleries stay within the
viewport and provide 44 px clear, restart, start-value, stepper, and commit
controls. Closing a list command restores the TipTap selection as well as DOM
focus, so reopening the numbering gallery retains its active style and start
settings. A deterministic phone workflow covers two ribbon advances, square
bullets, upper-Roman numbering, a custom start value, editor-focus recovery,
and Escape-to-invoker restoration. Tables now expose separate Design and
Layout contextual tabs, five keyboard-operated style presets, multi-cell fill,
and a reusable border pen with all, outside, inside, side, inside-horizontal,
and inside-vertical targets. Horizontal and vertical alignment, exact column-
width and row-height fields, equal row/column distribution, content/window
autofit, row/column operations, merge/split, repeat-header, and non-splitting
row controls are also available. Common cell shading, vertical alignment,
explicit per-edge borders, fixed grid widths, layout mode, and explicit row
heights round-trip through editable HTML and DOCX. Table geometry now models
the layout algorithm, auto/percentage/pixel preferred width, alignment, indent,
table-level cell margins, and cell-level margin overrides independently.
`tblLayout`, `tblW`, `jc`, `tblInd`, `tblCellMar`, `tcMar`, `tblGrid`, and
`tcW` round-trip without treating grid data as the preferred table width. A
real fixture proves that a centered 62.5% table with a first-cell margin
override keeps matching edit and preview geometry. Table insertion or removal
does not reinterpret an imported autofit table as fixed merely because it has
grid widths. Percentage column preferences round-trip as `tcW` while stable
pixel `tblGrid` widths remain browser-layout fallbacks, including across merged
cells. The Layout ribbon reports the rendered column width for autofit tables
and labels noncanonical preferred widths as Current width instead of
misreporting them as fixed.
The Table Properties dialog now owns four coherent tabs: Table, Row, Column,
and Cell. It combines auto, percentage, or centimeter preferred table width,
left/center/right placement, and left indent with selected-row height and rule,
page splitting, repeated headings, percentage or centimeter current-column
width, selected-cell vertical
alignment, and four-edge margins. The complete validated draft commits in one
TipTap transaction and one undo record; Cancel and Escape preserve the last
committed state and restore the exact ribbon invoker. Values the user did not
edit retain their exact imported pixels and partial inheritance instead of
round-tripping through the two-decimal centimeter display. The deterministic
styled-table workflow checks truthful 62.5% centered import values, all four
tabs and touch-sized controls at 390 px, an atomic change to 72% left placement
with a 0.5 cm indent, a 1.2 cm exact row, disabled page splitting and repeated
headings, a 4 cm column, centered cell content, a 0.4 cm left cell margin,
cancellation safety, exact untouched 8 px and 16 px margin edges, and matching
preview geometry with clean browser diagnostics. Table-level outer and inside
DOCX borders resolve to cell edges on import without flattening mixed styles.
Direct table and cell theme colors now resolve through the package theme,
including `themeTint`, `themeShade`, `themeFillTint`, and `themeFillShade`.
The resolved colors remain identical in edit and preview. Untouched run text,
run shading, cell fills, and independent cell borders retain their semantic
theme references, tint, and shade on export with the resolved RGB as a valid
fallback; an explicit color edit exports direct RGB and clears stale semantics.
DOCX table styles now resolve the default or referenced table style through a
bounded, cycle-safe `basedOn` chain. `tblLook` flags or bitmasks, row and column
band sizes, grid spans, and row grid offsets select whole-table, banded row or
column, first/last row or column, and corner-cell conditions in Word precedence
order. Conditional fills, per-edge borders, fonts, colors, bold, italic,
underline, and strikethrough enter the same edit and preview model before
direct table, cell, paragraph, and run formatting. Conditional paragraph
alignment, direction, indents, spacing and line rules, pagination rules, and
tab stops follow that same precedence. Theme-derived style colors are
materialized as stable RGB on export. A real conditional-style DOCX fixture is
covered by deterministic A3S Test evidence for header, body-band, last-row, and
paragraph-layout presentation with a clean browser diagnostic run.
Splittable rows taller than a complete continuation-page body now fragment at
measured paragraph boundaries instead of overflowing as one atomic row. The
same layout path repeats leading heading rows on each continuation page. A real
DOCX fixture with one 120-paragraph body row deterministically produces three
physical pages; A3S Test verifies the page-two and page-three cell breaks and
repeated headings, keyboard access to the final marker and following paragraph,
and an error-free browser run.
Desktop insertion keeps an 8 × 10 keyboard matrix without duplicate cell
semantics, while phone insertion uses focused row and column controls with
44 px touch targets. A deterministic phone workflow now proves 3 × 3 insertion,
row extension, compact-ribbon deletion, and editor-focus recovery.
Percentage-width column authoring now retains pixel grid fallbacks, and nested
tables remain editable through targeted inner-table commands and DOCX
round-trips. Outer rows paginate at nested-row boundaries instead of treating a
tall inner table as one atomic block. Row-spanning cells participate in every
covered physical row with contiguous content ranges and in-cell page-break
widgets; combined row and column spans retain `vMerge` and `gridSpan` through
DOCX round-trips.

Imported Word list identities now survive the DOCX-to-controlled-HTML boundary
as explicit numbering, abstract-numbering, and level metadata. Export reuses a
single generated DOCX numbering instance across separated runs with the same
native identity, retaining continuation semantics instead of assigning each
run an unrelated list. This establishes the identity boundary required by the
remaining multilevel fidelity work.

Native `numFmt` values supported by the OOXML writer and compound `lvlText`
patterns now survive that same boundary, including non-Latin numbering
families and references to parent levels. Continue Numbering adopts the
preceding native identity and definition; an explicit gallery style change
clears stale imported format and text metadata before export. Native suffix,
level alignment, physical or logical indentation, hanging or first-line
offsets, and `lvlRestart` rules now follow the same round-trip path. Logical
`start` and `end` indentation remain distinct for RTL numbering instead of
being flattened to physical left and right values.

Conditional table-style paragraph properties now retain contextual spacing and
outline levels through the same default-style, paragraph-style, inherited table
style, conditional region, and direct-format precedence chain. Both properties
enter the controlled editor model, participate in format copy and explicit
format clearing, and return to DOCX paragraph properties. Focused package and
round-trip evidence covers direct overrides against `wholeTable` and
`firstRow` conditions.

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

Current implementation evidence includes one project-owned file
materialization boundary shared by Word body images, header/footer images,
Presentation images, and imported Office/PDF files. File inputs can reset
immediately for same-file reselection without invalidating an in-flight read.
Unit coverage verifies single-read ownership and the A3S Test picture fixture
verifies that a relative-path PNG reaches the Word canvas and opens its
contextual picture toolbar without a false read error. The phone picture
workflow now exercises one Picture Properties dialog for centimeter width and
height, per-image aspect-ratio locking, square, tight, through, top-and-bottom,
and no-wrap free-floating placement, wrap-side selection, alignment, text
distance where applicable, signed
horizontal and vertical anchor offsets with column, paragraph, margin, or page
reference frames, four-edge percentage cropping, and
alternative text. No-wrap images use a one-pixel pagination anchor instead of
reserving their rendered height, remain absolutely positioned in edit, preview,
and print surfaces, and use the existing behind-text drawing-layer option to
select whether text paints above or below the image. The complete draft
commits as one undo entry,
untouched imported dimensions keep their exact source values, the selected
image remains active, and apply, cancel, and Escape restore the exact Picture
Properties command so the next keyboard action cannot mutate document content.
Image nodes now carry a stable object identity together with their native DOCX
drawing-property, floating-anchor, and edit identifiers. Transaction mapping
keeps the pre-existing object stable when a copy is inserted on either side,
creates a complete independent identity for the copy, and retains identity
through move, cut-and-paste, delete, undo, and redo. DOCX regeneration assigns
unique `wp:docPr` identifiers, preserves `wp14:anchorId` plus untouched
`wp14:editId` values for inline and supported floating drawings, advances the
edit identifier when the image is changed or moved, and validates every
emitted image relationship without using the relationship ID as object
identity.
Word body bookmarks now use paired inline boundaries rather than a flattened
single marker, so a target can span paragraphs while remaining one atomic
editable range. New names follow Word's 40-character letter-first convention;
imported hidden names remain valid, malformed or duplicate identities are
normalized deterministically, and copying a complete range assigns a new
name, internal identity, and native numeric identifier while preserving the
original. Internal links export through `w:anchor` without an external
relationship, supported web and email destinations retain hyperlink
relationships, and deleting a bookmark gives every dangling internal link an
explicit state that undo repairs. DOCX inspection reports ambiguous boundary
pairs, normalized names, column-scoped bookmarks, missing targets, unsupported
destinations, and advanced hyperlink metadata separately. Body bookmarks are
also available in the cross-reference picker: their live display follows target
edits, deletion produces an explicit missing-reference state, undo repairs it,
and copy normalization retargets self-references to the copied identity. Native
DOCX import/export preserves ordinary bookmark `REF` fields with the supported
hyperlink switch and diagnoses advanced switches separately. Unit round trips
cover cross-paragraph native ranges, internal-versus-external relationships,
and bookmark fields; a deterministic phone workflow covers dialog focus,
reference insertion, deletion, the missing-reference transition, and exact
editor-focus recovery.
Footnote and endnote references now share one live identity graph with their
editable definitions. Each kind renumbers independently from reference order;
copy creates a new identity with cloned note content, deleting either complete
side removes its pair, and undo or redo restores the graph atomically. New
footnotes stay with their reference section while new endnotes use the final
section. Controlled HTML normalization and DOCX inspection detect repeated,
missing, unreferenced, or nested identities, and the high-priority note parser
preserves native references through export, import, and a second export. Unit
coverage exercises reverse insertion, copy, deletion from either side,
undo/redo, final-section endnotes, malformed controlled HTML, diagnostics, and
native footnote/endnote round trips; a deterministic phone workflow covers
live renumbering, paired deletion, restoration, accessibility, and browser
diagnostics.
Body `PAGE`, `NUMPAGES`, `SECTION`, `SECTIONPAGES`, `DATE`, and `TIME` fields
now resolve from the measured Worker/WASM page containing each atom, including
physical-page membership for continuous sections. Pagination refreshes numeric
results without adding history or advancing clock fields; scoped F9 refreshes
all field kinds as one undoable action. Transaction mapping preserves moved
field identities, gives copied atoms fresh redo-stable identities, and keeps
insertion plus its initial refresh in one history step. Native DOCX simple
fields round-trip twice, complete inline complex fields import as editable
atoms, and incomplete, nested, cross-paragraph, deleted, or instructionless
structures remain text with an explicit compatibility warning. Unit coverage
exercises live page contexts, continuous sections, copy and undo/redo identity,
clock-safe automatic refresh, structural diagnostics, and native DOCX output;
a deterministic phone workflow covers physical PAGE and NUMPAGES results, F9,
atomic undo/redo, accessibility, and browser diagnostics.
Caption order and cross-reference validity now share one live transaction
graph: deleting or reordering a caption renumbers surviving targets, updates
every linked field, and renders dangling references as an explicit missing
state. Caption numbers are included in the accessible caption name, and a
deterministic phone workflow proves dialog focus, insertion, deletion, and the
missing-reference transition.

Exit evidence:

- Relationship-bearing objects reopen with working targets in Word and WPS.
- Object layout has deterministic fixtures at page and section boundaries.

### 5. Review and Long-Document Quality

- Finish anchored comments, replies, resolve/reopen, tracked insert/delete,
  accept/reject, citations, navigation, find/replace, outline, and references.
- Define conflict behavior when a controlled host update changes a reviewed
  range.
- Keep outlines, search results, revision review, and comment review on bounded
  windows, and incrementally derive page chrome for large files.

Current implementation evidence includes a persistent Word-style navigation
pane with a typed heading hierarchy, active-heading tracking, collapsible
branches, keyboard traversal, and responsive left-side placement. Its search
returns full-text results with section context, bounded excerpts, match
highlights, and safe selection-based jumps that do not create history entries.
Compact result selection closes the modal pane before restoring the exact body
selection and focus. The same pane now offers a page view with real rasters
cropped from the live paginated surface, physical-page and restarted-page-number
labels, active-page tracking, arrow/Home/End keyboard traversal, and
selection-safe jumps to each page. The current and adjacent pages enter a
serialized capture queue first; the rest use bounded viewport admission,
debounced source-mutation refresh, and off-screen image release. Measured text
is retained only as a loading or failure fallback. Capture readiness does not
depend on unrelated document fonts or animation frames that background agent
tabs can suspend.
The dedicated Find/Replace task pane now has deterministic phone-width coverage
for query entry, match navigation, single replacement, disabled-action focus
recovery, content synchronization, and modal close-to-invoker focus restoration.
The compact revision pane now exposes recording controls inside its modal
boundary, reports whether new changes are actually being recorded, retains
focus when that state changes, and keeps the truthful active state after the
last pending revision is resolved. A deterministic phone workflow covers
enabling tracking, creating and accepting an insertion, empty-state recovery,
and exact invoker restoration. Revision collections above 48 items now reuse
the same bounded window as document navigation, mounting at most 32 contiguous
rows while physical spacers preserve the complete scroll range. Arrow,
PageUp/PageDown, Home, and End keep the whole review queue keyboard reachable;
individual accept/reject decisions retain the same action focus on the next
revision. A real DOCX fixture with 120 native OOXML insertions proves bounded
mounting, first/last access, the 120-to-119 decision transition, focus
continuity, spacer geometry, and an error-free browser run.
Comment collections above 48 items now mount at most 32 contiguous cards while
the active comment and comments with unsent replies remain pinned. Anchor
geometry is collected in one linear pass, connector rendering is limited to
mounted cards, and measured card heights are reused when a card leaves the
window. Arrow, PageUp/PageDown, Home, and End traverse the complete collection;
keyboard navigation keeps the document selection, physical page, review
window, and focus synchronized. Deleting a comment moves focus and selection
to its adjacent surviving comment. A real DOCX fixture with 120 native OOXML
comments proves anchor import, bounded mounting, first/last access, the
120-to-119 deletion transition, page synchronization, and an error-free browser
run. New comment drafts stay inside the visible review rail even for
document-wide selections, and discard confirmation is limited to drafts or
replies that contain written content. Cancelling a task-pane switch now returns
keyboard focus to the exact
unfinished comment, reply, or citation field on desktop and compact layouts,
so protected content can be edited immediately. The same editing context is
retained when users cancel comment deletion, citation deletion, or an internal
citation switch. Citation source validation is now attached to the exact
invalid tag or title field, moves keyboard focus to that field, and clears only
when its value changes. A deterministic phone workflow covers validation,
source persistence, citation insertion, and exact close-to-invoker focus
restoration. Raster-quality page thumbnails and bounded raster admission are
implemented. Page buttons now use a bounded window with physical scroll spacers
and sparse current/focus pins. A deterministic 120-page DOCX workflow proves
Home/End navigation, selection, focus retention, and bounded mounting. The same
window model now bounds heading and search-result rows above 48 items to at most
32 contiguous rows plus sparse active, selected, and keyboard-roving pins. The
real 120-page fixture contributes 120 headings and 120 matches; deterministic
browser evidence proves first/last keyboard access, exact final-result
selection, physical spacer geometry, and zero console or page errors. Comments
now have the same bounded long-review evidence. Page-chrome descriptor
derivation is incremental as well: a tail edit reuses the stable physical-page
prefix, including resolved first/default/even headers, footers, page numbers,
preview text, and jump targets. The deterministic 120-page DOCX workflow proves
that a final-page edit reuses 118 descriptors and derives only the two-page
boundary suffix while page navigation, focus, console, and page errors remain
stable.

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
  scrolling; at phone widths, give Source or Preview the complete workspace
  through one explicit, touch-sized switch.
- Keep the preview a calm reading surface with a bounded line length and one
  visual hierarchy across headings, prose, quotations, code, tables, and task
  lists.

Exit evidence: canonical Markdown fixtures round-trip without semantic drift,
custom structural Extensions have serialization tests, source/visual undo
remains one coherent history, and desktop split plus phone full-pane behavior
pass visual and deterministic E2E regression.

## Priority 3: Spreadsheet

The first WPS-alignment slice now adopts the Writer-proven shared ribbon
behavior without importing Writer's document model. Spreadsheet owns one
command catalog for stable tab order, labels, locations, and shortcut metadata;
Undo and Redo live in quick access; the ribbon compacts by group priority,
persists a double-click collapse, and temporarily exposes a selected tab.
Conditional Formatting is located under Home and Styles, executable ascending
and descending sorts are under Data, and workbook recalculation is available
from Formulas and F9. Focused unit and component tests plus desktop and compact
Playwright coverage prove the command placement and interaction path. The
deterministic A3S Test manifest validates independently and records the same
desktop workflow when a browser capability is installed.
The second slice adds Paste, Cut, and Copy as the first Home group instead of
leaving core clipboard behavior discoverable only through shortcuts and context
menus. Buttons and `Cmd/Ctrl+V`, `Cmd/Ctrl+X`, and `Cmd/Ctrl+C` all cross the
same typed command port, retain the permission-resilient browser/local
clipboard fallback, and restore grid focus after a successful command. Desktop
and compact browser regression executes the complete copy-paste-cut path.
The third slice adds Format Painter beside the clipboard commands. A single
click captures the selected native cell-style pattern for one target, while a
double click keeps the session locked across repeated and cross-sheet targets.
Single-cell targets expand to the source dimensions and larger selections tile
the pattern. Values, formulas, comments, links, and merges stay untouched;
another click or Escape cancels the mode. A 50,000-cell source and target guard,
duplicate-target suppression, one controlled batch per application, accessible
pressed/live state, and copy cursor keep the Web interaction bounded and
predictable. Focused model, hook, controller, and component tests plus desktop
and compact browser regression cover the complete one-shot and locked paths.
The fourth slice adds AutoFilter under Data and Sort and Filter. A single-cell
selection expands to the bounded contiguous current region, while an explicit
multi-row selection keeps its exact coordinates. `Cmd/Ctrl+Shift+L` toggles
the filter and `Alt+ArrowDown` opens the selected header menu. Arrow keys,
Space, Enter, and Escape operate the value checklist and restore grid focus.
The model rejects empty data, merged intersections, and pivot sheets before a
controlled mutation, preserves filter-owned row visibility when toggled, and
round-trips the active range and hidden rows through XLSX. Focused model, hook,
controller, component, import/export, desktop, and compact regression cover the
complete enable, apply, disable, and restore path.
The fifth slice adds Freeze Panes under View and Window. A current-cell freeze
uses the WPS boundary of rows above and columns left, while dedicated top-row
and first-column presets cover the two common one-axis workflows. The active
menu exposes Unfreeze Panes, disables a no-op boundary, supports Arrow, Home,
End, Enter, and Escape, and returns focus to the grid without a delayed focus
guard closing a repeated command. A pure workbook model owns the boundary,
selection retention, no-op checks, and one controlled update. Focused model,
controller, component, focus, XLSX round-trip, desktop, compact, and
schema-validated A3S Test coverage prove the complete freeze and unfreeze path.
The sixth slice exposes the existing row and column structure commands under
Home and Cells through one WPS-familiar Rows and Columns menu. Its six actions
insert selected rows above or below, insert selected columns to the left or
right, or delete selected rows or columns. Availability remains derived from
the live workbook selection and protection state; the ribbon does not create a
second mutation model. Arrow, Home, End, Enter, and Escape operate the menu,
successful commands return focus to the grid, and desktop plus compact
Playwright coverage exercises an insert-and-restore workflow. Focused catalog,
component, and focus tests plus the schema-validated A3S Test manifest keep the
information architecture and command boundary stable.
The seventh slice moves cell merging into Home and Alignment as a WPS-familiar
split control. Its primary action and `Ctrl+M` execute Merge and Center, while
the disclosure menu also offers Merge Cells, Merge Across, Unmerge Cells, and
Unmerge and Fill. Availability comes from the current selection and Fortune's
native merge ranges. Each intent is translated into one `batchCallApis`
transaction, so alignment, native merge state, and fill propagation share one
controlled update. Arrow, Home, End, Enter, and Escape operate the menu;
successful actions return focus to the grid and Escape returns to the exact
disclosure invoker. Focused model, controller, component, and focus tests plus
XLSX round-trip, desktop, compact, and deterministic A3S Test coverage protect
the workflow.
The eighth slice exposes Clear under Home and Editing with Clear All, Clear
Formats, Clear Contents, Clear Comments, and Clear Hyperlinks. Delete and
Backspace execute Clear Contents through the same typed command. Content-only
clearing retains formats, comments, hyperlinks, and merge geometry;
format-only clearing retains values, formulas, comments, hyperlinks, and merge
geometry while subtracting the selected range from direct, border,
conditional, and alternating formats. Clear All removes cell state but retains
merge geometry. Each intent produces one controlled workbook batch, successful
ribbon actions restore grid focus, and Arrow, Home, End, Enter, and Escape
operate the disclosure menu on desktop and compact Web.

- Replace remaining dense and main-thread workbook work with an A3S-owned
  sparse model and virtualized viewport.
- Expand Worker/WASM formula parity, dependency tracking, number formats,
  advanced sorting and filter predicates, validation, charts, pivots, and print
  layout.
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
- Keep the phone Find surface within the viewport with a 40 px input and 40 px
  previous, next, and close actions. The deterministic phone workflow must
  prove an exact cell result and return focus to the grid after Escape.
- Keep workbook task panes as bounded desktop regions and focus-contained phone
  dialogs. At phone widths the ribbon, grid, and worksheet footer remain inert;
  Tab and Shift+Tab stay in the pane; Escape cancels a dirty draft before it
  closes the pane and restores the exact ribbon invoker. Deterministic browser
  evidence covers the modal semantics, focus order, isolation, screenshot,
  accessibility tree, console, and page errors.
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
- Keep the chart inspector as a bounded desktop region and a focus-contained
  dialog whenever the responsive layout overlays the canvas. The ribbon,
  workspace, and status bar remain inert; Tab and Shift+Tab stay inside;
  Escape cancels a dirty field before closing; and close restores the selected
  chart. Deterministic phone evidence covers semantics, focus order, isolation,
  screenshot, accessibility tree, console, and page errors.
- Keep comment review as a docked desktop strip and a full-editor modal on
  phones. The compact surface uses readable review text and 44 px primary
  actions, isolates the ribbon, workspace, and status bar, contains Tab and
  Shift+Tab, lets a dirty comment consume Escape before the panel closes, and
  restores the exact New Comment, View Comments, or comment-marker invoker.
  Deterministic phone evidence covers both invoker paths, modal semantics,
  focus order, isolation, screenshot, accessibility tree, console, and page
  errors.
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
  long-document thumbnail window, and focus-contained phone page drawer. Its
  toolbar-owned trigger must never cover rendered PDF content.
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
