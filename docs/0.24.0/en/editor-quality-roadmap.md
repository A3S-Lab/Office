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
   Office and Traditional Office.
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
Shared Office select menus now keep their compact desktop density while using
44 px option rows and a taller viewport-bounded scroll region on phones. The
deterministic Word font workflow traverses the complete grouped font list with
End, applies the final option while returning focus to the document, and proves
that Escape from a reopened picker restores the exact combobox trigger.

## Priority 1: Word-Compatible Document Editing

Document work is delivered as complete vertical slices in the order below.
New ribbon commands wait unless they are required by the current slice.

### 1. DOCX Round-Trip Foundation

- Move package relationships, styles, numbering, themes, comments, headers,
  footers, notes, fields, and media toward loss-preserving browser state.
- Preserve unsupported package parts and attributes instead of regenerating
  them from visible HTML.
- Build small, reviewable fixtures for every supported semantic and include
  documents produced by both Microsoft Word and Traditional Office.
- Report unsupported or normalized content explicitly before export.

The first package-state slice now registers the imported DOCX source Blob and
retains safe source-only OPC parts byte-for-byte across an edited export. It
merges content-type declarations and relationship graphs without overwriting
generated core parts or reusing conflicting relationship IDs. Export fails
closed when the declared source Blob is unavailable or its SHA-256 fingerprint
does not match the imported package. Invalidated digital signatures and active
VBA, ActiveX, or custom-ribbon parts are deliberately removed and reported.
The second slice preserves relationship-free ignorable attributes, elements, and
structurally valid, non-conflicting `mc:AlternateContent` blocks inside
`word/settings.xml`, including strict/transitional and UTF-8/UTF-16 sources.
Generated settings semantics stay authoritative. The relationship-identity
slice reconnects source font-table metadata only to source-only internal
obfuscated-font payloads, rewriting colliding relationship IDs and rejecting
external, mistyped, duplicate, or payload-path-colliding references. Package
XML decoding accepts strict or transitional UTF-8/UTF-16 input and emits UTF-8.
The retained fonts remain native-DOCX state; browser edit, preview, and PDF
surfaces continue to use registered fonts or substitution. The stable-identity
slice now retains passive, relationship-free extension trees at the roots and
uniquely matched nodes of `word/styles.xml` and `word/numbering.xml`. Style
identity uses type plus style ID. Imported abstract-numbering,
concrete-numbering, and level metadata follows regenerated IDs, while
source-only, duplicate, malformed, and ambiguous one-to-many mappings fail
closed. The drawing-identity slice additionally retains relationship-free
passive extensions from non-OOXML ignorable namespaces on uniquely matched
pictures in document, header, footer, footnote, and endnote parts. Normalized
anchor plus drawing-property IDs survive sanitized body, page-chrome, and note
HTML plus strict/transitional UTF-8/UTF-16 sources. Native note pictures retain
layout, wrapping, crop, and layer metadata through public import/export. Export
repairs missing note-part image relationships, allocates collision-free IDs,
and validates embedded media targets. Microsoft/OOXML semantic namespaces,
source-only, changed, duplicate, namespace-spoofed, relationship-bound, and
ambiguous drawings fail closed. Generated Word geometry and media remain
authoritative; legacy VML, shapes, SmartArt, and drawing-bearing control
wrappers normalize. Equivalent passive extensions now follow uniquely matched,
unchanged paragraphs and their paragraph properties using native `w14:paraId`
plus `w14:textId`. The identity
survives sanitized body and page-chrome HTML. Text edits rotate the version ID,
formatting-only edits and moves retain it, and copies or splits receive new
paragraph IDs. Changed text versions, duplicate identities, relationship-bound
content, and Microsoft/OOXML semantic branches fail closed; generated paragraph
semantics remain authoritative. Stable table scopes now preserve equivalent
passive extensions on `w:tbl`/`w:tblPr`, `w:tr`/`w:trPr`, and
`w:tc`/`w:tcPr`. Rows retain native `w14:paraId` plus `w14:textId`; tables and
cells derive conservative identities from directly owned row and paragraph
IDs. Row text or structural edits rotate the version, formatting-only edits and
moves retain it, copies receive independent IDs, and nested DOM collections no
longer leak into outer-table export. Missing, duplicate, cross-kind,
relationship-bound, and Microsoft/OOXML semantic identities fail closed while
generated table semantics remain authoritative. Native positive footnote and
endnote IDs now survive reorderings, while copies receive fresh IDs. Signed
native comment and reply IDs retain thread parentage and resolved state.
Relationship-free passive extensions follow uniquely matched `w:footnote`,
`w:endnote`, `w:comment`, and `w15:commentEx` roots, and valid `commentsIds`
durable IDs are rebound to regenerated final comment paragraphs. Duplicate,
namespace-spoofed, deleted, relationship-bound, and unsupported modern
reaction/people sidecars fail closed. Text-stable direct paragraphs and runs
inside uniquely matched notes, comments, and replies now retain eligible
passive extensions on paragraph, run, and run-property scopes. Safe unmodeled
note properties survive, while unchanged plain-text comments recover
relationship-free source run segmentation and formatting. Stable hyperlink
wrappers retain safe tooltips, passive metadata, and eligible run formatting.
Generated note destinations remain authoritative; unchanged comments and
replies recover validated HTTP(S), `mailto`, or internal-anchor destinations,
and external relationship IDs are deduplicated or rewritten after collisions.
Text-stable static rich-text and plain-text content controls now recover
eligible inline or contiguous block wrappers, aliases, tags, locking, signed
IDs, Word 2013 appearance and color, end-character formatting, passive
metadata, and stable runs. Footnote and endnote tables now remain native
editable OOXML blocks instead of flattening each row into text. A rich-text
block control can span stable paragraphs, tables, and nested tables when its
row/cell shape, grid spans, merge state, nested block shape, and paragraph text
match uniquely; generated geometry and supported formatting remain
authoritative. Collisions rewrite only the affected control IDs. Text or table
structure edits, duplicate text or properties, missing or malformed hyperlink
relationships, wrong target types or modes, unsafe or relative targets,
combined external-plus-anchor destinations, namespace spoofing, active bindings
or placeholder state, form or nested controls, relationship-bound content,
math, drawing-bearing control wrappers, and unsupported wrappers remain
explicit fail-closed or normalization boundaries.

Exit evidence:

- Import-edit-export-reopen passes in Word and Traditional Office for the declared fixture
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

The Writer Insert ribbon now follows the Traditional Office Pages, Table, Illustrations,
Links, Header and Footer, and Text sequence, including page-number visibility
with the other page-chrome commands. Page Layout exposes direct normal, narrow,
moderate, wide, orientation, paper-size, and one-to-three-column presets. Custom
margins and advanced columns open the matching controlled Page Setup tab. A
deterministic desktop workflow proves live landscape and two-column rendering,
advanced-tab routing, Escape close, accessibility output, and empty browser
diagnostics.

References, Review, and View now follow the Office-oriented information sequence
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
passive label. Clicking it or pressing the Traditional Office `Ctrl+Shift+G` shortcut opens
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

A deterministic Traditional Office Writer layout gate now covers the first declared native
page-parity slice. The fixture uses an A4 page, installed Arial runs, explicit
automatic paragraph spacing, a fixed centered table, direct fills and borders,
and physical cell margins. Traditional Office exports the reference PDF through its desktop
automation API; the browser captures both page surfaces at 794 by 1123 CSS
pixels. The comparison requires one physical page, matching title/body/table
landmarks within one pixel, no more than a 2% thresholded pixel difference,
and a mean absolute channel error no greater than 1.0. The current measured
result is a one-pixel maximum landmark delta, a 1.4612% thresholded difference,
and 0.7308 mean absolute error. Original OOXML line multiples remain the DOCX
round-trip authority instead of being replaced by the browser-only Traditional Office metric.

Four focused matrices now extend that gate beyond Arial and table geometry.
The common-font matrix contains 30 rows across Arial, Times New Roman, Calibri,
Segoe UI, and Microsoft YaHei at two sizes and two automatic spacing multiples.
The CJK matrix contains 36 Latin and Chinese rows across Microsoft YaHei,
SimSun, SimHei, FangSong, KaiTi, and DengXian. The explicit document-grid matrix
contains 18 rows across three fonts and two sizes. Browser layout uses the
measured Traditional Office per-font advance without changing the source OOXML spacing
multiple. DOCX import and export also retain each section's `docGrid` type and
line pitch and each run's `snapToGrid` override. The deterministic browser
suite requires one page, exact semantic metric markers, and empty console and
page-error diagnostics; the Traditional Office comparison rejects a text-band top drift above
three pixels or a consecutive-band advance drift above four pixels. The current
maximum top/advance deltas are 3/1 pixels for common fonts, 2/2 pixels for CJK
fonts, and 1/1 pixels for the document grid.

The fourth matrix adds 30 Latin, CJK, Arabic, Hebrew, and mixed-formatting rows
at two sizes. DOCX import now chooses the Word `ascii`, `hAnsi`, `eastAsia`, or
`cs` font slot from the run text and honors `bCs`, `iCs`, `szCs`, `cs`, `rtl`,
and `rFonts` hints before calculating the Traditional Office line-height factor. The same
fixture proves paragraph RTL direction, complex-script emphasis, mixed-run
bold and italic formatting, one-page pagination, and empty browser diagnostics.
Its real Traditional Office 12.1 reference has 30 matching text bands with a three-pixel
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
styles. Writer now has one command catalog for stable Office-oriented tab and
group placement plus shortcut metadata. Its displayed Traditional Office font-size,
alignment, line-spacing, heading, spelling, field-refresh, comment, and
track-changes shortcuts execute inside the document while host text inputs and
modal surfaces retain their native keyboard behavior. Traditional Office copy-format and
paste-format shortcuts share a permission-free local formatting clipboard with
the one-shot format painter; applying it keeps links and review marks intact
and commits direct character and paragraph formatting in one transaction.
Header and footer surfaces use the same Traditional Office alignment and format-copy commands,
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

- Relationship-bearing objects reopen with working targets in Word and Traditional Office.
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
Character- and paragraph-formatting revisions now form separate review kinds.
Paragraph tracking captures one canonical full-property baseline across a
multi-paragraph command, preserves it through later pending edits, and resolves
the group atomically. Accept, reject, and the source formatting command have
independent undo boundaries. Strict and transitional `w:pPrChange` import and
export preserve the supported old paragraph properties, while malformed or
unsupported records stay structural diagnostics. Browser/Yrs collaboration
converges the distinct immutable decision kind and prevents `suggest` peers
from changing paragraph-revision metadata. Focused A3S Test suites prove both
rejection paths against the public Playground with screenshots, accessibility,
and empty console/page-error evidence.
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
Controlled host replacements now have an explicit review-conflict contract.
The host value remains authoritative and does not enter local undo history,
while stable IDs and reviewed text distinguish harmless range movement from
comment text changes, removed anchors, removed revisions, and revision-kind
reuse. Intentional comment deletion removes both the thread and anchor without
warning; document switches are isolated by `artifactId`. React, Vue, and Web
Component hosts receive the same typed conflict event, an accessible in-editor
warning exposes the condition to users, and orphaned comment records survive
later local edits until the host removes or restores them. Deterministic unit
and adapter regressions cover movement, mutation, removal, persistence,
deduplication, restoration, and document switching.

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

The first Traditional Office-alignment slice now adopts the Writer-proven shared ribbon
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
uses the Traditional Office boundary of rows above and columns left, while dedicated top-row
and first-column presets cover the two common one-axis workflows. The active
menu exposes Unfreeze Panes, disables a no-op boundary, supports Arrow, Home,
End, Enter, and Escape, and returns focus to the grid without a delayed focus
guard closing a repeated command. A pure workbook model owns the boundary,
selection retention, no-op checks, and one controlled update. Focused model,
controller, component, focus, XLSX round-trip, desktop, compact, and
schema-validated A3S Test coverage prove the complete freeze and unfreeze path.
The sixth slice exposes the existing row and column structure commands under
Home and Cells through one Office-familiar Rows and Columns menu. Its six actions
insert selected rows above or below, insert selected columns to the left or
right, or delete selected rows or columns. Availability remains derived from
the live workbook selection and protection state; the ribbon does not create a
second mutation model. Arrow, Home, End, Enter, and Escape operate the menu,
successful commands return focus to the grid, and desktop plus compact
Playwright coverage exercises an insert-and-restore workflow. Focused catalog,
component, and focus tests plus the schema-validated A3S Test manifest keep the
information architecture and command boundary stable.
The seventh slice moves cell merging into Home and Alignment as an Office-familiar
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

The ninth slice establishes a maximum-dimension sparse Spreadsheet contract.
An XLSX worksheet may retain 1,048,576 logical rows and 16,384 logical columns
without allocating its empty range. Virtual navigation, search, selection,
formula dependency scans, formatting, filtering, statistics, collaboration,
and host projection visit only materialized cells. Editing a far blank row
creates exactly that row. Import and export retain the logical dimensions plus
compact validation, protection, passwordless editable, and conditional-format
ranges. Focused unit and round-trip tests and a deterministic A3S Test workflow
cover `Control+End`, read-only far navigation, one-row materialization,
persisted editing, accessibility, console output, and page errors.

The tenth slice completes the first missing Traditional Office font-emphasis path. Home and
Font now exposes Strikethrough with the Traditional Office `Cmd/Ctrl+5` shortcut. The button,
shortcut metadata, and editor keymap come from the Spreadsheet command catalog;
both entry points use the existing typed `setCellFormat` command and Fortune's
native `cl` cell style. Successful commands restore grid focus, create one
undoable controlled update, remain scoped away from host text inputs and modal
surfaces, and continue through Clear Formats and Format Painter without a
parallel style model. Focused catalog, controller, component, desktop browser,
and schema-validated local A3S Test coverage protect the workflow.

The eleventh slice adds cell borders to Home and Font through an Office-familiar
split control. Top, bottom, left, right, no, all, outside, inside, horizontal,
vertical, and diagonal targets share ten native line styles and an exact color
input. The A3S-owned immutable command normalizes the live range, compacts
overlapping native records, preserves malformed vendor state, emits one
controlled update, and restores grid focus without calling Fortune's mutable
border command. Diagonal formatting is capped at 4,096 cells because the
native renderer requires one record per cell. Focused model, controller,
component, focus, desktop and compact Canvas rendering, undo, and
schema-validated local A3S Test coverage protect the complete workflow. GitHub
Actions keeps A3S Test excluded and runs the Playwright rendering contract.

The twelfth slice adds Fill to Home and Editing with Down, Right, Up, and Left
commands. The menu uses Arrow, Home, End, Enter, and Escape, while
`Cmd/Ctrl+D` and `Cmd/Ctrl+R` share the same typed command path and retain grid
focus. Each direction copies the appropriate selection edge through Fortune's
native relative-formula, number-series, and style semantics in one history
record. Preflight rejects multiple selections, merged intersections, pivot
sheets, read-only rows or columns, protected cells, inactive sheets, and more
than 50,000 target cells. The focused adapter satisfies Fortune's sparse-matrix
precondition by creating only missing row arrays inside the selection and
restores the exact sparse shape on a native failure. Unit, controller,
component, focus, desktop, and compact Playwright regressions cover all four
directions, formula and bold-style propagation, one-step Undo and Redo, and
`Ctrl+R` browser-refresh ownership. A schema-validated A3S Test ACL remains a
local release gate; GitHub Actions neither installs nor runs A3S Test.

The thirteenth slice completes the common Traditional Office number-format path under Home and
Number. General, Number, CNY Currency, Accounting, Percentage, Short Date,
Time, Scientific, Fraction, and Text now map to explicit Fortune `ct` codes and
types instead of presenting date, time, currency, scientific, fraction, and
text cells as an undifferentiated Custom state. The grouped picker exposes all
presets, direct Currency and Percentage buttons advertise their shortcuts, and
`Cmd/Ctrl+Shift+~`, `!`, `$`, `%`, `#`, `@`, and `^` share the same typed
selection command. Each action preserves values and formulas, produces one
native controlled update and one Undo record, and retains its exact format code
through XLSX export and reopen. Decimal adjustment skips date, time, fraction,
text, and unknown custom codes and no longer edits digits inside locale tags.
Focused parser, catalog, controller, component, XLSX, desktop, and compact
Playwright coverage protects the complete path. A3S Test remains a local-only
release gate and is absent from GitHub Actions.

The fourteenth slice adds an Office-familiar Cell Styles gallery under Home and
Styles. It owns 17 built-in choices grouped as Common, Data and Model, and
Titles and Totals, renders the native appearance in each preview, and supports
Home, End, and two-dimensional arrow navigation across desktop and compact
layouts. A single typed selection command applies Fortune's native font, fill,
emphasis, and per-cell border properties to at most 10,000 populated or blank
cells. It preserves values, formulas, comments, links, merges, alignment, and
number formats, publishes one controlled value, and creates one Undo record.
Current-state recognition resolves the focused cell's native border records and
does not introduce a second style marker. The XLSX path now round-trips direct
font family, size, color, emphasis, solid fill, alignment, wrap, rotation, four
side and diagonal borders, and the original number-format XF; theme, indexed,
automatic, and tint colors resolve to stable RGB. Focused model, controller,
component, import/export, desktop, compact, accessibility, blank-cell, and Undo
coverage protects the path. The deterministic A3S Test ACL remains a local
release gate; GitHub Actions uses the equivalent Playwright regression only.

The fifteenth slice makes Increase Decimal and Decrease Decimal first-class
Home and Number commands instead of toolbar-only helpers. The command reads
each selected cell's native format independently, so mixed Currency,
Accounting, Percentage, Number, and Scientific selections retain their format
families. Date, time, fraction, text, and unrecognized custom codes remain
unchanged. Equal results compact into rectangular native calls, the complete
gesture commits through one batch and one Undo record, and a 10,000-cell guard
prevents a blank maximum-size selection from becoming dense. Focused model,
catalog, controller, component, and desktop browser coverage protect mixed
format semantics and focus restoration.

The sixteenth slice adds the complete Office-style Format Cells workflow to Home
and Number and the editor-scoped `Cmd/Ctrl+1` shortcut. One modal surface owns
Number, Alignment, Font, Border, Fill, and Protection tabs, including custom
number codes, wrapping and rotation, mixed font emphasis, per-edge and
diagonal borders, no-fill state, and compact locked or hidden ranges. The
dialog captures the exact worksheet, range, active cell, and source cells when
it opens. A later selection change therefore cannot retarget Apply. Mixed
selections retain every untouched field; only explicitly changed fields are
unified. Applying all six tabs emits one controlled `onChange` and creates one
Undo record, while Cancel and Escape emit neither. Dense `data` and sparse
`celldata` worksheets retain their original representation, formatting a blank
cell materializes only that cell, and removing an absent property does not
materialize it. General formatting is bounded to 10,000 cells and diagonal
borders to 4,096 cells. The shortcut remains with host inputs, formula and cell
editing, and modal controls. Focused Rstest model, controller, component,
history, and focus coverage plus desktop and compact Playwright regression
exercise launch, six-tab Apply, reopen, and one-step Undo. A3S Test remains an
optional local-only release aid; GitHub Actions and Pages neither install nor
run it and use Rstest, Playwright, and Cargo instead.

The seventeenth slice adds the Traditional Office AutoSum split command before Fill under Home
and Editing. Sum remains the primary action and owns the editor-scoped `Alt+=`
shortcut; the disclosure menu adds Average, Count, Maximum, and Minimum with
Arrow, Home, End, Enter, and Escape behavior. A single blank target infers the
nearest contiguous numeric or formula run above before falling back to the
left. Explicit totals rows and columns discover numeric/formula axes, preserve
labels and text-only axes, and compact consecutive outputs into native ranges.
Every function writes native formulas, retains target formatting, updates the
formula bar, and commits through one Fortune batch, one controlled `onChange`,
and one Undo record. Preflight rejects occupied eligible targets, multiple
selections, inactive or read-only sheets, merged or protected targets, pivot
sheets, invalid Excel coordinates, and more than 1,000 formula outputs. Sparse
`celldata` is read through indexed coordinates without a dense projection.
Focused planner, command, editor, catalog, ribbon, focus, desktop, and compact
Playwright coverage verifies inference, menu navigation, formula synchronization,
style retention, one-step Undo, and responsive layout. GitHub Actions and Pages
continue to use Rstest and Playwright without installing or invoking A3S Test.

The eighteenth slice replaces the standalone Home and Editing Find button with
the Office-style Find and Select menu. Find retains `Cmd/Ctrl+F`; Go To adds
`Ctrl+G` and `F5` for one bounded direct A1 range, quoted cross-sheet
references, and worksheet- or workbook-scoped names. The shared dialog reports
empty, invalid, ambiguous, hidden, missing, multi-area, unsupported, and
out-of-bounds targets before submission. One view-only navigation primitive
activates the target sheet, sets the exact range, scrolls to the focused cell,
and restores grid focus for Go To and participant following. It emits no
controlled `onChange`, creates no Undo record, and never materializes sparse
cells. Command ownership excludes host inputs, formula and cell editing,
popovers, and modal controls. Focused resolver, shortcut, controller, ribbon,
editor, sparse, and browser regressions protect the workflow. Its deterministic
A3S Test ACL is a local release gate; GitHub Actions and Pages continue to use
Rstest and Playwright only.

The nineteenth slice adds the Office-style Paste split command to Home and
Clipboard. Its primary action pastes All; the disclosure adds quick Values,
Formulas, and Formatting actions plus Paste Special. `Cmd/Ctrl+Alt+V` opens a
single dialog with ten content modes, Add, Subtract, Multiply, Divide, Skip
blanks, and Transpose. A versioned same-editor clipboard snapshot retains
formulas, native styles, comments, validation, explicit protection,
hyperlinks, border fragments, complete in-range merges, and column widths.
External clipboard content remains rectangular TSV and cannot acquire trusted
native formatting. Relative and mixed formula references translate from the
source origin to the destination; absolute axes stay fixed. Planning rejects
pivot worksheets, partial or conflicting merges, protected or read-only
targets, worksheet and XLSX bounds, unsupported external formula state, and
divide by zero before mutation. One accepted request updates at most 50,000
destination cells, emits one controlled workbook value, creates one Undo
record, selects the pasted range, and restores the current grid after a dialog
remount. Focused model, hook, command, shortcut, ribbon, context-menu, dialog,
fallback, desktop, and compact regressions cover the contract. The deterministic
A3S Test suite remains a local-only release gate and records accessibility,
empty console, and empty page-error evidence; Playwright owns bounded visual
evidence while the pinned standalone adapter's CDP screenshot command remains
unreliable.

The twentieth slice adds Hyperlink to Insert and Links with the grid-scoped
`Cmd/Ctrl+K` shortcut. A single accessible Insert/Edit dialog owns Web page,
direct or continuous A1 cell-range, and visible worksheet targets, display-text
changes, and explicit Remove. An A3S-owned immutable model writes Fortune's
native `webpage`, `cellrange`, and `sheet` records without using the vendor
helpers that replace displayed values or hyperlink formatting. Dense `data`
and sparse `celldata` representation, formulas, values, native formatting,
comments, unrelated links, and malformed vendor records remain intact. Apply
and Remove each publish one controlled value and one Undo record, preserve the
captured selection, and restore the exact ribbon trigger or current grid after
a remount. Validation rejects unsafe URLs, malformed or out-of-bounds ranges,
missing or hidden worksheets, protected cells, pivot output, and read-only
views before mutation. Focused model, command, dialog, editor, ribbon, desktop,
and compact regressions cover the workflow. Its deterministic ACL is checked
and run locally with A3S Test 1.0.0, agent-browser 0.26.0, and Web protocol 15;
GitHub Actions and Pages do not install or invoke A3S Test.

The twenty-first slice adds Data Validation to Data and Data Tools. One
accessible dialog snapshots the active worksheet, focused cell, and every live
selection, then creates, edits, or removes list, whole-number, decimal, date,
and text-length rules. Lists accept bounded comma-separated values, one
continuous row or column, and named ranges. Numeric and text-length rules use
typed comparison operators; dates accept ISO, integer Excel serial, and
`DATE(...)` boundaries. The browser stores normalized ISO dates for Fortune,
XLSX import respects the 1900 or 1904 date system, XLSX export writes stable
native formulas, and imported decimal rules no longer reject integers through
Fortune's decimal-only vendor type. One accepted Apply or Remove covers at most
10,000 selected cells, rewrites compact `dataValidationRanges` without
materializing blanks, preserves direct or compact records outside the target,
publishes one controlled value, creates one Undo record, retains every captured
selection, and restores the exact invoker after a remount. Invalid boundaries,
two-dimensional list sources, out-of-bounds regions, protection, merges, pivot
output, and read-only views fail before mutation. Focused model, XLSX, command,
dialog, history, ribbon, desktop, and compact coverage plus a deterministic
local A3S Test 1.0.0 ACL protect the workflow; GitHub Actions and Pages do not
install or invoke A3S Test.

The twenty-second slice adds native Spreadsheet Tables/ListObjects through
Insert and the grid-scoped `Cmd/Ctrl+T` shortcut. A single cell expands to its
finite current region, an explicit multi-row selection remains exact, and one
accessible dialog captures that range plus header semantics. Creation assigns
a workbook-unique name, canonicalizes unique headers, rejects empty,
out-of-bounds, merged, protected, pivot, overlapping-table, worksheet-
AutoFilter, and over-100,000-cell targets, and commits one controlled value and
one Undo record. Selecting a table opens a contextual Table Design ribbon with
editable names, all 60 OOXML Light/Medium/Dark styles, first/last-column
emphasis, row/column stripes, responsive two-dimensional gallery navigation,
and exact trigger focus restoration. Table appearance is resolved for visible
Canvas cells without mutating controlled cell formats. Convert to Range
materializes fills, text emphasis, and borders through a bounded matrix or
sparse `celldata` path before removing the semantic record, and fails closed
when a detected structured reference would break. Row/column commands reconcile
ranges, column identities, filter offsets, and canonical headers; merge and
worksheet AutoFilter operations reject table intersections. Native XLSX table
parts, relationships, content types, names, columns, header/totals flags,
built-in styles, and supported filters round-trip. Yjs uses parent-scoped
creation claims, ID-keyed table records, explicit order, validation, and
field-local patches so two clients' independent design edits converge. Focused
Rstest, desktop and compact Playwright, and a deterministic local-only A3S Test
1.0.0 ACL protect the slice; GitHub Actions does not install or invoke A3S
Test. Structured-reference calculation, calculated columns, complete totals
authoring, slicers, external/query tables, and advanced integrations remain
open.

The twenty-third slice completes the frequent Traditional Office font-size and border
keyboard path under Home and Font. Visible Grow Font and Shrink Font commands
share `Cmd/Ctrl+Shift+.` / `Cmd/Ctrl+]` and `Cmd/Ctrl+Shift+,` /
`Cmd/Ctrl+[` aliases with one typed command. Mixed selections advance each
cell independently through the shared 9–72 point scale, treat unformatted
blanks as 10 points, compact equal results into native rectangles, and commit
one Fortune batch and one Undo record. The 10,000-cell guard rejects oversized
live ranges before reading them. `Cmd/Ctrl+Shift+&` applies a thin black
outside border and `Cmd/Ctrl+Shift+_` clears selected borders through the
existing immutable border command. Ribbon buttons, border-menu `kbd` hints,
and `aria-keyshortcuts` use the same catalog metadata. Host inputs, formula or
cell editing, and modal controls retain their native keys. Focused Rstest,
desktop and compact Playwright, and a deterministic local-only A3S Test 1.0.0
ACL protect the slice; GitHub Actions and Pages do not install or invoke A3S
Test.

The twenty-fourth slice completes native Spreadsheet underline fidelity.
Home and Font now uses an Office-style split control for none, single, double,
single-accounting, and double-accounting underline; Format Cells exposes the
same typed choices and preserves untouched mixed selections. `Cmd/Ctrl+U`
turns every active variant off and enables single underline from the off state,
with one controlled update, one Undo record, and grid-focus restoration.
Fortune `un` values 0–4 remain the only browser state, while native XLSX fonts
round-trip exact OOXML `u` values and SheetJS fallback imports normalize
Boolean, numeric, and named variants. Focused Rstest, desktop and compact
Playwright, and a deterministic local-only A3S Test 1.0.0 ACL protect the
slice with clean console and page-error evidence; GitHub Actions and Pages do
not install or invoke A3S Test.

The twenty-fifth slice completes common Traditional Office text orientation and row/column
visibility. Home and Alignment exposes Horizontal, both 45-degree angles,
stacked Vertical, Rotate Up, and Rotate Down through one keyboard-operated
radio menu. Fortune `rt` stores numeric angles, `tr='3'` stores stacked text,
and XLSX import/export maps OOXML 0–180 plus 255 exactly, including negative
angles represented above 90. One orientation intent accepts at most 10,000
cells and remains one native batch, one controlled update, and one Undo record.
Rows and Columns adds Hide and Unhide plus grid-scoped `Cmd/Ctrl+9`,
`Cmd/Ctrl+0`, `Cmd/Ctrl+Shift+9`, and `Cmd/Ctrl+Shift+0`; preflight rejects
more than 10,000 rows or 1,000 columns before allocating indexes. Host inputs,
formula or cell editing, popovers, dialogs, read-only views, and repeated keys
retain native behavior. Focused Rstest and XLSX tests, desktop and compact
Playwright Canvas pixel evidence, and local A3S Test 1.0.0 semantic regression
protect the slice; GitHub Actions and Pages do not install or invoke A3S Test.

The twenty-sixth slice completes common direct-color absence states and Traditional Office
font-emphasis aliases. Automatic Color deletes Fortune `fc`, No Fill deletes
`bg`, and `Ctrl+2`, `Ctrl+3`, and `Ctrl+4` join Bold, Italic, and Underline
through the shared catalog, typed command path, one-step Undo, and final grid
focus. Imported direct font, solid-fill, and border colors retain validated
theme, indexed, automatic, palette, rendered-RGB, and tint origins through
unrelated edits and Yjs transport. XLSX reconstruction emits those semantic
references only when the rendered value and non-conflicting palette slot still
agree; changed or conflicting values degrade to explicit RGB. Focused Rstest,
a real import/edit/export/reopen fixture, desktop and compact Playwright, and a
local A3S Test 1.0.0 semantic regression protect the slice. GitHub Actions and
Pages do not install or invoke A3S Test.

The twenty-seventh slice preserves independent OOXML diagonal-down and
diagonal-up cell-border directions through Ribbon and Format Cells editing,
Paste Special, built-in styles, Yjs transport, row insertion, XLSX export, and
reopen. A visible-cell Canvas hook paints canonical imported diagonal metadata
without densifying the sheet, while a 4,096-cell bound and one indexed read
keep each mutation finite. One command remains one native batch, one controlled
update, and one Undo record.

The twenty-eighth slice adds static local-date `Ctrl+;` and minute-precision
local-time `Ctrl+Shift+;` entry through Home and Number. Each command targets
only the active cell, preserves unrelated styles, writes the scalar and native
number format in one batch, and creates one Undo record. Read-only, inactive,
protected, merged, pivot, and invalid targets fail before mutation.

The twenty-ninth slice adds exact copy-from-above semantics. `Ctrl+'` copies
the source formula text without translating relative references and falls back
to its scalar when no formula exists. `Ctrl+Shift+'` copies only the calculated
or displayed value and removes the target formula. One native write preserves
all target styles, produces one controlled update and one Undo record, and
synchronizes the formula bar after success. Formula copying fails closed for
external, malformed, array, dynamic-array, data-table, and spill sources;
target protection, merge, pivot, bounds, activity, and editability guards apply
to both commands. Focused Rstest, desktop and compact Playwright, and the
local-only A3S Test 1.0.0 regression verify exact formulas, cached values,
style retention, revision counts, focus, accessibility, and clean browser
diagnostics without adding A3S Test to GitHub Actions or Pages.

The thirtieth slice completes the common Spreadsheet font-dialog keyboard
entry. `Cmd/Ctrl+Shift+F` and `Cmd/Ctrl+Shift+P` route through the existing
typed Format Cells command and captured selection, open Font, and focus the
font-family or font-size combobox respectively. The command catalog remains the
single source for editor keys and `aria-keyshortcuts`; malformed routes fail
closed, native host and modal controls remain outside ownership, and Escape
returns to the exact grid node. Font previews preserve the requested text color
while selecting a dark or light neutral canvas from relative luminance so light
fonts stay visible. Focused Rstest, desktop and compact Playwright, and the
local-only A3S Test 1.0.0 regression protect routing, focus, accessibility,
contrast, and empty browser diagnostics without changing Actions.

- Continue replacing remaining vendor-specific dense paths with the A3S-owned
  sparse traversal contract and virtualized viewport.
- Expand Worker/WASM formula parity, dependency tracking, number formats,
  advanced sorting and filter predicates, validation, charts, pivots, and print
  layout.
- Keep lightweight selection statistics responsive for dense and sparse sheets.
- Expand the native format engine from the common preset set into host-selected
  currencies, locale-aware dates, custom conditional sections, and explicit
  unsupported-format diagnostics without rewriting imported codes.
- Keep font family, vertical alignment, and text wrapping available from the
  primary ribbon and backed by the native workbook cell-style model.
- Keep worksheet lifecycle actions in the shared Office interaction system:
  destructive deletion uses a safe default, invalid names remain editable with
  local accessible feedback, and temporary menus and dialogs restore focus.
- Keep phone worksheet renaming focused on the active task: unrelated footer
  tools yield the available width, invalid feedback expands into a readable
  second row, and successful or cancelled edits restore the normal workbook
  status controls without covering the grid.
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
Excel and Traditional Office.

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
their semantics through PowerPoint and Traditional Office round trips, and one gesture creates
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
   Word/Traditional Office round-trip evidence.
3. Update public types and Playground documentation only after the contract is
   stable.
4. Move to the next Word slice. Do not start a lower-priority editor milestone
   merely because its UI is easier to demonstrate.
5. Revisit the priority order only with measured user demand, compatibility
   risk, and engineering evidence.
