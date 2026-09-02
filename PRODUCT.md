# A3S Office Product Direction

A3S Office is a browser-first editing component for teams that need familiar
office workflows inside their own products. It should let an experienced
Traditional Office user find routine commands, predict keyboard behavior, and
move between files without relearning the editor shell.

## Traditional Office Alignment Contract

Alignment targets information architecture and operating habits before visual
similarity. A3S Office uses its own product identity and implementation while
matching the command locations, tab order, shortcut expectations, disclosure
patterns, focus behavior, and responsive priorities that make Traditional
Office familiar.

The delivery order is:

1. Web before native shells.
2. Writer before Spreadsheet, Presentation, and PDF.
3. Daily editing paths before long-tail command coverage.

Writer establishes the shared interaction baseline: a quick-access row,
priority-aware and collapsible ribbon, temporary tab expansion, stable command
metadata, Office-oriented shortcuts, contextual tools, and a useful status bar.
Each later editor adopts that baseline only where it matches the editor's own
file model and workflows.

## Current Spreadsheet Milestone

Spreadsheet is the next vertical track after the Writer shell baseline. Its
first milestone keeps the Traditional Office tab sequence, moves Undo and Redo
into quick access, places Conditional Formatting under Home and Styles, exposes
real sort
commands under Data, and makes workbook recalculation discoverable and
executable through F9. The shared ribbon must remain adaptive, persistently
collapsible, and temporarily expandable without moving the workbook surface.
The second milestone makes the existing permission-resilient Paste, Cut, and
Copy behavior discoverable in the Home clipboard group and routes both ribbon
clicks and Traditional Office shortcuts through the same typed command port.
The third milestone adds Format Painter beside those clipboard commands. A
single click applies the captured native cell-format pattern once, a double
click locks it across repeated or cross-sheet targets, and another click or
Escape exits without copying values, formulas, comments, links, or merges.
Large source and target ranges fail with recoverable feedback, while desktop
and compact Web retain the same pressed state, copy cursor, and grid focus.
The fourth milestone adds AutoFilter under Data and Sort and Filter. A
single-cell selection expands to its finite current region, while an explicit
multi-row selection remains exact. `Cmd/Ctrl+Shift+L` toggles the filter and
`Alt+ArrowDown` opens the active header menu; arrows, Space, Enter, and Escape
provide a complete keyboard path. Empty data, merged ranges, and pivot sheets
fail without mutating the workbook, and filter ranges plus filtered row state
survive controlled updates and XLSX round trips.
The fifth milestone adds Freeze Panes under View and Window. The current-cell
command freezes rows above and columns left of the selection, matching the
Traditional Office workflow, while dedicated presets freeze the top row or
first column and an
active worksheet exposes Unfreeze Panes. One controlled update retains the
selection, the menu supports Arrow, Home, End, Enter, and Escape with exact
grid-focus restoration, and the boundary survives XLSX export and reopen.
The sixth milestone makes the existing row and column structure commands
discoverable under Home and Cells through one Office-familiar Rows and Columns
menu. Users can insert selected rows above or below, insert selected columns to
the left or right, and delete selected rows or columns. Each item derives its
availability from the live selection, uses the shared typed command port, and
returns focus to the grid for uninterrupted keyboard editing.
The seventh milestone moves cell merging into the Office-familiar Home and
Alignment group. The split control keeps Merge and Center as the primary
action, while its menu exposes Merge and Center, Merge Cells, Merge Across,
Unmerge Cells, and Unmerge and Fill. `Ctrl+M` follows the same typed command
path. Every action reads the native Fortune merge model, emits one controlled
workbook batch, and restores grid focus; the menu retains Arrow, Home, End,
Enter, and Escape behavior on desktop and compact Web.
The eighth milestone puts Clear under Home and Editing. Its menu exposes Clear
All, Clear Formats, Clear Contents, Clear Comments, and Clear Hyperlinks, while
Delete and Backspace remain the fast path for Clear Contents. Content-only
clearing retains formats, comments, links, and merge geometry; format-only
clearing retains content, comments, links, and merge geometry while removing
direct, border, conditional, and alternating formats. Every mode uses one typed
command and one controlled workbook batch, then returns focus to the grid.
The ninth milestone establishes the maximum-dimension sparse workbook contract.
A sheet retains 1,048,576 logical rows and 16,384 logical columns without
allocating its empty range, while navigation, editing, formatting, filtering,
statistics, collaboration, import, and export visit only materialized cells.
The tenth milestone completes the first missing Traditional Office
font-emphasis path by placing Strikethrough in Home and Font. The ribbon button
and `Cmd/Ctrl+5`
share the existing typed cell-format command, preserve grid focus, produce one
undoable controlled update, and expose the active native `cl` style without a
second formatting model.
The eleventh milestone adds an Office-familiar cell-border split control to Home
and Font. It supports top, bottom, left, right, no, all, outside, inside,
horizontal, vertical, and diagonal borders with ten native line styles and an
exact color input. One immutable A3S-owned command compacts overlapping native
range records, preserves malformed vendor records, emits one controlled update,
and returns focus to the grid. Diagonal writes stay bounded to 4,096 cells and
undo through the shared workbook history.
The twelfth milestone adds Fill under Home and Editing. Fill Down, Right, Up,
and Left copy the appropriate selection edge through Fortune's native formula
and style semantics; `Cmd/Ctrl+D` and `Cmd/Ctrl+R` provide the Traditional
Office fast paths.
One editable, unmerged, unprotected, non-pivot range may fill at most 50,000
target cells. The adapter materializes only missing row arrays inside that
range before calling Fortune and restores the original sparse shape if the
native command fails. A successful fill remains one undoable operation and
returns focus to the workbook grid.
The thirteenth milestone completes the common Traditional Office number-format
path under Home and Number. General, Number, CNY Currency, Accounting,
Percentage, Short Date,
Time, Scientific, Fraction, and Text are explicit native presets instead of
collapsing into Custom. The format picker, direct Currency and Percentage
buttons, and the seven standard `Cmd/Ctrl+Shift` formatting shortcuts share one
typed `ct` command. Each intent updates only the selected cells' format code and
Fortune value type, preserves values and formulas, creates one undo record, and
round-trips its exact XLSX format code. Imported combined date-time and other
unmodeled codes remain visible as disabled Custom state rather than being
silently rewritten.
The fourteenth milestone adds the Office-familiar Cell Styles gallery under Home
and Styles. Its 17 built-in choices are grouped as Common, Data and Model, and
Titles and Totals, with a faithful visual preview and two-dimensional keyboard
navigation. One command applies native Fortune font, fill, emphasis, and
per-cell border properties to at most 10,000 populated or blank cells while
preserving values, formulas, links, comments, merge geometry, alignment, and
number formats. The selected style is inferred from those native properties
instead of a private marker, emits one controlled update, and creates one Undo
record. XLSX import and export retain direct font, fill, alignment, wrapping,
rotation, border, and original number-format XF state; theme, indexed, and tint
colors resolve to stable RGB for the browser model.
The fifteenth milestone hardens Increase Decimal and Decrease Decimal as
catalogued Home and Number commands. Mixed selections retain each cell's
currency, accounting, percentage, number, or scientific format family instead
of inheriting the focused cell's code. Identical results compact into native
rectangles, one command emits one workbook batch and one Undo record, and date,
time, fraction, text, and unknown custom formats remain untouched. A 10,000-cell
bound prevents blank maximum-size ranges from becoming dense.
The sixteenth milestone adds the complete Office-style Format Cells workflow under
Home and Number with the editor-scoped `Cmd/Ctrl+1` shortcut. Number, Alignment,
Font, Border, Fill, and Protection share one dialog that preserves untouched
mixed values, captures its exact source range, retains dense or sparse worksheet
representation, and publishes one controlled update and one Undo record. The
general command accepts at most 10,000 cells and diagonal borders retain their
stricter 4,096-cell bound; host inputs, formula or cell editing, and modal
controls keep their native shortcut behavior.
The seventeenth milestone adds the Traditional Office AutoSum split command
before Fill under Home and Editing. Sum is the primary action and `Alt+=` fast
path; the menu adds
Average, Count, Maximum, and Minimum with complete Arrow, Home, End, Enter, and
Escape behavior. One blank target infers the nearest contiguous numeric or
formula block above before looking left, while an explicit totals row or column
can emit multiple formulas without replacing label or text axes. The command
preserves target formatting, rejects occupied, merged, protected, read-only,
inactive, or pivot targets, caps one gesture at 1,000 formula cells, keeps sparse
worksheets sparse, and commits one native batch, one controlled update, and one
Undo record.
The eighteenth milestone replaces the standalone Home and Editing Find button
with the Office-style Find and Select menu. Find retains `Cmd/Ctrl+F`; Go To adds
`Ctrl+G` and `F5` for direct A1 cells, continuous ranges, quoted cross-sheet
references, and worksheet- or workbook-scoped names. Invalid, ambiguous,
hidden, multi-area, unsupported, and out-of-bounds targets fail inside the
shared dialog. Successful navigation activates the target sheet, selects and
scrolls to the focused cell, and restores grid focus without publishing
`onChange`, adding Undo history, or materializing sparse cells. Host inputs,
formula or cell editing, popovers, and modal controls retain their native keys.
The nineteenth milestone adds the Office-style Paste split command and Paste
Special dialog to Home and Clipboard. Quick commands cover All, Values,
Formulas, and Formatting; `Cmd/Ctrl+Alt+V` opens ten content modes plus Add,
Subtract, Multiply, Divide, Skip blanks, and Transpose. A same-editor rich
snapshot retains formulas, styles, comments, validation, protection,
hyperlinks, borders, merges, and column widths, while external clipboard text
uses a bounded TSV snapshot. Relative and mixed formula references translate
at the destination. One paste publishes one controlled workbook value and one
Undo record, restores the current grid after a dialog-driven remount, and caps
the destination at 50,000 cells. Pivot sheets, partial merge intersections,
protected or read-only targets, out-of-bounds writes, unsupported formula
state, and divide-by-zero operations fail before mutation.
The twentieth milestone adds Hyperlink under Insert and Links with the
grid-scoped `Cmd/Ctrl+K` shortcut. One accessible Insert/Edit dialog supports
HTTP(S), one direct A1 cell or continuous range, and visible worksheet targets,
plus explicit Remove. The command captures the exact active cell and selection,
uses Fortune's native `webpage`, `cellrange`, and `sheet` records through an
A3S-owned immutable update, and preserves dense or sparse representation,
content, formulas, formatting, comments, unrelated hyperlinks, and malformed
vendor records. One Apply or Remove emits one controlled value and one Undo
record, then restores the exact ribbon invoker or current grid after a remount.
Read-only, protected, pivot, missing, hidden, unsafe, invalid, and out-of-bounds
requests fail before mutation; host inputs, formula or cell editing, and modal
controls retain their native `Cmd/Ctrl+K` behavior.
The twenty-first milestone adds Data Validation under Data and Data Tools. One
accessible dialog creates, edits, or removes list, whole-number, decimal, date,
and text-length rules across one or more captured ranges, with optional input
messages and explicit invalid-input blocking. List sources accept bounded
comma-separated values, one continuous row or column, or a named range. Date
boundaries accept ISO dates, Excel serials, and `DATE(...)`, normalize to a
Fortune-compatible ISO value, honor the XLSX 1900 or 1904 date system, and
export as stable native formulas. Apply and Remove each publish one controlled
workbook value and one Undo record while retaining dense or sparse cells,
values, formulas, formatting, comments, hyperlinks, unrelated direct rules,
and compact ranges outside the selection. Preflight rejects more than 10,000
selected cells, malformed boundaries, two-dimensional list sources,
out-of-bounds ranges, protection, merges, pivot output, and read-only views.
Cancel, Escape, Apply, and Remove restore the exact ribbon invoker or current
grid after a controlled remount. Version 0.41.0 aligns direct and formula-bar
editing with the native error-alert contract: Stop blocks with an accessible
notice, while Warning and Information offer explicit keep-or-return branches.
The shared validation hook preserves authored titles, messages, and the current
input, commits an accepted invalid value once through the controlled API, and
keeps selection, focus, Undo, and collaboration on the same bounded path.
Version 0.42.0 adds the follow-up custom-formula rule to that same dialog and
controlled edit boundary. Authors can enter a local formula with an optional
`=` prefix; relative references are anchored to each selected range, the
proposed value replaces the active cell during evaluation, and common local
functions plus sheet-qualified cell/range references are checked synchronously.
The 255-character formula and 1,024-referenced-cell limits, fail-closed
handling for external/whole-row/column/missing-sheet/uncached-formula inputs,
and native XLSX round trip are explicit. The public Data Validation template
now includes a required-owner custom rule, while paste and object-level batch
writes retain their separate preflight boundary.
The twenty-second milestone adds native Spreadsheet Tables/ListObjects under
Insert and Tables with the grid-scoped `Cmd/Ctrl+T` shortcut. A single-cell
selection expands to its finite current region, an explicit multi-row
selection remains exact, and one accessible dialog captures the range and
header-row choice. Creation assigns a workbook-unique table name, canonicalizes
unique column headers, rejects overlapping tables, worksheet AutoFilters,
merges, protection, pivots, invalid ranges, and more than 100,000 cells, and
publishes one controlled workbook value and one Undo record. Selecting a table
opens the contextual Table Design ribbon with editable names, 60 native OOXML
Light, Medium, and Dark styles, first/last-column emphasis, row/column stripes,
and Convert to Range. Rendering resolves table appearance only for the visible
Canvas cells without rewriting cell formats; conversion materializes that
appearance into native sparse-safe cell styles and borders before removing the
ListObject. Row and column structure commands reconcile ranges, columns,
filters, and canonical headers. XLSX import and export preserve native table
parts, relationships, content types, names, ranges, built-in styles, and the
supported filter criteria. Yjs stores tables as ordered ID-keyed records with
creation claims and field-local conflict handling, so independent design edits
from two browser clients converge. The follow-up structured-reference slice
resolves bounded table names/display names, contiguous columns,
header/data/totals selectors, worksheet-qualified references, and table-local
calculated-column formulas through the shared Rust/WASM and JavaScript
calculation paths. The calculated-column slice now infers a shared current-row
formula, persists it as typed table metadata, and fills only newly inserted
body rows. Existing values, formulas, and manual exceptions remain
authoritative; conflicting formulas fail closed. Native XLSX
`<calculatedColumnFormula>` metadata and Yjs table records round-trip with the
editable leading `=` form. The totals-row follow-up adds one Table Design
surface for enabling the row and assigning each column a label, one of ten
native aggregates, or a bounded custom formula. Native aggregates author
filtered-row-aware `SUBTOTAL` formulas. Dense and sparse worksheets, direct
cell edits, table structure and name changes, Yjs convergence, and native XLSX
`totalsRowFunction`, `totalsRowLabel`, and `totalsRowFormula` round trips share
one fail-closed model. Slicers and external/query tables remain explicit
compatibility gaps.
The twenty-third milestone completes the frequent Traditional Office font-size
and border keyboard path under Home and Font. Grow Font owns `Cmd/Ctrl+Shift+.` and
`Cmd/Ctrl+]`; Shrink Font owns `Cmd/Ctrl+Shift+,` and `Cmd/Ctrl+[`; Outside
Borders owns `Cmd/Ctrl+Shift+&`; and Clear Borders owns
`Cmd/Ctrl+Shift+_`. Mixed-size selections advance each cell through the shared
9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, and 72 point scale instead
of flattening to the focused cell. Equal results compact into native
rectangles, blank cells start from 10 points, and one gesture covers at most
10,000 cells before one Fortune batch and one Undo record. Visible ribbon
buttons, border-menu hints, and `aria-keyshortcuts` expose the same catalog
metadata on desktop and compact Web. Grid ownership excludes host inputs,
formula or cell editing, and modal controls.
The twenty-fourth milestone preserves the complete native Spreadsheet
underline family. Home and Font exposes an Office-style split control for none,
single, double, single-accounting, and double-accounting underline; Format
Cells uses the same typed states and retains untouched mixed selections.
`Cmd/Ctrl+U` turns every active variant off instead of downgrading an advanced
style to single, while an unformatted cell still enables single underline.
Fortune's native `un` values remain the only browser model, and XLSX fonts
round-trip the exact OOXML `u` value, including an omitted value as single.
Every ribbon, dialog, or shortcut intent emits one controlled update and one
Undo record, restores grid focus, and keeps host inputs, the formula bar,
active cell editing, and modal controls outside editor shortcut ownership.
The twenty-fifth milestone completes the common Traditional Office
text-orientation and row/column visibility path. Home and Alignment exposes
Horizontal, Angle
Counterclockwise, Angle Clockwise, Vertical, Rotate Up, and Rotate Down through
one keyboard-accessible radio menu. Numeric rotation uses Fortune `rt`, stacked
vertical text uses `tr='3'`, and native XLSX import/export maps the complete
OOXML 0–180 plus 255 range without reversing negative angles. One orientation
intent covers at most 10,000 cells and emits one Fortune batch and one Undo
record. Home and Cells now exposes Hide and Unhide for rows and columns, while
grid-scoped `Cmd/Ctrl+9`, `Cmd/Ctrl+0`, `Cmd/Ctrl+Shift+9`, and
`Cmd/Ctrl+Shift+0` share the same typed commands. Visibility mutation is capped
before allocation at 10,000 rows or 1,000 columns and never captures host
inputs, active editing, popovers, modal controls, read-only views, or repeated
keydown events.
The twenty-sixth milestone completes the common Traditional Office direct-color
reset and font-emphasis alias path. Home and Font adds Automatic Color for text and No
Fill for cell backgrounds; each action removes the direct Fortune `fc` or `bg`
property instead of replacing it with a hard-coded RGB value. Grid-scoped
`Ctrl+2`, `Ctrl+3`, and `Ctrl+4` join the existing Bold, Italic, and Underline
commands, with one catalog owning visible copy, `aria-keyshortcuts`, and the
editor keymap across desktop and compact layouts. Imported direct XLSX font,
solid-fill, and border colors retain theme, indexed, automatic, and tint origin
metadata across unrelated cell-style edits and Yjs transport. Export rebuilds
matching source palette slots and emits semantic references only while their
rendered RGB is unchanged; edited values or conflicting palette identities
fall back to explicit RGB instead of exporting a misleading reference.
The twenty-seventh milestone preserves both native OOXML diagonal-border
directions independently. Home and Font exposes Diagonal Down and Diagonal Up
as separate catalogued commands, Format Cells can compose a crossed border,
and Paste Special, built-in styles, Yjs collaboration, row insertion, XLSX
export, and reopen retain the exact `diagonalDown` and `diagonalUp` flags with
their shared native line. A3S paints the canonical per-cell metadata only for
visible Canvas cells because Fortune renders range slashes but not imported
cell slash records; ordered native slash and no-border ranges can still
override earlier metadata. Each command remains limited to 4,096 cells, reads
existing directions in one indexed pass, emits one controlled update and one
Undo record, and preserves legacy Fortune `border-slash` as diagonal-down.
The twenty-eighth milestone adds static current-date and current-time entry to
Home and Number through one discoverable Date and Time menu. Grid-scoped
`Ctrl+;` writes the user's local calendar day as a normalized Excel 1900
serial with `yyyy-MM-dd`; Control-only `Ctrl+Shift+;` writes local hour and
minute as a day fraction with `hh:mm`, dropping seconds and milliseconds. Each
intent targets only the active cell while retaining a broader selection,
replaces an old formula through Fortune's value API, preserves unrelated
styles, and sends the value plus native `ct` format through one workbook batch
and one Undo record. Read-only, inactive, locked, protected, merged, pivot, and
out-of-XLSX-bounds targets fail before mutation. Host inputs, formula or cell
editing, modals, composing events, repeated keydown, and Meta-only gestures
remain outside shortcut ownership.
The twenty-ninth milestone adds exact copy-from-above semantics to the
Spreadsheet command runtime. Grid-scoped `Ctrl+'` copies the source formula
text without translating relative references and falls back to its scalar when
the cell above is not a formula. `Ctrl+Shift+'` copies only the calculated or
displayed source value and removes any target formula. Each command writes only
the active cell through one native batch, preserves the broader selection and
all target formatting, publishes one controlled update, and creates one Undo
record. Top-row, out-of-bounds, inactive, read-only, protected, merged, pivot,
array, dynamic-array, data-table, external, and malformed-formula cases fail
closed at the appropriate formula or target boundary. Host inputs, the formula
bar, active cell editing, popovers, dialogs, composing events, and repeated
keydown remain outside shortcut ownership.
The thirtieth milestone completes the common keyboard entry into Spreadsheet
font formatting. `Cmd/Ctrl+Shift+F` opens the existing Format Cells dialog on
the Font tab with the font-family control focused, while
`Cmd/Ctrl+Shift+P` focuses font size. Both aliases route through the same typed
`openFormatCells` command, captured selection snapshot, modal focus scope, and
catalog metadata as `Cmd/Ctrl+1`; no secondary font state or alternate Apply
path is introduced. Malformed routes fail closed, host inputs, the formula bar,
cell editing, composing or repeated events, and existing dialogs retain native
keys, and Escape restores the exact grid node. The font preview preserves the
requested text color while selecting a dark or light canvas from relative
luminance so white and other light fonts remain readable without falsifying
their style.
The thirty-first milestone preserves native Spreadsheet rich-text cells across
XLSX import, cell-wide font formatting, export, and reopen. Shared strings and
inline strings retain ordered runs for font family, size, RGB color, bold,
italic, strikethrough, and all supported underline variants. Theme, indexed,
automatic, and tint color identities remain semantic only while the rendered
value still matches its origin; an edited color exports as explicit RGB.
Imported runs use Fortune's existing `ct.s` representation with
`ct.t='inlineStr'`, and export emits native inline-string runs with exact
leading or trailing whitespace preservation. Format Cells applies one font
intent immutably to every visible run, keeps one controlled update and one Undo
record, and number-format edits retain the rich-string type. Parsing and
materialization fail closed beyond 32,767 characters, 512 runs per cell,
10,000 cells, or 100,000 total runs. Focused import/export/reopen, scanner,
formatting, desktop/compact Playwright, and local-only A3S Test 1.0.0 evidence
cover this boundary; at that milestone, partial-run authoring remained a
documented gap.
The thirty-second milestone adds bounded partial-run formatting to the live
Spreadsheet cell editor and formula bar. A non-collapsed text selection routes
font family, size, color, bold, italic, strikethrough, and underline through a
single optional rich-text command port before whole-cell formatting. Plain
strings become native inline-string runs only when the command succeeds;
existing runs split and coalesce without flattening unaffected content, and
unchanged semantic color origins remain eligible for exact XLSX export. Direct
ribbon actions restore the selected text when the editor remains active. The
model rejects malformed coordinates, invalid values, UTF-16 surrogate splits,
formula cells, and results beyond 32,767 characters or 512 runs. One immutable
controlled update preserves host history and collaboration ownership.
The thirty-third milestone preserves native rich-text runs through direct
formula-bar and F2 insertion or deletion. The controlled source runs are the
only formatting authority: an exact cell operation must authenticate the
coordinate and emitted text before one contiguous replacement can inherit the
replaced or preceding run. Untouched boundaries and theme, indexed, automatic,
or tint color identities remain stable; structural operations and text-stable
focus callbacks cannot infer an edit. One commit remains one controlled host
revision and one Undo record, with the existing UTF-16, 32,767-character, and
512-run limits enforced before publication.
The thirty-fourth milestone adds authenticated formatted rich-text paste to
formula-bar and F2 selections. A paste event stages one bounded, one-shot
authority over the exact worksheet object, coordinate, controlled source text,
UTF-16 selection, plain clipboard text, and sanitized font runs. Projection
consumes it only when an authenticated cell operation emits the predicted
replacement. Untouched source runs keep their native semantic identities;
pasted font family, size, explicit RGB color, bold, italic, strikethrough, and
underline apply only inside the inserted range. Eligible plain and empty cells
become native inline strings after the same proof. Clipboard HTML is discarded
after parsing and capped at 256,000 characters; mismatched, malformed,
structural, or oversized input falls back to inherited or plain text without
authorizing styles.
The thirty-fifth milestone preserves every native non-solid OOXML Spreadsheet
pattern fill. A dedicated cell field carries the exact pattern identity,
rendered foreground/background colors, and validated theme, indexed,
automatic, or tint origins while Fortune `bg` remains the background-color
projection. The visible-cell Canvas hook paints the base and procedural overlay
behind text without rasterizing hidden cells; conditional and table fills stay
authoritative. Format Painter, Paste Special Formats, unrelated formatting,
Yjs transport, export, and reopen retain the metadata, while an explicit fill,
No Fill, Clear Formats, or built-in Cell Style clears it. Unsupported or
malformed metadata fails closed, and conflicting semantic palettes export
literal RGB. Focused Rstest, desktop/compact Canvas evidence, and a local-only
A3S Test 1.0.0 gallery gate cover all 17 pattern identities without adding A3S
Test to GitHub Actions or Pages.
The thirty-sixth milestone preserves native OOXML Spreadsheet gradient fills.
Linear fills retain their exact angle; path fills retain validated inner-
rectangle geometry; both keep two through 256 ordered stops with rendered RGB
and optional theme, indexed, automatic, or tint identities. Fortune `bg`
projects the first stop and invalidates stale metadata after an explicit color
change. The visible-cell Canvas path uses native linear gradients or at most 96
rectangular path contours behind text, while conditional and table fills remain
authoritative. Format Painter, Paste Special Formats, unrelated formatting,
Yjs transport, exact export, and reopen preserve the field. Explicit fills,
No Fill, Clear Formats, and built-in Cell Styles clear it. Invalid geometry,
colors, order, type, or stop budgets fail closed; palette conflicts emit
literal RGB. Focused Rstest, desktop/compact Playwright, and a local-only A3S
Test 1.0.0 gallery gate cover the slice without changing Actions or Pages.
The thirty-seventh milestone turns preserved native fills into one complete
Format Cells authoring surface. A typed fill union covers none, solid, every
native OOXML pattern, and linear or path gradients without introducing a
parallel style state. Pattern foreground/background colors and gradient angle,
inner-rectangle geometry, and two through 256 ordered stops are editable with
an exact Canvas preview; midpoint insertion interpolates the neighboring
colors. Switching modes retains inactive drafts, mixed selections remain
untouched until an explicit fill edit, and invalid geometry or stop order
blocks Apply. Editing one semantic color invalidates only that color origin,
while untouched native identities remain exportable. One Apply still produces
one controlled workbook update and one Undo record, and authored patterns and
gradients survive XLSX export and reopen. Focused Rstest, desktop/compact
Playwright authoring flows, and local-only A3S Test 1.0.0 suites cover Apply,
reopen, Undo, responsive containment, accessibility, and empty browser
diagnostics without adding A3S Test to Actions or Pages.
The thirty-eighth milestone returns to Writer with native all-caps and
small-caps character effects. One typed `none | all-caps | small-caps` state
keeps the effects mutually exclusive across the Home ribbon, body and page
chrome editors, Format Painter, formatting-revision snapshots, Undo, and the
standard `Cmd/Ctrl+Shift+A` and `Cmd/Ctrl+Shift+K` shortcuts. Semantic text is
never rewritten: CSS renders the effect, while DOCX import, export, reopen, and
`w:rPrChange` use native `w:caps` or `w:smallCaps` properties and explicit
resets. Paragraphs containing either effect deliberately use browser line
measurement because uppercase conversion can change the UTF-16 glyph stream
and small caps changes font metrics; eligible unaffected paragraphs continue
through Worker/WASM shaping. Focused Rstest, desktop Playwright, and the
local-only A3S Test shortcut suite cover mutual exclusion, native round trips,
Undo, accessibility, and empty browser diagnostics without changing Actions or
Pages.
The thirty-ninth milestone completes Writer's native underline model. One typed
mark preserves all 18 WordprocessingML values, direct RGB or theme/tint/shade
color identity, and an explicit `none` reset across body text, page chrome,
Format Painter, formatting revisions, Undo, DOCX export, and reopen. Accessible
split controls expose the complete style family and Automatic Color, while
`Cmd/Ctrl+U`, `Cmd/Ctrl+Shift+D`, and `Cmd/Ctrl+Shift+W` cover the common single,
double, and words-only workflows. Browser CSS provides bounded visual
projections without flattening the native metadata. Underline remains a paint
effect, so the change does not exclude otherwise eligible paragraphs from the
Worker/WASM layout path. Focused and full Rstest, formatting, lint, typecheck,
and the production package build form the release evidence.
The fortieth milestone restores Writer's native strikethrough identity instead
of flattening `w:strike` and `w:dstrike` into one boolean. One typed
`none | single | double` mark preserves independent inherited flags and explicit
resets through body text, page chrome, Format Painter, tracked formatting
revisions, Undo, DOCX export, and reopen. Accessible split controls expose all
three states in the Home ribbon, selection toolbar, and header/footer ribbon;
the compact page-chrome toolbar keeps a direct toggle. Export writes both native
flags so a child run can override either inherited state, while import resolves
double above single deterministically. The editor deliberately claims no direct
strikethrough shortcut and disables TipTap's unrelated `Mod+Shift+S` binding.
Strikethrough remains paint-only, so eligible paragraphs retain Worker/WASM
layout. Focused Rstest, Playwright, local-only A3S Test, formatting, lint,
typecheck, and production builds form the release evidence without adding A3S
Test to Actions or Pages.
The forty-first milestone adds Writer's native East Asian emphasis-mark model.
One closed `none | dot | comma | circle | underDot` state preserves all five
`w:em` values, including an explicit `none` that overrides inherited formatting
without conflating it with removal of the direct property. The shared
`Cmd/Ctrl+D` advanced font dialog keeps mixed selections independent, offers
Follow style and every native value, restores the captured selection and focus,
and commits with the other character properties through one transaction and one
Undo record. Body text, headers, footers, footnotes, endnotes, inherited styles,
Format Painter, tracked formatting revisions, strict/transitional DOCX import,
export, and reopen share the same typed model. Canonical DOM attributes project
the four visible states through CSS; those paragraphs deliberately use
browser-authoritative line measurement because the marks extend outside the
text line, while a computed explicit none remains eligible for Worker/WASM
layout. Malformed or namespace-spoofed properties fail closed. Focused Rstest,
desktop/compact Playwright, and a local-only A3S Test 1.0.0 suite form the
release evidence without adding A3S Test to Actions or Pages.
The forty-second milestone adds Writer's native hidden-text model. One
three-state TextStyle attribute distinguishes inheritance, native `w:vanish`,
and an explicit `w:vanish w:val="0"` reset. The shared `Cmd/Ctrl+D` dialog keeps
mixed selections untouched until the semantic checkbox changes, while the
standard `Cmd/Ctrl+Shift+H` shortcut toggles the same command and remains one
Undo step. Body text, page chrome, notes, inherited styles, Format Painter,
tracked formatting revisions, strict/transitional DOCX import, exact export,
and reopen retain the state; temporary export markers restore the original
character style instead of replacing its identity. Hidden text is suppressed
by default and in every read-only preview or PDF snapshot. The View ribbon
reveals it only in editable Writer surfaces with a dotted underline. Malformed,
duplicated, child- or text-bearing, extra-attribute, and namespace-spoofed
properties fail closed; unchanged comment XML is only source-preserved and is
not presented as rich comment editing. Focused Rstest, desktop/compact
Playwright, and the local-only pinned A3S Test suite cover authoring, reveal,
shortcut, Undo, accessibility, and empty browser diagnostics without changing
Actions or Pages.
The forty-third milestone adds Writer's native outline, shadow, emboss, and
imprint effects as four nullable TextStyle properties. Missing direct values
inherit, enabled values map to exact `w:outline`, `w:shadow`, `w:emboss`, or
`w:imprint` leaves, and explicit false values remain native resets. Outline and
shadow may coexist; enabling emboss or imprint clears every conflicting effect
inside the same advanced-font transaction and one Undo record. Document
defaults, paragraph and character styles, body text, page chrome, notes,
Format Painter, tracked formatting revisions, accept or reject decisions,
strict/transitional DOCX import, exact export, and reopen share the model. A
canonical run-property order plus collision-safe nested export markers retain
both hidden text and original character-style identity. Malformed, duplicated,
misplaced, namespace-spoofed, child- or text-bearing, extra-attribute, unknown,
and conflicting input fails closed. Bounded CSS and PDF projections remain
paint-only so eligible paragraphs stay on Worker/WASM layout. Focused Rstest,
desktop/compact Playwright, and the local-only pinned A3S Test 1.0.0 suite cover
the public Playground template, conflict-safe authoring, Undo, accessibility,
and empty browser diagnostics without adding A3S Test to Actions or Pages.
The forty-fourth milestone adds Writer's native character-border model. One
typed `w:bdr` value preserves all 25 visible WordprocessingML line styles plus
explicit `nil` and `none`, direct or theme colors with tint and shade, widths
from 2 through 96 eighth-points, text spacing from 0 through 31 points, and
explicit shadow and frame flags. The Home Font group exposes a direct toggle;
the shared advanced-font dialog keeps mixed selections untouched and separates
Follow style, explicit no-border, and complete line-style, color, width,
spacing, shadow, and frame authoring. Applying the draft with other character
properties, Format Painter, and direct toggling remains one intent and one Undo
record. Document defaults, paragraph and character styles, body text, page
chrome, notes, tracked formatting revisions, strict/transitional DOCX import,
exact export, and reopen share the bounded model. Malformed, duplicated,
misplaced, namespace-spoofed, child- or text-bearing, extra-attribute,
art-style, out-of-range, and unresolved-theme properties fail closed. Visible
borders use browser-authoritative line measurement because CSS border and
padding change inline geometry; explicit resets remain eligible for
Worker/WASM and CSS/PDF line styles stay bounded visual approximations. No
shortcut is invented. Focused Rstest, desktop/compact Playwright, and the
local-only pinned A3S Test 1.0.0 suite cover the public Playground workflow,
single-step Undo, accessibility, and empty browser diagnostics without adding
A3S Test to Actions or Pages.
The forty-fifth milestone adds bounded Writer document comparison and
reviewed-copy combine. Review imports DOCX, HTML, or TXT against the current
controlled document, deterministically aligns same-layout paragraphs and
headings, and generates existing insertion, deletion, character-formatting,
and paragraph-formatting revisions in one transaction and one Undo record.
Combine accepts only a reviewed copy whose reject-all snapshot exactly
reproduces the current baseline, then retains current paragraph and section
identities. Changed complex structures, layout differences, unresolved current
revisions, malformed review state, empty structural changes, and bounded matrix
or content limits fail before mutation. Native `w:ins`, `w:del`,
`w:rPrChange`, and `w:pPrChange` survive export and reopen; browser-owned
block-container semantics remain an explicit native paragraph-mark boundary.
The public Playground template, bilingual documentation, focused Rstest,
responsive Playwright, and pinned local-only A3S Test suite form the release
evidence without adding A3S Test to Actions or Pages.

The forty-sixth milestone completes the first automatic Spreadsheet calculated
column workflow. When a table contains one consistent current-row formula such
as `=[@Units]*[@[Unit price]]`, the controlled model records it as a calculated
column and fills only empty cells in newly inserted body rows. Values, formulas,
styles, and manual exceptions are never overwritten; a conflicting column
loses its automatic rule instead of guessing. Dense Fortune matrices and
sparse `celldata` both remain supported. The rule is validated at the
collaboration boundary, survives Yjs field-level convergence, and round-trips
through native XLSX `<calculatedColumnFormula>` elements with the editable
leading `=` restored. Focused model, reconciliation, collaboration, XLSX, and
Playground tests plus the local-only A3S Test gate form the release evidence;
slicers and external/query tables remain open.

The forty-seventh milestone completes common Spreadsheet totals-row authoring.
The creation dialog can start a table with totals enabled, and the contextual
Table Design menu enables or disables the row and edits a per-column label,
`sum`, `average`, `count`, `countNums`, `max`, `min`, `stdDev`, `stdDevP`,
`var`, `varP`, or bounded custom formula. Native functions generate
filtered-row-aware `SUBTOTAL` formulas and the first eligible column defaults
to `Total`. Structural edits move the row, direct edits reconcile metadata,
renames rewrite only generated formulas, and occupied, merged, protected,
overlapping, AutoFilter, out-of-bounds, external, unsafe, and over-budget
requests fail before mutation. Dense and sparse worksheets, Yjs collaboration,
XLSX totals metadata, Rust/WASM calculation, and the JavaScript fallback are
covered by focused tests, responsive Playwright, and the local-only A3S Test
gate without adding A3S Test to Actions or Pages.

The forty-eighth milestone adds offline Spreadsheet multi-key Custom Sort.
One dedicated `spreadsheetSort` editor extension captures an immutable range,
detects an editable header, and applies up to 64 unique ascending or descending
value keys with stable numeric and natural text ordering and blanks last.
Complete cell rows move in one native batch and one Undo record; relative
formula rows translate while absolute references remain fixed. Read-only,
locked, merged, pivot, malformed, out-of-bounds, and million-cell-over-budget
requests fail before mutation. Semantic tables, worksheet AutoFilter, and
coordinate-owned hyperlink, imported-formula, or border sidecars also fail
closed until one transaction can reconcile their row ownership. Focused
Rstest, desktop and compact Playwright, and the local-only A3S Test gate cover
the dialog, priority changes, formula movement, accessibility, diagnostics,
and one-step Undo. Table-aware sorting is completed by the fifty-ninth
milestone below; left-to-right direction is completed by the fifty-second.

The forty-ninth milestone aligns partial Spreadsheet sorting with the
Traditional Office Sort Warning flow. AutoFilter and Sort now share one
dense/sparse current-region
model. A partial column or single-cell selection that touches adjacent data
opens an accessible warning which defaults to expanding the selection and also
offers the exact rectangle when it is independently safe. The same planner
serves Custom Sort and quick ascending/descending commands; protection, merge,
pivot, table, AutoFilter, coordinate sidecar, row-count, and area checks can
disable either candidate without weakening the other. Single cells cannot be
sorted as exact ranges.

The selected range, expanded destination, active column, worksheet, and intent
are frozen together. Application is authorized only while the controlled live
selection still matches the original selection, so a host remount or selection
change invalidates the request before mutation. Header inference now accepts
only real non-formula text values and no longer treats formatted numbers or
formula results as headers. Focused Rstest, desktop and compact Playwright, and
the local-only A3S Test gate cover exact and expanded choices, quick and custom
paths, focus, accessibility, formula movement, one-step Undo, and browser
diagnostics. Table-aware sorting is completed by the fifty-ninth milestone
below; left-to-right direction is completed by the fifty-second.

The fiftieth milestone adds offline custom-list order to every Spreadsheet sort
key. Seven immutable Chinese and English month/weekday sequences are always
available. The same dialog accepts newline or comma input for two through 256
unique entries, limits each entry and total text, and retains at most 32 user
lists only while that editor instance remains mounted. Controlled workbook
content stays untouched. Explicit typed cross-editor persistence is completed
by the fifty-fourth milestone below rather than introduced as hidden state.

Each accepted request carries a normalized copy of its sequence, so command
execution does not depend on mutable UI state. Matching trims and normalizes
width/case, listed values sort first, unmatched values use natural ascending
order, blanks remain last, and later keys break equal ranks stably. Focused
model, dialog, and Hook Rstest, desktop/compact Playwright, and the pinned
local-only A3S Test gate cover invalid input, mounted-session reuse,
formula-safe row movement, one-step Undo, accessibility, and empty browser
diagnostics.

The fifty-first milestone adds offline appearance order to every Spreadsheet
sort key. The Custom Sort dialog can place one effective cell color, effective
font color, or conditional-format icon at the top or bottom. Repeated levels
express a complete visual priority without introducing a separate mutable
preference model, and later value, custom-list, or appearance keys continue to
break ties stably.

One bounded appearance model combines direct native styles with calculated
conditional-format output for the frozen range. Conditional presentation wins
over direct presentation, solid/no-fill and automatic-font identities remain
explicit, and native pattern or gradient fills are never flattened into a
misleading solid-color key. The command reconstructs the effective snapshot
from controlled sheet state and live range cells immediately before execution,
validates its shape and target identities, then reuses the existing complete-row
batch, relative-formula translation, and one-step Undo boundary. Focused model,
dialog, command, and Hook Rstest, desktop/compact Playwright, and the pinned
local-only A3S Test gate cover four-level same-column color/font/icon priority, responsive
containment, formula movement, Undo, focus, accessibility, and empty browser
diagnostics. Durable host settings arrive in the fifty-fourth milestone and
structural table/AutoFilter reconciliation in the fifty-ninth; left-to-right
direction is completed below.

The fifty-second milestone adds offline Traditional Office-compatible left-to-right
Spreadsheet sorting. A nested Sort Options dialog switches the existing Custom
Sort surface between column keys that move complete rows and row keys that move
complete columns. Horizontal mode starts from the active row, disables header
retention, retains up to 64 stable value, custom-list, color, font-color, or
conditional-icon keys, and maps appearance placement to left or right.

One direction-neutral matrix engine owns both orientations. It validates every
absolute key and the complete appearance snapshot before mutation, translates
relative formula references along the moved row or column axis, preserves
absolute references, and rejects an out-of-bounds translation atomically. One
accepted request still crosses Fortune as one native range write and one Undo
record. Existing million-cell, protection, merge, pivot, table, AutoFilter,
hyperlink, formula-metadata, and border-sidecar guards remain authoritative.
Focused model, dialog, command, and Hook Rstest, desktop/compact Playwright, and
the pinned local-only A3S Test gate cover multi-row priority, complete-column
movement, formula translation, Undo, focus, responsive containment,
accessibility, and empty browser diagnostics.

The fifty-third milestone completes the common offline text comparison controls
in Spreadsheet Sort Options. Value keys now choose Simplified Chinese pinyin or
stroke order and can distinguish case; the case-sensitive ascending contract
places lowercase before uppercase. Text containing digits uses lexical
character order, while real numeric cells continue to compare numerically and
blank cells remain last.

One `spreadsheet-sort-collation` boundary compiles the requested local
`Intl.Collator` once per accepted operation and is reused by the same stable
matrix comparator in both orientations. It probes observable pinyin or stroke
ordering instead of trusting runtime-specific resolved labels. Invalid settings
and runtimes without the requested ordering fail before mutation instead of
silently falling back.
The dialog request carries normalized options, while older command callers that
omit them retain pinyin, case-insensitive defaults. Focused model, dialog,
command, and focus Rstest, desktop/compact repeated Playwright, and the pinned
local-only A3S Test gate cover pinyin/stroke differences, lowercase ties,
numeric text, formula translation, Undo, responsive containment,
accessibility, and empty browser diagnostics. Deep testing also closes the
controlled-remount race for a grid overlay focused directly by a host or test.

The fifty-fourth milestone makes authored Spreadsheet custom sort lists
reusable across workbook changes and editor remounts through an explicit typed
host boundary. `SpreadsheetEditor.sortCustomListStore` synchronously loads and
saves the same bounded canonical sequences used by the dialog; omitting the
property retains the existing mounted-session behavior and controlled workbook
content never owns a preference.

The provided `LocalStorageSpreadsheetSortCustomListStore` receives a Storage
object instead of selecting a backend by name. It uses a versioned key,
normalizes and deduplicates at most 32 lists, ignores corrupt or unsupported
payloads, and downgrades a failed write to a visible session-only list rather
than discarding user input. Saved lists are grouped separately in the order
selector and keep their identity when a request already carries the same
sequence. Focused Rstest, desktop/compact Playwright, and the pinned local-only
A3S Test gate cover storage round trips, invalid payloads, write fallback,
full-page reload reuse, formula-safe sorting, one-step Undo, accessibility, and
empty browser diagnostics. The same bounded preference surface now exposes a
responsive, keyboard-accessible manager: built-in sequences remain read-only;
user sequences are created, edited, deleted, and reordered in one staged
update; duplicate or over-budget drafts fail locally; and unchanged
confirmation does not cross the store boundary. Active sort keys adopt edited
sequences and fall back to ascending value order after deletion. A rejected
write keeps the complete changed set in the mounted session with visible
feedback.

The fifty-fifth milestone adds common worksheet AutoFilter conditions through
an A3S-owned accessible dialog reached from each active filter header. Users can
author equal/not-equal, contains/does-not-contain, begins/ends-with, numeric
comparison and between/not-between ranges, blanks, and nonblanks. Each
column owns its predicted hidden rows; active columns combine, while manual row
visibility remains a separate provenance record across replacement, clearing,
filter removal, and one-step Undo. Dense `data` and sparse `celldata` share the
same bounded evaluator.

One typed `spreadsheetAutoFilter` extension reauthenticates the editable active
sheet, exact live filter range, and target column immediately before a single
controlled update. The condition dialog retains its frozen selection through a
Fortune remount, and stale requests fail without mutation. Worksheet-level
native OOXML `filterColumn` criteria now import and export through the same
closed filter union already used by tables. Unsafe vendor sort entries are no
longer exposed from the filter menu; the A3S-owned formula-safe sort commands
remain authoritative. Focused Rstest, desktop/compact Playwright, and the
pinned local-only A3S Test gate cover multi-column filtering, manual hidden-row
overlap, XLSX reopen, Undo, selection, accessibility, and clean diagnostics.
The subsequent criteria milestones and current architecture contract complete
same-column AND/OR, wildcard, Top/Bottom, and dynamic evaluation; the
fifty-ninth milestone below completes table/AutoFilter-owned Custom Sort.

The fifty-sixth milestone adds exactly two same-column Custom AutoFilter
conditions joined by AND or OR. The dialog exposes a deliberate second-condition
step, restores imported or controlled pairs, validates each numeric operand,
and includes the Traditional Office negative prefix and suffix operators. The
closed model stores a fixed tuple of nonrecursive custom conditions, so command
reauthorization and collaboration validation cannot admit nested or unbounded
predicate trees. Dense and sparse worksheet evaluation composes the two
matchers without changing multi-column or manual-hide ownership.

Native worksheet and table OOXML now read and write general two-item
`<customFilters and="1|0">` groups while retaining canonical between and
not-between shortcuts. Negative prefix and suffix conditions use native
`notEqual` wildcard forms and preserve escaped literal wildcard characters.
Focused model, command, dialog, collaboration, worksheet/table XLSX, responsive
Playwright, and local-only A3S Test coverage protect the bounded contract.
The subsequent criteria milestones and current architecture contract complete
wildcard, Top/Bottom, and dynamic evaluation; the fifty-ninth milestone below
completes table/AutoFilter-owned Custom Sort.

The fifty-seventh milestone completes local Top/Bottom AutoFilter authoring and
evaluation for numeric columns. The owned **Top 10 Items** menu action opens the
same typed dialog with a bounded ten-item default and supports top or bottom
1–500 items and 1–100 percent. Validation accepts integers only. Ranking ignores
blank, text, Boolean, and nonfinite values, rounds percentage counts upward,
and compares every numeric cell with the selected boundary so ties remain
visible rather than being cut at an arbitrary row.

The dense and sparse worksheet paths share the same threshold evaluator and
retain independent multi-column and manual-hide ownership. Imported or authored
worksheet and table criteria continue to use native `<top10 top="..."
percent="..." val="..."/>` OOXML and are recomputed when a workbook reopens.
Focused model, command, dialog, menu, worksheet/table XLSX, responsive
Playwright, and local-only A3S Test coverage protect application, Undo,
accessibility, and clean diagnostics. Later criteria work completes wildcard
and dynamic evaluation, while the fifty-ninth milestone completes
table/AutoFilter-owned Custom Sort. Large-rank Worker/WASM offload remains
open.

The fifty-eighth milestone completes local wildcard AutoFilter authoring
and evaluation. Two explicit custom-condition variants preserve arbitrary
expressions while existing equality and other text criteria retain their
literal operand semantics. `*` matches zero or more normalized Unicode
characters, `?` matches one, and `~` escapes `*`, `?`, or `~`. The dialog
documents the grammar, permits the variants in the fixed two-condition tuple,
and applies the shared 32,767-character filter-text ceiling before the command
boundary.

The evaluator compiles each expression once into a controlled token sequence
instead of executing or interpolating user input. Ordered literal segments and
a bounded bitset path prevent long late-mismatch expressions from repeatedly
rescanning the same text. Dense and sparse worksheets, command
reauthorization, and collaboration validation share the closed types.
Native worksheet and table `<customFilter>` import retains arbitrary positive
or negative patterns, while canonical leading/trailing-star forms continue to
normalize to contains, begins-with, or ends-with criteria. Focused Rstest,
desktop/compact Playwright, and local-only A3S Test coverage protect Unicode,
escaping, compound authoring, XLSX reopen, Undo, accessibility, and clean
diagnostics. Large aggregate/rank Worker/WASM offload remains open; the
fifty-ninth milestone below completes table/AutoFilter-owned Custom Sort and
the sixtieth completes source-level 1904 date-system retention.

The fifty-ninth milestone integrates Spreadsheet Custom Sort with native table
and worksheet AutoFilter ownership. Selecting any cell or rectangle inside one
structure opens the Sort Warning with only its exact sortable owner enabled.
Table totals are excluded, structural headers remain fixed, and Sort Options
locks movement to top-to-bottom rows. The request carries the owner kind,
header state, and table ID when applicable; the command derives that fingerprint
again immediately before mutation, so changed ownership fails without editing
cells.

Filter-active owner sorts publish one controlled workbook change and one Undo
record. Typed worksheet and table criteria are evaluated again against the
sorted dense or sparse cells. Native Fortune value/color filters that expose
only opaque `caljs` state instead move their `rowhidden` ownership through the
stable source-row permutation, while independent manual hiding is retained.
The command recovers hidden rows omitted by the native range reader from the
controlled snapshot, limits the sort and filter-rescan budgets to one million
cells each, and leaves unrelated filter regions on the native range-write path.
Focused model, command, Hook, and dialog Rstest plus a local-only A3S Test gate
and desktop/compact Playwright cover table formulas, native AutoFilter
visibility, focus, responsive containment, one-step Undo, accessibility, and
clean browser diagnostics.

The sixtieth Spreadsheet milestone makes the workbook date system an explicit
controlled-model fact instead of an import-time assumption. Missing state
means the standard 1900 system; imported 1904 workbooks retain an explicit
`dateSystem: '1904'`, raw date-typed serials including zero, and native
`workbookPr date1904` on export. Date-typed formula caches remain numeric, so a
save and reopen cannot introduce local-time or historical-time-zone drift.

Dynamic relative, month, and quarter AutoFilters, filter-menu profiling,
filter reconciliation after edits or owned-range sorts, and `Ctrl+;`
current-date authoring all consume that one workbook epoch. The 1900 path
continues to reject the fictional leap-day serial 60, while 1904 serial zero
maps to 1904-01-01. Collaboration validates and merges date-system changes
independently from calculation settings. Focused model, command,
collaboration, and real-package XLSX reopen tests plus a local A3S Test browser
gate cover exact serial retention, imported filter visibility, controlled
reopen, accessibility, and clean diagnostics.

The sixty-first Writer milestone makes ordered-list numbering changes a
first-class review intent. Changing a common decimal, lower/upper letter, or
lower/upper Roman list style, or changing its starting value, stores one
canonical snapshot on the ordered-list node. The review panel presents one
Numbering card for the complete range. Accept keeps the current numbering;
reject restores the original style, start value, and retained Office numbering
identity without touching list text. The initial change, final decision, and
Undo each remain one atomic transaction.

Strict and transitional WordprocessingML `w:numberingChange` records import
only when the paragraph numbering properties, revision metadata, original
single-level definition, and contiguous per-item sequence are unambiguous.
Export writes one native record per list item with sequential original values
and enables native revision settings without leaking browser transport markers.
The importer and diagnostics fail closed for malformed, duplicated,
conflicting, namespace-spoofed, unsupported-format, or multi-level definitions.
The model caps one snapshot at 64 KiB and one package at 65,536 numbering
records; ordinary typing avoids numbering-tree scans because tracking runs only
for structural list transactions.

Browser Yjs and native Yrs carry the same live node metadata and immutable
`changeKind: "numbering"` decision records. Native persistence reopens the
browser fixture, authenticated text suggestions preserve the numbering intent,
and attempted metadata removal is rejected before state changes. Focused
Rstest, DOCX export/import/reopen, responsive browser coverage, and the pinned
local A3S Test numbering suite cover review, rejection, Undo, accessibility,
and clean diagnostics. Complex and multi-level numbering changes, move ranges,
and section/table/row/cell property revisions remain explicit follow-up work.

The sixty-second Spreadsheet milestone completes the first editable formula
conditional-format path. Home → Conditional Formatting authors bounded local
rules with relative/absolute references, finite worksheet-qualified references,
independent text/fill colors, ordered precedence, and Stop-if-true behavior;
imported XLSX `expression` rules use the same editor instead of becoming a
file-only record. A shared synchronous evaluator reads only cached values,
limits each formula to 255 Unicode characters and 1,024 referenced cells, scans
bounded blank ranges without densifying the workbook, and fails closed for
external, whole-row/column, missing-sheet, or uncached-formula references.
Native differential styles, `sqref`, priorities, and formulas survive export
and reopen. The public **新建 → 公式条件格式** template and focused Rstest,
desktop/compact Playwright, accessibility, and local A3S Test gates make this
boundary inspectable without a remote service. Formula-function breadth,
large-range Worker evaluation, and advanced visual rule families remain
explicit follow-up work.

The sixty-third Spreadsheet milestone adds dependent local dropdown lists to
the existing Data Validation dialog. A list source may use a bounded
`=INDIRECT(...)` expression composed of quoted text, single-cell references,
and concatenation. Relative drivers are re-evaluated from each selected
range's top-left anchor, so a Region column can select a workbook-local named
range or one-dimensional area for the next column. Empty drivers show an empty
list; external books, whole-row/column references, missing or hidden sheets,
uncached formulas, and two-dimensional results fail closed. The authored rule
remains compact in `WorkSpreadsheetContent`, while the Fortune projection is
limited to 1,024 source cells and 10,000 materialized cells. Native XLSX list
formulas and names survive import, export, and reopen, and the public Data
Validation template demonstrates the Region → Regional owner workflow. The
dialog's source field, formula note, and Sigma affordance keep the dependency
visible without introducing a second settings model or a remote service.

The sixty-fourth Writer milestone adds a bounded native picture-transform slice.
The contextual Picture ribbon and responsive Picture Properties dialog expose
90-degree rotation plus horizontal and vertical reflection as one typed
per-image state. Each command is one controlled TipTap update and one Undo
record, and the same CSS projection is consumed by editing, preview, and PDF
capture. DOCX `a:xfrm` `rot`, `flipH`, and `flipV` values import, export, and
reopen for the supported quarter-turn/reflection subset; arbitrary-angle,
malformed, and unsupported DrawingML values receive explicit diagnostics and
normalize safely. Focused model, UI, DOCX round-trip, responsive browser, and
local A3S Test coverage document the boundary without adding a remote service.

## Current Presentation Milestone

The second Presentation-animation milestone extends the bounded sequence with
composable entrance and exit effects. One slide-owned object may have one
editable animation from each class: appear, fade, fly in, or zoom for entrance;
disappear, fade out, fly out, or zoom out for exit. Array order remains playback
order, and each item starts on click, with the previous item, or after the
previous item. Duration is bounded from 100 milliseconds through 60 seconds,
delay from zero through 60 seconds, and fly direction to left, right, up, or
down. A slide accepts at most 256 animation records. Two effects for one object
may share a cue only when their time intervals do not overlap.

The Animation ribbon tab follows the Traditional Office object workflow: it selects Entrance or
Exit, offers the matching effect and fly direction, edits the shared trigger and
timing, changes sequence order, and previews the current slide. Invalid trigger
or timing changes are disabled or rejected before mutation. Every accepted
command emits one controlled presentation update and one Undo record.
Collaboration validates the complete sequence and stable object targets;
deleting an object removes both classes, while object and slide copies create
fresh animation and object identities without breaking their mapping.
Slideshow and presenter view consume pending click cues before navigating.
Automatic cues start with the slide, sequential same-object effects compose in
one browser animation list, and reduced-motion environments jump to each cue's
final visibility while retaining the same sequence semantics.

The supported subset exports as native PresentationML timing trees with
`presetClass="entr"` or `presetClass="exit"` and matching in/out effect
transitions. Import maps both classes back into the editable Work model with
effect, order, trigger, duration, delay, direction, and object identity intact
across a second export and reopen. Malformed timing, namespace spoofing,
unavailable or ambiguous targets, duplicate per-class object effects,
overlapping same-object cues, unsupported effects, and inconsistent transition
metadata fail closed with compatibility diagnostics. The remaining alignment
work is explicit: emphasis effects, motion paths, trigger-on-object, the broader
native animation catalog, audio and video, rehearsal and recording, and full
preservation of unsupported timing trees.

The public Playground exposes the **进入与退出动画** example in the New template
grid, with all eight effects and all three trigger modes in a two-cue
entrance-then-exit story. Focused Rstest covers
the model, commands, ribbon, slideshow, collaboration, clipboard identity, and
native PPTX export/import/reopen. A responsive Playwright workflow and the
pinned local-only A3S Test suite complete the release evidence without
installing or invoking A3S Test in Actions or Pages.

## Current PDF Milestone

The first PDF-workbench milestone adds one page-organization surface for seven
operations: insert a blank page, delete, rotate, reorder, extract, merge another
PDF, and split. The operation engine and `pdf-lib` load lazily in a dedicated
Web Worker instead of adding page-rewrite work to the rendering thread. Insert,
delete, rotate, reorder, and merge replace the current PDF with one complete
`Blob` and add exactly one page-history record. Extract and split leave the
source unchanged and publish typed `PdfPageOrganizationExport` files through
`onPageExport`, or download them when the host omits that callback. Toolbar
Undo and Redo consume native PDF annotation/form history first, then the page
history.

The bounded contract accepts a primary PDF up to 256 MiB, a merge source up to
128 MiB, and 1 through 4,096 pages in each result. Malformed and encrypted
inputs fail closed. Any mutation rejects a signed source; delete, reorder, and
merge additionally reject forms, outlines, or tagged structures whose page
references cannot be rewritten safely. Extract and split may copy pages from
those structures, but report that document-level outlines, forms, tags,
attachments, scripts, and signatures are not copied. Page organization is
hidden while a collaboration session or evidence overlay is active because
both modes bind review state to an immutable source identity.

Focused Rstest covers the planner, Worker protocol, controller, dialog, toolbar,
and framework adapters. Desktop Playwright performs real drag-and-drop reorder,
downloads and independently parses exports, saves the mutation, and verifies
PDFium reopen; compact coverage verifies overflow discovery and exact focus
restoration. The public Playground exposes **组织 PDF 页面**, both documentation
languages describe the same API and limits, and the pinned A3S Test 1.0.0 ACL
remains a local-only release gate. GitHub Actions and Pages do not install or
invoke A3S Test.

## Product Rules

- Do not add a visible command that has no executable behavior.
- Keep command labels, locations, shortcuts, and availability in one catalog
  per editor when they represent stable product behavior.
- One user intent produces one controlled update and one undo record.
- Keyboard handling stays scoped to the active editor and never captures host
  form inputs or modal text entry.
- Desktop and compact Web layouts preserve high-priority editing commands
  before paging or hiding secondary groups.
- Temporary surfaces restore the exact invocation or editing focus on close.
- Native-file compatibility remains an acceptance gate; familiar chrome does
  not compensate for data loss or incorrect rendering.

## Release Evidence

Every aligned slice requires focused command/controller tests, component tests,
TypeScript and formatting checks, a production build, and desktop and compact
Web browser coverage. A deterministic A3S Test ACL is a local-only release gate;
GitHub Actions neither installs nor runs A3S Test. A local browser-driver outage
may block ACL execution, but the manifest must still validate and equivalent
Playwright coverage must pass before the change is submitted.
