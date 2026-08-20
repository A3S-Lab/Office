# A3S Office Product Direction

A3S Office is a browser-first editing component for teams that need familiar
office workflows inside their own products. It should let an experienced WPS
user find routine commands, predict keyboard behavior, and move between files
without relearning the editor shell.

## WPS Alignment Contract

Alignment targets information architecture and operating habits before visual
similarity. A3S Office uses its own product identity and implementation while
matching the command locations, tab order, shortcut expectations, disclosure
patterns, focus behavior, and responsive priorities that make WPS familiar.

The delivery order is:

1. Web before native shells.
2. Writer before Spreadsheet, Presentation, and PDF.
3. Daily editing paths before long-tail command coverage.

Writer establishes the shared interaction baseline: a quick-access row,
priority-aware and collapsible ribbon, temporary tab expansion, stable command
metadata, WPS-oriented shortcuts, contextual tools, and a useful status bar.
Each later editor adopts that baseline only where it matches the editor's own
file model and workflows.

## Current Spreadsheet Milestone

Spreadsheet is the next vertical track after the Writer shell baseline. Its
first milestone keeps the WPS tab sequence, moves Undo and Redo into quick
access, places Conditional Formatting under Home and Styles, exposes real sort
commands under Data, and makes workbook recalculation discoverable and
executable through F9. The shared ribbon must remain adaptive, persistently
collapsible, and temporarily expandable without moving the workbook surface.
The second milestone makes the existing permission-resilient Paste, Cut, and
Copy behavior discoverable in the Home clipboard group and routes both ribbon
clicks and WPS shortcuts through the same typed command port.
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
command freezes rows above and columns left of the selection, matching the WPS
workflow, while dedicated presets freeze the top row or first column and an
active worksheet exposes Unfreeze Panes. One controlled update retains the
selection, the menu supports Arrow, Home, End, Enter, and Escape with exact
grid-focus restoration, and the boundary survives XLSX export and reopen.
The sixth milestone makes the existing row and column structure commands
discoverable under Home and Cells through one WPS-familiar Rows and Columns
menu. Users can insert selected rows above or below, insert selected columns to
the left or right, and delete selected rows or columns. Each item derives its
availability from the live selection, uses the shared typed command port, and
returns focus to the grid for uninterrupted keyboard editing.
The seventh milestone moves cell merging into the WPS-familiar Home and
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
The tenth milestone completes the first missing WPS font-emphasis path by
placing Strikethrough in Home and Font. The ribbon button and `Cmd/Ctrl+5`
share the existing typed cell-format command, preserve grid focus, produce one
undoable controlled update, and expose the active native `cl` style without a
second formatting model.
The eleventh milestone adds a WPS-familiar cell-border split control to Home
and Font. It supports top, bottom, left, right, no, all, outside, inside,
horizontal, vertical, and diagonal borders with ten native line styles and an
exact color input. One immutable A3S-owned command compacts overlapping native
range records, preserves malformed vendor records, emits one controlled update,
and returns focus to the grid. Diagonal writes stay bounded to 4,096 cells and
undo through the shared workbook history.
The twelfth milestone adds Fill under Home and Editing. Fill Down, Right, Up,
and Left copy the appropriate selection edge through Fortune's native formula
and style semantics; `Cmd/Ctrl+D` and `Cmd/Ctrl+R` provide the WPS fast paths.
One editable, unmerged, unprotected, non-pivot range may fill at most 50,000
target cells. The adapter materializes only missing row arrays inside that
range before calling Fortune and restores the original sparse shape if the
native command fails. A successful fill remains one undoable operation and
returns focus to the workbook grid.
The thirteenth milestone completes the common WPS number-format path under Home
and Number. General, Number, CNY Currency, Accounting, Percentage, Short Date,
Time, Scientific, Fraction, and Text are explicit native presets instead of
collapsing into Custom. The format picker, direct Currency and Percentage
buttons, and the seven standard `Cmd/Ctrl+Shift` formatting shortcuts share one
typed `ct` command. Each intent updates only the selected cells' format code and
Fortune value type, preserves values and formulas, creates one undo record, and
round-trips its exact XLSX format code. Imported combined date-time and other
unmodeled codes remain visible as disabled Custom state rather than being
silently rewritten.
The fourteenth milestone adds the WPS-familiar Cell Styles gallery under Home
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
The sixteenth milestone adds the complete WPS-style Format Cells workflow under
Home and Number with the editor-scoped `Cmd/Ctrl+1` shortcut. Number, Alignment,
Font, Border, Fill, and Protection share one dialog that preserves untouched
mixed values, captures its exact source range, retains dense or sparse worksheet
representation, and publishes one controlled update and one Undo record. The
general command accepts at most 10,000 cells and diagonal borders retain their
stricter 4,096-cell bound; host inputs, formula or cell editing, and modal
controls keep their native shortcut behavior.
The seventeenth milestone adds the WPS AutoSum split command before Fill under
Home and Editing. Sum is the primary action and `Alt+=` fast path; the menu adds
Average, Count, Maximum, and Minimum with complete Arrow, Home, End, Enter, and
Escape behavior. One blank target infers the nearest contiguous numeric or
formula block above before looking left, while an explicit totals row or column
can emit multiple formulas without replacing label or text axes. The command
preserves target formatting, rejects occupied, merged, protected, read-only,
inactive, or pivot targets, caps one gesture at 1,000 formula cells, keeps sparse
worksheets sparse, and commits one native batch, one controlled update, and one
Undo record.
The eighteenth milestone replaces the standalone Home and Editing Find button
with the WPS-style Find and Select menu. Find retains `Cmd/Ctrl+F`; Go To adds
`Ctrl+G` and `F5` for direct A1 cells, continuous ranges, quoted cross-sheet
references, and worksheet- or workbook-scoped names. Invalid, ambiguous,
hidden, multi-area, unsupported, and out-of-bounds targets fail inside the
shared dialog. Successful navigation activates the target sheet, selects and
scrolls to the focused cell, and restores grid focus without publishing
`onChange`, adding Undo history, or materializing sparse cells. Host inputs,
formula or cell editing, popovers, and modal controls retain their native keys.
The nineteenth milestone adds the WPS-style Paste split command and Paste
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
