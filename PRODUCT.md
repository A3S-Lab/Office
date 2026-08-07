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
TypeScript and formatting checks, a production build, desktop and compact Web
browser coverage, and a deterministic A3S Test ACL. A browser-driver outage may
block local ACL execution, but the manifest must still validate and equivalent
browser coverage must pass before the change is submitted.
