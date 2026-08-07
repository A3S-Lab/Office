# Local A3S Test E2E

These deterministic browser scenarios use the separately installed
[`a3s-test`](https://github.com/A3S-Lab/Test) CLI version 0.4.4 or newer and
its A3S Browser adapter.
Evidence is written under `.a3s-test/`, which is ignored by Git.

Build the Playground and start its static preview in one terminal:

```bash
bun run playground:build
bun run playground:preview
```

Validate and run the ACL suites from another terminal:

```bash
bun run test:e2e:check
bun run test:e2e
```

The suites cover focused Word page-color, a desktop Writer quick-access and
adaptive, collapsible-ribbon workflow covering comfortable, priority-compacted,
persistent collapsed, and temporary tab states without document layout
movement, plus executable WPS formatting and review shortcuts scoped to the
document, accessibility, screenshot, console, and page-error evidence;
the complete phone Page Setup flow
with grouped header/footer formatting controls, phone paragraph-spacing and
pagination popovers with committed document semantics and exact invoker-focus
restoration, a phone list-gallery workflow covering bounded ribbon paging,
square bullets, upper-Roman numbering, a custom start value, viewport-safe
touch controls, TipTap selection and editor-focus recovery, and
Escape-to-invoker focus restoration, compact comments-drawer, a phone citation
workflow with
field-level validation, save, insertion, and close-to-invoker focus
restoration, full-text navigation, the complete phone Find/Replace workflow,
page-preview, bounded page-button virtualization for a
real 120-page DOCX with animation-frame-independent Home/End keyboard focus,
bounded heading-outline and full-text-result windows for the same real DOCX,
including sparse current/selection pins, physical spacer geometry, exact
selection of result 120, and same-frame alignment with the final document page,
selected-text AI question-composer regressions, and durable picture insertion
from a project-relative fixture after the browser-managed file input is reset,
plus a phone-width Picture Properties workflow covering centimeter size,
aspect-ratio locking, wrapping, alignment, text distance, alternative text,
retained image selection, and exact invoker restoration after apply and cancel,
and a
phone tracked-changes workflow with in-pane recording controls, truthful empty
states, exact decision focus, and close-to-invoker restoration, plus a real
120-comment native OOXML DOCX workflow covering bounded card and connector
mounting, Home/End reachability, document-selection and physical-page
synchronization, the 120-to-119 deletion transition, adjacent focus continuity,
and non-blocking status feedback, plus a real
120-revision DOCX workflow covering bounded mounted rows, physical spacer
geometry, Home/End reachability, keyboard acceptance, the 120-to-119 count
transition, and same-action focus continuity, plus a phone
table workflow covering touch-sized dimension controls, exact 3 × 3 insertion,
row extension, compact-ribbon deletion, and focused editing recovery, plus a
desktop table-border workflow proving that the reusable border pen does not
flatten existing edges, outside borders retain independent cell-edge data in
edit and preview, and browser diagnostics stay clean, plus a real theme-backed
DOCX table workflow proving tint/shade resolution for cell fill and four
independent border edges across edit and preview with screenshot,
accessibility, console, and page-error evidence, plus a real inherited-table-
style DOCX workflow proving based-on resolution, first-row fill, white bold
header text and double border, horizontal banding, last-row presentation, and
conditional paragraph alignment, spacing, pagination, and indent semantics
across edit and preview. The same workflow verifies truthful 62.5% centered
geometry in the unified Table Properties dialog, compact 390 px touch controls,
all four Table, Row, Column, and Cell tabs, and one atomic change covering 72%
left placement with a 0.5 cm indent, a 1.2 cm exact row, disabled row splitting
and repeated headings, a 4 cm column, centered cell content, and a 0.4 cm left
cell margin. It also proves cancellation and Escape safety, invoker-focus
restoration, exact retention of untouched imported 8 px and 16 px margin
edges, and matching preview geometry, with screenshot, accessibility, console,
and page-error evidence. A real multi-page
DOCX table workflow proves that a splittable
row continues at paragraph boundaries across three physical pages, repeated
heading rows remain visible
on both continuation pages, and the final row content and following paragraph
remain keyboard reachable, plus a
phone caption and cross-reference workflow covering dialog focus, accessible
numbering, keyboard deletion, and the live missing-reference transition;
the phone Spreadsheet
worksheet-rename workflow plus viewport-safe, touch-sized Find controls,
exact-result navigation, and grid-focus restoration, a safe-area-aware bottom
context menu with 44 px actions, bounded scrolling, screenshot and
accessibility evidence, together with modal
workbook task panes that isolate the ribbon, grid, and worksheet footer,
contain forward and reverse focus, and restore the exact ribbon invoker;
Presentation focus
recovery across object cut and paste, presenter-view keyboard navigation and
phone layout, plus a phone chart-inspector workflow covering responsive modal
semantics, background isolation, forward and reverse focus containment, dirty
draft cancellation, screenshot and accessibility evidence, and exact
selected-chart focus restoration, plus phone comment review covering full-
editor modal semantics, touch-sized review controls, background isolation,
forward and reverse focus containment, dirty-comment cancellation, and exact
New Comment and View Comments invoker restoration; PDF import
from a project-relative fixture plus focus-synchronized thumbnail keyboard
navigation, plus a phone page-drawer workflow proving toolbar ownership,
selection, modal close behavior, and an unobstructed document canvas; the
phone Markdown source, visual, and split-mode workflow with
unambiguous accessible controls, a source-first full-workspace phone switch,
synchronized preview content, and touch-sized actions; and the public
Playground-to-documentation-center navigation contract, including Simplified
Chinese as the default language, language and release-version switching, the
documentation tree, and the same-deployment return to the Playground homepage.
The responsive Office-shell workflow also crosses from a persistent desktop
sidebar into a 390 px workspace, proves the phone drawer stays closed without
moving search focus, and restores the prior persistent sidebar on return.
Each suite owns only its browser surface. Keep the preview process under the
terminal that started it and stop that process separately when testing is
complete.
