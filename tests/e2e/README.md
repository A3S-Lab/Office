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

The suites cover focused Word page-color, the complete phone Page Setup flow
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
plus phone-width picture alternative-text editing that retains the image
selection and restores the exact ribbon invoker after save and cancel, and a
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
phone caption and cross-reference workflow covering dialog focus, accessible
numbering, keyboard deletion, and the live missing-reference transition;
the phone Spreadsheet
worksheet-rename workflow plus viewport-safe, touch-sized Find controls,
exact-result navigation, and grid-focus restoration, together with modal
workbook task panes that isolate the ribbon, grid, and worksheet footer,
contain forward and reverse focus, and restore the exact ribbon invoker;
Presentation focus
recovery across object cut and paste plus presenter-view keyboard navigation
and phone layout; PDF import
from a project-relative fixture plus focus-synchronized thumbnail keyboard
navigation, plus a phone page-drawer workflow proving toolbar ownership,
selection, modal close behavior, and an unobstructed document canvas; the
phone Markdown source, visual, and split-mode workflow with
unambiguous accessible controls; and the public
Playground-to-documentation-center navigation contract, including Simplified
Chinese as the default language, language and release-version switching, the
documentation tree, and the same-deployment return to the Playground homepage.
Each suite owns only its browser surface. Keep the preview process under the
terminal that started it and stop that process separately when testing is
complete.
