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

The suites cover focused Word page-color, compact comments-drawer, full-text
navigation, page-preview, and selected-text AI question-composer regressions;
the phone Spreadsheet
worksheet-rename workflow; Presentation focus recovery across object cut and
paste plus presenter-view keyboard navigation and phone layout; PDF import
from a project-relative fixture plus focus-synchronized thumbnail keyboard
navigation; and the public Playground-to-documentation-center navigation
contract, including the documentation tree and the same-deployment return to
the Playground homepage.
Each suite owns only its browser surface. Keep the preview process under the
terminal that started it and stop that process separately when testing is
complete.
