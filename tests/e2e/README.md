# Local A3S Test E2E

These deterministic browser scenarios use the separately installed
[`a3s-test`](https://github.com/A3S-Lab/Test) CLI and its A3S Browser adapter.
Evidence is written under `.a3s-test/`, which is ignored by Git.

Build the Playground and start its static preview in one terminal:

```bash
bun run playground:build
bun run playground:preview
```

Validate and run the ACL suite from another terminal:

```bash
bun run test:e2e:check
bun run test:e2e
```

The suite owns only its browser surface. Keep the preview process under the
terminal that started it and stop that process separately when testing is
complete.
