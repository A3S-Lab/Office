# Contributing to A3S Office

Thank you for improving A3S Office.

## Before opening a change

1. Search existing issues and pull requests.
2. Open an issue for behavior changes or new public APIs.
3. Keep changes focused and include tests for modified behavior.
4. Do not include documents, credentials, or customer data in fixtures.

## Local development

Install Node.js 20 or newer, Bun 1.3 or newer, and Rust 1.85 or newer.

```bash
bun install
bun run typecheck
bun run test
bun run build
cargo fmt --all -- --check
cargo check --workspace --all-targets
cargo test --workspace
bun run playground
```

Before submitting a pull request, also run:

```bash
bun run format
bun run lint
bun run playground:build
```

## Source structure

- `src/react.tsx` contains the public React adapters.
- `src/vue.ts` contains the Vue 3 adapters.
- `src/web-component.tsx` contains the Custom Elements adapters.
- `src/core.ts` defines the framework-free public API.
- `src/internal/` contains the editor engine.
- `src/styles/` contains the editor stylesheet entry and internal styles.
- `playground/` is the browser integration example.
- `tests/` contains Core and framework adapter tests.
- `crates/core/` contains the native OOXML engine.
- `crates/cli/` contains the `a3s-office` CLI, standard MCP server, Skill, and
  native integration tests.
- `integrations/a3s-use/` contains the external A3S Use capability manifest.
- `scripts/package-a3s-use-extension.sh` creates an installable A3S Use package.

## Public feature completion

A user-facing editor capability is complete only when all applicable surfaces
ship in the same change:

1. Add a concise, discoverable README entry instead of relying only on a long
   capability table or changelog item.
2. Update both `docs/latest/en/` and `docs/latest/zh/`, and link the capability
   from the documentation home while it is part of the current `main` release
   window.
3. Add a normal Playground template or another user-facing workflow. A hidden
   `?e2e=` fixture does not count as product discovery.
4. Cover the public entry with a component test and the local A3S Test Web
   gate. A3S Test remains outside GitHub Actions and Pages.
5. Verify the entry is visible and operable at desktop and compact viewport
   sizes before publishing Pages.

## Pull requests

Describe the user-facing problem, the chosen behavior, and the validation that
was run. Add a migration note when changing a public content model or Props
type. Keep generated `dist/` and `playground-dist/` files out of commits.

All contributions are licensed under the repository's MIT License.
