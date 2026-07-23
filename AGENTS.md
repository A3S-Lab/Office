# AGENTS.md

You are an expert in JavaScript, Rspack, Rsbuild, Rslib, and library development. You write maintainable, performant, and accessible code.

## Commands

- `bun run build` - Build the library for production
- `bun run dev` - Turn on watch mode, watch for changes and rebuild the library
- `bun run typecheck` - Type-check the public package and editor engine
- `bun run test` - Run tests
- `bun run test:watch` - Run tests in watch mode
- `bun run playground` - Start the local integration playground
- `bun run playground:build` - Build the integration playground

## Docs

- Rslib: https://rslib.rs/llms.txt
- Rsbuild: https://rsbuild.rs/llms.txt
- Rspack: https://rspack.rs/llms.txt
- Rstest: https://rstest.rs/llms.txt

## Tools

### Biome

- Run `bun run lint` to lint your code
- Run `bun run format` to format your code

## Repository rules

- Keep public APIs in `src/core.ts`, `src/react.tsx`, `src/vue.ts`, and
  `src/web-component.tsx`.
- Keep implementation details under `src/internal/`.
- Do not expose an internal module through package exports without a public API
  review.
- Preserve controlled component semantics: hosts own content and persistence.
- Keep React, Vue, Web Component, and Core behavior aligned.
- Add tests for changed public behavior.
- Keep code, comments, and documentation in English. Localized user-interface
  strings may remain in their target language.
