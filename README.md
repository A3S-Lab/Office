<p align="center">
  <img src="./assets/readme/hero.svg" width="100%" alt="A3S Office brings five format-native editors into a host-owned product boundary">
</p>


<p align="center">
  <strong>Language / 语言:</strong>
  <a href="README.md">English</a> ·
  <a href="README.zh-CN.md">中文</a>
</p>

<p align="center">
  <strong>Open-source browser editors and deterministic native automation for real Office files.</strong>
</p>

<p align="center">
  Document · Spreadsheet · Presentation · Markdown · PDF
</p>

<p align="center">
  <a href="https://a3s-lab.github.io/Office/playground/"><strong>Open the Playground</strong></a>
  ·
  <a href="https://a3s-lab.github.io/Office/docs/">Read the documentation</a>
  ·
  <a href="#quick-start">Embed an editor</a>
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI status" src="https://img.shields.io/github/actions/workflow/status/A3S-Lab/Office/ci.yml?branch=main&amp;style=flat-square&amp;label=CI"></a>
  <a href="https://www.npmjs.com/package/@a3s-lab/office"><img alt="npm version" src="https://img.shields.io/npm/v/@a3s-lab/office?style=flat-square&amp;color=1456f0"></a>
  <a href="#project-status"><img alt="Project status: pre-1.0" src="https://img.shields.io/badge/status-pre--1.0-7457c8?style=flat-square"></a>
  <a href="./LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/license-MIT-16845b?style=flat-square"></a>
</p>

<p align="center">
  <a href="#proof-not-promises">Product</a> ·
  <a href="#the-core-design">Design</a> ·
  <a href="#quick-start">Quick start</a> ·
  <a href="#collaboration-without-a-bundled-cloud">Collaboration</a> ·
  <a href="#automation-without-ui-scraping">Automation</a> ·
  <a href="#capabilities-and-boundaries">Boundaries</a> ·
  <a href="./ROADMAP.md">Roadmap</a> ·
  <a href="./CONTRIBUTING.md">Contributing</a>
</p>

---

A3S Office is an open-source Office engine for teams building editing into
their own products. It provides complete browser surfaces for documents,
spreadsheets, presentations, Markdown, and PDFs, plus a separate Rust
automation plane for files and coding agents.

The product boundary is deliberate: A3S Office owns format-aware editing,
layout, import, export, and typed mutations. Your application keeps control of
content, persistence, identity, authorization, collaboration transport, and AI
providers. Core editing and file workflows require no A3S backend or bundled
cloud service.

## Proof, not promises

These are committed visual-regression baselines from the real
[Playground](https://a3s-lab.github.io/Office/playground/), not concept art.

<p align="center">
  <a href="./visual-tests/__snapshots__/linux/desktop-1280/document.png">
    <img src="./visual-tests/__snapshots__/linux/desktop-1280/document.png" alt="A3S Office Document editor with an Office-style ribbon and paginated project brief" width="100%">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png" alt="A3S Office Spreadsheet editor showing a quarterly execution plan">
      </a>
      <br><sub><strong>Spreadsheet</strong> — formulas, tables, formatting, sort, filter, and print workflows</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/presentation.png" alt="A3S Office Presentation editor with slide thumbnails and a slide canvas">
      </a>
      <br><sub><strong>Presentation</strong> — structured slides, objects, animations, and presenter flows</sub>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/markdown.png" alt="A3S Office Markdown editor in synchronized source and preview mode">
      </a>
      <br><sub><strong>Markdown</strong> — GFM source, visual editing, and synchronized preview</sub>
    </td>
    <td width="50%" valign="top">
      <a href="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png">
        <img src="./visual-tests/__snapshots__/linux/desktop-1280/pdf.png" alt="A3S Office PDF editor with search, annotation, page organization, save, and download controls">
      </a>
      <br><sub><strong>PDF</strong> — PDFium rendering, forms, annotations, page organization, and save</sub>
    </td>
  </tr>
</table>

## Why A3S Office exists

Embedding Office work is not one problem. A dependable product needs four
things at the same time:

1. **A complete interaction surface** — ribbons, dialogs, panes, shortcuts,
   responsive layouts, accessibility, and predictable focus.
2. **Format-native behavior** — DOCX, XLSX, PPTX, Markdown, and PDF cannot be
   flattened into one lowest-common-denominator model without losing meaning.
3. **A clear ownership boundary** — the editor should not dictate where files
   live, who users are, how permissions work, or which AI provider is allowed.
4. **Deterministic automation** — agents and backend jobs need typed,
   conflict-aware mutations instead of UI scraping.

A3S Office keeps those concerns separate while exposing them through one
package and one set of bounded contracts.

## Five surfaces, one product boundary

- **Document** (`DOCX`, `HTML`, `TXT`) — structured authoring, live
  pagination, tables, equations, references, review, and PDF output.
- **Spreadsheet** (`XLSX`, `XLS`, `ODS`, `CSV`) — sparse worksheets,
  formulas, tables, formatting, sort/filter, validation, pivots, charts, and
  print workflows.
- **Presentation** (`PPTX`) — slides, typed scene objects, masters/layouts,
  transitions, bounded entrance/exit animations, notes, slideshow, and
  presenter view.
- **Markdown** (`MD`) — GFM source, visual mode, split preview, direct round
  trips, and native automation.
- **PDF** — PDFium rendering, search, forms, annotations, history, save, and
  page organization.

The surface is requested lazily. Large or expensive work is isolated behind
cancellable Workers, Rust WebAssembly, viewport-bounded rendering, and an
explicit PDFium runtime. React, Vue 3, Web Components, and the framework-neutral
Core API use the same controlled content models.

## The core design

<p align="center">
  <a href="./assets/readme/architecture.svg">
    <img src="./assets/readme/architecture.svg" width="100%" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

The two execution planes solve different jobs:

- **Browser editing** provides complete interactive surfaces through React,
  Vue 3, Web Components, and the Core API. Workers, Rust/WASM, and PDFium keep
  layout, calculation, parsing, and rendering bounded.
- **Native automation** provides deterministic file reads, validation,
  mutation, batching, CLI commands, a standard MCP server, and an Office Skill
  without launching desktop Office.

Both planes are controlled by the host:

| A3S Office owns | Your product owns |
| --- | --- |
| Format-native models and commands | Content persistence and version history |
| Import, export, layout, and rendering | Identity, authorization, and policy |
| Editor UI and responsive interaction | Application shell and navigation |
| Typed collaboration and agent ports | Rooms, transport, model providers, and AI lifecycle |

Neither plane requires an A3S backend. Read the
[browser architecture](./docs/latest/en/browser-editor-architecture.md) and
[native engine design](./docs/latest/en/native-office-engine.md) for the exact
boundaries.

## Quick start

### Try the complete product

The fastest first success is the
[live Playground](https://a3s-lab.github.io/Office/playground/). It exposes
normal document templates, recent capabilities, file import, and every editor
without a local install.

To run the same Playground locally, use Node.js 20+, Bun 1.3+, and Rust 1.85+:

```bash
git clone https://github.com/A3S-Lab/Office.git
cd Office
bun install --frozen-lockfile
bun run playground
```

### Embed a controlled React editor

```bash
bun add @a3s-lab/office react react-dom
```

Import the stylesheet once, give the editor an explicit-height host, and retain
the complete value emitted by `onChange`:

```tsx
import { useState } from 'react';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

const initialContent: DocumentContent = {
  type: 'document',
  html: '<h1>Project brief</h1><p>Start editing here.</p>',
  pageSize: 'a4',
  pageColor: '#ffffff',
};

export function ProjectBrief() {
  const [content, setContent] = useState(initialContent);

  return (
    <main style={{ height: '100dvh', minHeight: 0 }}>
      <DocumentEditor
        content={content}
        onChange={setContent}
        theme="system"
      />
    </main>
  );
}
```

The editor owns the editing transaction. Your application decides when, where,
and how the controlled value is stored.

### Choose an entry point

| Entry point | Use it for |
| --- | --- |
| `@a3s-lab/office/react` | Lazy React editor components and preload helpers |
| `@a3s-lab/office/vue` | Vue 3 adapters with `v-model:content` |
| `@a3s-lab/office/web-component` | Framework-agnostic Custom Elements |
| `@a3s-lab/office/core` | Models, templates, import/export, file workflows, and Yjs bindings |
| `@a3s-lab/office/styles.css` | Shared editor and interaction-system styles |

Copyable integrations live in the
[component documentation](https://a3s-lab.github.io/Office/docs/components/).

## Files stay files

The Core API can import and export files without mounting an editor:

```ts
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '@a3s-lab/office/core';

const shell = createArtifact('blank-document');
const artifact = await importOfficeFile(file, {
  artifactId: shell.id,
  onProgress: ({ stage, progress }) => {
    console.info(stage, Math.round(progress * 100));
  },
});

const output = await createArtifactBlob(artifact);
```

Imported DOCX artifacts are source-backed. Safe, unedited package parts and
stable native identities are preserved under bounded rules; unsupported or
unsafe structures produce compatibility diagnostics or fail explicitly rather
than being attached to the wrong content. Structured Document values can also
cross process boundaries through the versioned snapshot codec.

See the
[Document](./docs/latest/en/components/document.mdx),
[Spreadsheet](./docs/latest/en/components/spreadsheet.mdx),
[Presentation](./docs/latest/en/components/presentation.mdx), and
[PDF](./docs/latest/en/components/pdf.mdx) references for file-specific
contracts.

## Collaboration without a bundled cloud

Every editor exposes the same transport-neutral collaboration boundary through
Yjs/Yrs content and Awareness. Browser users, native replicas, CLI sessions,
MCP clients, and A3S Code can participate in the same host-owned document.

A3S Office provides format-specific bindings, local undo, validated presence,
remote selections or locations, comments, suggestions, and conflict-local
typed mutations. The host provides rooms, authentication, authorization,
delivery, offline buffering, persistence, and the `Y.Doc`.

This separation makes collaboration optional infrastructure rather than a
mandatory account or storage service. Start with the
[collaboration guide](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
or the runnable
[A3S Boot example](./examples/collaboration-server/).

## Automation without UI scraping

The Rust CLI, standard MCP server, typed Rust API, and packaged Office Skill
share the same bounded file contracts:

```bash
# Validate and inspect a file.
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- view report.docx outline --json

# Apply an exact guarded mutation.
cargo run -p a3s-office-cli -- set report.docx /body \
  --find Draft --replace Final --json

# Expose the same contracts over standard MCP.
cargo run -p a3s-office-cli -- mcp
```

Native collaboration replicas can exchange standard Yjs updates and state
vectors, then apply typed Document, Markdown, Spreadsheet, Presentation, and
PDF mutations without interpreting Office's private CRDT schema.

Read the [automation guide](https://a3s-lab.github.io/Office/docs/automation/)
and [CLI reference](./docs/latest/en/cli-reference.md).

Before a coding agent uses cached guidance, verify the packaged Skill contract
with `cargo run -p a3s-office-cli -- skills manifest a3s-office --json` (or
`a3s-office skills manifest a3s-office --json` after installation). The
manifest includes byte counts and SHA-256 values for `SKILL.md` and every
bundled reference.

For browser-editor work, use the repository's local operator instead of
assembling ad-hoc shell conditionals. It is a Commander-based CLI with one
declarative matrix for Writer, Spreadsheet, Presentation, Markdown, and PDF:

```bash
bun run office:ops -- plan all --json
bun run office:ops -- capabilities --json
bun run office:ops -- doctor --json
bun run office:ops -- gate writer --run \
  --browser-driver standalone \
  --cdp-port 9345
bun run office:ops -- visual spreadsheet --project compact-768
```

`office-ui-ops plan <surface> --json` is the typed inventory entry point for that
matrix (`bun run office:ops -- plan <surface> --json`).

The declarative operator checks the focused A3S Test ACLs; `gate --run`
executes them as the primary interaction contract, while Playwright is only a
supplemental desktop/compact pixel baseline. Screenshots and diagnostics stay
below `.a3s-test/office-ops/`. It never polls CI or changes a committed visual
baseline. Use `a3s agent start/observe/act/finish` for bounded exploratory
sessions and `a3s cua certification --json` before native GUI work. The locked
CUA Driver 0.10.0 Windows profiles are currently unsupported, so Windows
browser-editor evidence uses A3S Test Web/CDP. On Windows,
supplying `--cdp-port` automatically compiles a native `.exe` adapter under
`.a3s-test/office-ops/`, keeping interactive arguments out of `.cmd` parsing.
`bun run office:ops -- wps-probe --connector` captures the bounded WPS COM
reference used for UI/OOXML parity; that probe is evidence, not a product
runtime. Use `--connector-type straight|elbow|curved` to select the typed WPS
reference shape through the Commander CLI. `wps-fields-probe --profile
numeric|common --json` records the installed WPS numeric or common field
instructions and feeds the corresponding A3S Test fixtures; it is likewise an
explicit local reference capture, never a CI prerequisite.
`wps-ui-probe --profile shell|fields|all --json` records the installed WPS
Writer window shell, Ribbon/status bars, and field-related command IDs as a
typed JSON UX reference. It starts and closes one owned COM instance, so the
receipt can guide UI decisions without turning WPS into a product runtime or
CI dependency.

For exploratory browser work, the same Commander operator exposes typed A3S
Test agent actions: `click`, `hover`, `focus`, `double-click`, `context-click`,
`fill`, `type`, `check`, `uncheck`, `select`, `drag`, `press`, `wheel`,
`viewport`, waits, assertions, screenshots, accessibility, console, and page
errors. Use the explicit target grammar (`@e7`, `css=...`, `role=role|name`,
`label=...`, `placeholder=...`, `testid=...`, `automation=...`, or `text=...`)
and keep the A3S Test `observe → one action → observe` lifecycle. This keeps
Codex editor operations typed and reproducible across all five surfaces.

## Current release

Version `0.99.0` makes table tblStyleRowBandSize revisions reviewable:

- **Row band size** — relationship-free `tblStyleRowBandSize` priors become
  reviewable `table-formatting` (`rowBandSize`) with accept/reject, live
  track-changes, and native DOCX export.
- **Boundary** — broader `tblPr` shapes stay opaque metadata or fail-closed;
  opaque table fixtures stay on `tblPrException`.

Earlier releases stay on the product
[What's new](https://a3s-lab.github.io/Office/docs/changelog.html) timeline and
the exhaustive engineering [changelog](./CHANGELOG.md). The
[live Playground](https://a3s-lab.github.io/Office/playground/) exposes these
workflows from the release-labelled template grid.

## Capabilities and boundaries

A3S Office favors explicit fidelity boundaries over silent approximation.
`Supported` means a path has deterministic behavioral or native round-trip
evidence. `Partial` means a useful path exists with a documented boundary.

### Document

**Strong paths:** structured authoring, pagination, tables, references, review,
large plain-document windows, source-aware DOCX round trips, and Document PDF
export with searchable Latin Helvetica vectors (optional host-registered CJK).

**Boundary:** long-tail DrawingML, fields, exact layout parity, full PDF/UA
structure trees, and unregistered non-Latin PDF glyphs remain partial or
fail-soft.

### Spreadsheet

**Strong paths:** sparse editing, dependency-aware formulas, tables,
sort/filter, validation, charts, pivots, XLSX round trips, and retention of the
workbook's native 1900 or 1904 date system.

**Boundary:** formula breadth, external data, macros, add-ins, and specialist
analysis remain incomplete.

### Presentation

**Strong paths:** scene editing, masters/layouts, transitions, bounded entrance
animations, presenter workflows, and PPTX round trips.

**Boundary:** broader shapes, effects, media, animation families, and exact
print/video output remain partial or unsupported.

### Markdown

**Strong paths:** GFM source and visual editing, synchronized preview,
collaboration, direct round trips, and native automation.

**Boundary:** Markdown remains format-native rather than acting as a conversion
layer for every rich-document construct.

### PDF

**Strong paths:** PDFium rendering, search, forms, annotations, save, and
bounded page organization.

**Boundary:** existing content-stream editing, OCR, signatures, optimization,
and trustworthy redaction need explicit providers or future engines.

The complete capability inventory, priorities, and exit evidence live in the
[capability roadmap](./ROADMAP.md). Collaboration delivery has its own
[roadmap](./COLLABORATION_ROADMAP.md).

## Project status

A3S Office is pre-1.0. Public models and component props are usable, but a
minor release may still include breaking model changes. Required migrations
are documented in the [changelog](./CHANGELOG.md).

The goal is predictable browser editing and honest file preservation, not
pixel parity with every desktop-suite feature. Unsupported semantics remain
visible compatibility boundaries. The
[editor quality roadmap](./docs/latest/en/editor-quality-roadmap.md) records
the evidence required to move each boundary.

## Development

Run the standard checks from the repository root:

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

Focused local A3S Test gates cover first-open focus and IME composition:

```bash
bun run test:e2e:initial-focus:check
bun run test:e2e:initial-focus
bun run test:e2e:presentation-chinese-ime:check
bun run test:e2e:presentation-chinese-ime
```

The controlled Document IME lifecycle also runs against Playwright's pinned
WebKit engine in CI:

```bash
bun run playground:ime:webkit
```

GitHub Actions does not require A3S Test. See
[CONTRIBUTING.md](./CONTRIBUTING.md),
[the E2E guide](./tests/e2e/README.md), and
[the visual-test guide](./visual-tests/README.md) for the complete validation
matrix.

## Documentation

- [Live Playground](https://a3s-lab.github.io/Office/playground/)
- [Documentation center](https://a3s-lab.github.io/Office/docs/)
- [What's new](https://a3s-lab.github.io/Office/docs/changelog.html)
- [React, Vue, Web Component, and Core API](https://a3s-lab.github.io/Office/docs/components/)
- [Real-time collaboration](https://a3s-lab.github.io/Office/docs/components/collaboration.html)
- [CLI, MCP, and Office Skill](https://a3s-lab.github.io/Office/docs/automation/)
- [Browser editor architecture](./docs/latest/en/browser-editor-architecture.md)
- [Native Office engine](./docs/latest/en/native-office-engine.md)
- [Editor quality roadmap](./docs/latest/en/editor-quality-roadmap.md)
- [Engineering changelog](./CHANGELOG.md)

## Community and security

Contributions are welcome. Read the
[contribution guide](./CONTRIBUTING.md) and
[Code of Conduct](./CODE_OF_CONDUCT.md) before opening a change. Report
suspected vulnerabilities through the private process in
[SECURITY.md](./SECURITY.md), not through a public issue.

## License

A3S Office is available under the [MIT License](./LICENSE). Bundled PDFium and
other third-party assets carry additional notices in
[THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md).
