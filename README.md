<p align="center">
  <img src="assets/readme/hero.svg" width="1200" alt="A3S Office — five browser editors behind one typed boundary">
</p>

<p align="center">
  <strong>AI-native Office surfaces for the browser, backed by deterministic native automation.</strong>
</p>

<p align="center">
  Edit documents, Markdown, spreadsheets, presentations, and PDFs inside your product.<br>
  Keep persistence, identity, collaboration, authorization, and AI in your host application.
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI status" src="https://img.shields.io/github/actions/workflow/status/A3S-Lab/Office/ci.yml?branch=main&amp;style=flat-square&amp;label=CI"></a>
  <a href="https://a3s-lab.github.io/Office/"><img alt="Open the live Playground" src="https://img.shields.io/badge/Live_Playground-open-2f6fed?style=flat-square"></a>
  <a href="#project-status"><img alt="Project status: pre-1.0" src="https://img.shields.io/badge/status-pre--1.0-7a5bd6?style=flat-square"></a>
  <a href="LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/license-MIT-159469?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://a3s-lab.github.io/Office/">Playground</a> ·
  <a href="#quick-start">Quick start</a> ·
  <a href="#five-format-native-surfaces">Editors</a> ·
  <a href="#native-automation">Automation</a> ·
  <a href="#architecture">Architecture</a> ·
  <a href="CONTRIBUTING.md">Contributing</a>
</p>

---

A3S Office is an open-source Office engine for product teams that need rich
browser editing and agent-ready file automation without adopting an A3S
backend. It ships complete editor surfaces, typed host contracts, browser file
workflows, and a separate Rust automation plane.

- **Embed complete editors** through React, Vue 3, Web Components, or the
  framework-neutral Core API.
- **Keep format-native behavior** instead of forcing every file through one
  lowest-common-denominator model.
- **Own the product boundary**: your application controls content, storage,
  permissions, collaboration, and model providers.
- **Automate deterministically** through the native CLI, standard MCP server,
  or packaged Office Skill.

## See it working

The images below are committed visual-regression baselines from the real
[Playground](https://a3s-lab.github.io/Office/), not conceptual mockups.

<p align="center">
  <a href="visual-tests/__snapshots__/linux/desktop-1280/document.png">
    <img src="visual-tests/__snapshots__/linux/desktop-1280/document.png" alt="A3S Office document editor with a ribbon and paginated project brief" width="1280">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/spreadsheet.png" alt="A3S Office spreadsheet editor showing a quarterly execution plan">
      </a>
      <br><sub><strong>Spreadsheet</strong> — workbook editing and formula workflows</sub>
    </td>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/presentation.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/presentation.png" alt="A3S Office presentation editor with slide thumbnails and a slide canvas">
      </a>
      <br><sub><strong>Presentation</strong> — structured slides, objects, and presenter flows</sub>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/markdown.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/markdown.png" alt="A3S Office Markdown editor in synchronized source and preview mode">
      </a>
      <br><sub><strong>Markdown</strong> — GFM source and synchronized preview</sub>
    </td>
    <td width="50%" valign="top">
      <a href="visual-tests/__snapshots__/linux/desktop-1280/pdf.png">
        <img src="visual-tests/__snapshots__/linux/desktop-1280/pdf.png" alt="A3S Office PDF editor with search, annotation, save, and download controls">
      </a>
      <br><sub><strong>PDF</strong> — PDFium rendering, forms, annotations, and save</sub>
    </td>
  </tr>
</table>

## Why A3S Office

- **Product-native UI** — Complete Office-style surfaces with no required
  backend, account system, or storage model.
- **Predictable state** — Controlled content values, typed callbacks, explicit
  file actions, conflict-aware document edits, and one document typography
  baseline across editing, preview, and PDF rendering. Word editing and
  read-only preview retain the same canonical TipTap tree, while browser PDF
  export captures that same Worker/WASM page layout instead of rebuilding
  content from compatibility HTML.
- **Framework choice** — React components, Vue 3 adapters, Custom Elements,
  and a framework-neutral Core API over the same engine.
- **Responsive computation** — Lazy editor chunks, cancellable Workers,
  CSS-compatible font-weight matching, Rust WebAssembly layout and
  calculation, and PDFium rendering.
- **AI without UI scraping** — Typed agent ports and host-defined selection
  actions receive structured context and editing commands.
- **Automation outside the browser** — The native Rust CLI, standard MCP
  server, and Office Skill share bounded mutation contracts.

## Quick start

### Try the product locally

Node.js 20+, Bun 1.3+, and Rust 1.85+ are required.

```bash
git clone https://github.com/A3S-Lab/Office.git
cd Office
bun install --frozen-lockfile
bun run playground
```

Then open the local URL printed by the development server. For a zero-install
tour, use the [live Playground](https://a3s-lab.github.io/Office/).

> [!NOTE]
> The `@a3s-lab/office` package is prepared for npm, but the first public npm
> release has not been published yet. Use the Playground or build from source
> until the release channel is available.

### Embed a controlled React editor

Once the package release is available, install it with its React peers:

```bash
bun add @a3s-lab/office react react-dom
```

Import the stylesheet once, give the editor an explicit-height host, and store
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

The editor owns editing, layout, import/export, and browser rendering. The host
owns persistence and decides when, where, and how the emitted content is saved.

### Export the live Word layout to PDF

Give each mounted document a stable `artifactId`, then pass the matching,
current artifact to `downloadArtifactPdf`. Export waits for the live pagination
surface and crops the physical pages computed by the editor:

```tsx
import { useState } from 'react';
import {
  createArtifact,
  downloadArtifactPdf,
} from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';

export function DocumentWithPdfExport() {
  const [artifact, setArtifact] = useState(() =>
    createArtifact('blank-document'),
  );

  if (artifact.content.type !== 'document') return null;

  return (
    <main style={{ display: 'flex', height: '100dvh', flexDirection: 'column' }}>
      <button
        type="button"
        onClick={() => void downloadArtifactPdf(artifact)}
      >
        Export PDF
      </button>
      <section style={{ flex: 1, minHeight: 0 }}>
        <DocumentEditor
          artifactId={artifact.id}
          content={artifact.content}
          onChange={(content) =>
            setArtifact((current) => ({
              ...current,
              content,
              revision: current.revision + 1,
              updatedAt: Date.now(),
            }))
          }
        />
      </section>
    </main>
  );
}
```

The live path preserves automatic page-break decorations, shaped text, table
continuations, and page chrome. It currently rasterizes each physical page
into the PDF; searchable text and vector output remain future fidelity work.

### Choose an entry point

- `@a3s-lab/office/react` — Lazy React editor components and preload helpers.
- `@a3s-lab/office/vue` — Vue 3 adapters with `v-model:content`.
- `@a3s-lab/office/web-component` — Custom Elements for framework-agnostic UI
  composition.
- `@a3s-lab/office/core` — Typed models, templates, import, export, and browser
  file workflows.
- `@a3s-lab/office/styles.css` — Shared editor and interaction-system styles.

Copyable React, Vue, and Web Component examples live in the
[integration guide](https://a3s-lab.github.io/Office/#guide).

## Five format-native surfaces

Each surface keeps a canonical model that matches its file format and user
interaction model.

- **Document** — Pagination, sections, focused page-setup tabs, clean page
  margins with on-demand header/footer editing, outline navigation, styles,
  precision table sizing and autofit, comments, tracked changes, citations,
  notes, host-defined selection menus, and shared edit/preview/PDF typography
  and page-chrome placement. Editing, read-only preview, and browser PDF export
  share one live pagination result. _DOCX import/export; PDF export._
- **Markdown** — GFM source, visual editing, synchronized and resizable split
  preview, source-native undo/redo with typing coalescing and selection restore,
  source-aware ribbon formatting and shortcuts, empty-source guidance, task
  lists, tables, links, images, code, and host-defined selection menus.
  _Markdown import/export._
- **Spreadsheet** — Multiple sheets, formulas, dependency-aware recalculation,
  live selection statistics, an A3S-owned worksheet bar with lifecycle and
  color controls, focus-safe formatting, history, worksheet navigation,
  cut/copy/paste and clear shortcuts, direct font-family, vertical-alignment,
  text-wrap, number/percent, and decimal controls, charts, validation,
  protection, comments, and print settings.
  _XLSX, XLS, ODS, and CSV import; XLSX and PDF export._
- **Presentation** — Typed slide scene graph, multi-selection, groups, object
  transforms, keyboard-accessible table-size insertion, native slide and object
  context actions, guides, comments, transitions, one-step slideshow from the
  beginning or current slide, and presenter view. _PPTX import/export; PDF
  export._
- **PDF** — PDFium rendering, navigation, search, form filling, annotations,
  annotation color, opacity, compatible stroke-width controls, history,
  responsive navigation and zoom controls, and save. _PDF open/save._

Document and Markdown accept public TipTap Extensions. Spreadsheet,
Presentation, and PDF expose stable host ports rather than their internal
command contexts. Editor engines and large runtime assets load only when that
surface is requested.

## Controlled by design

A3S Office is headless at the product boundary, not at the UI boundary. The
package includes complete toolbars, ribbons, panes, popovers, and dialogs while
leaving product infrastructure to the host.

**A3S Office owns:** editing models and commands; import, export, layout, and
rendering; editor UI and responsive interactions; typed selection and agent
ports.

**Your host owns:** persistence and version history; identity, permissions,
and collaboration; the application shell and navigation; AI providers,
prompts, policy, and request lifecycle.

Document and Markdown selection menus can be replaced with host-defined typed
actions. Each action receives an immutable selection snapshot, nearby text,
the complete controlled content, and conflict-aware `replaceText`,
`insertBefore`, `insertAfter`, and `copyText` commands. Markdown snapshots also
identify whether the source or visual surface owns the selection. Async edits
track unrelated visual-editor transactions and fail with `stale-selection`
instead of changing the wrong text.

## Browser file workflows

The Core API creates typed blank artifacts and performs browser-side import or
export without mounting an editor:

```ts
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '@a3s-lab/office/core';

const imported = await importOfficeFile(file);
const output = await createArtifactBlob(imported);
const blankDeck = createArtifact('blank-presentation');
```

Use `downloadArtifact` to start a browser download or
`createArtifactBlob` when your application owns upload and persistence.

## Native automation

The repository also contains a native Rust engine for deterministic reads,
validation, mutation, batch operations, semantic previews, screenshots, and
file watching across Office packages.

```bash
# Run from the repository root
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- view report.docx outline --json
cargo run -p a3s-office-cli -- set report.docx /body --find Draft --replace Final --json

# Start the standard MCP server
cargo run -p a3s-office-cli -- mcp
```

CLI, MCP, the typed Rust API, and the packaged Office Skill share the same
bounded contracts. They inspect and modify files without launching desktop
Office or scraping editor UI.

Read the [native engine design](docs/native-office-engine.md), the complete
[CLI reference](docs/cli-reference.md), or the Playground's
[automation guide](https://a3s-lab.github.io/Office/#guide/automation).

## Architecture

<p align="center">
  <a href="assets/readme/architecture.svg">
    <img src="assets/readme/architecture.svg" width="1200" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

The browser plane combines controlled editor surfaces, framework adapters,
Workers, Rust WebAssembly, and PDFium. The native plane keeps filesystem and
OOXML package concerns in a separate Rust core. Both planes expose typed
contracts; neither requires an A3S backend.

For engine ownership, Worker/WASM boundaries, delivery stages, and performance
gates, see [Browser editor architecture](docs/browser-editor-architecture.md).

## Project status

A3S Office is pre-1.0. Content models and component props are public, but a
minor release may still include breaking model changes. Required migrations
will be called out in the [changelog](CHANGELOG.md).

The project targets predictable browser editing and file preservation, not
pixel parity with every Microsoft Office or WPS feature. Unsupported OOXML
semantics and fidelity gaps remain explicit compatibility boundaries instead
of being silently approximated. Track product depth and release evidence in
the [editor quality roadmap](docs/editor-quality-roadmap.md).

## Development

Run focused checks from the repository root:

```bash
bun install --frozen-lockfile
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run build
```

Start the integration Playground with `bun run playground`. The full pull
request checklist also covers the Rust workspace, browser bundle budget, and
committed visual contracts; see [CONTRIBUTING.md](CONTRIBUTING.md).

## Documentation

- [Live Playground and integration guide](https://a3s-lab.github.io/Office/)
- [Browser editor architecture](docs/browser-editor-architecture.md)
- [Native Office engine](docs/native-office-engine.md)
- [CLI reference](docs/cli-reference.md)
- [Editor quality roadmap](docs/editor-quality-roadmap.md)
- [Changelog](CHANGELOG.md)

## Community and security

Contributions are welcome. Read the [contribution guide](CONTRIBUTING.md) and
[Code of Conduct](CODE_OF_CONDUCT.md) before opening a change. Report suspected
vulnerabilities through the private process in [SECURITY.md](SECURITY.md), not
through a public issue.

## License

A3S Office is available under the [MIT License](LICENSE). Bundled PDFium and
other third-party assets carry additional notices in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
