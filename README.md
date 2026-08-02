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
  <a href="https://a3s-lab.github.io/Office/docs/">Documentation</a> ·
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
- **Accessible responsive shell** — Compact sidebars, slide navigation, AI
  panes, dialogs, menus, and popovers share bounded keyboard navigation,
  background isolation, topmost Escape handling, and focus restoration.
  Shared color palettes expand to an eight-column touch layout on phones while
  preserving spatial arrow-key navigation.
- **Editor-scoped zoom** — Status controls and Ctrl/Cmd + mouse-wheel gestures
  share each surface's bounded zoom model without changing the host browser's
  page scale.
- **Predictable state** — Controlled content values, typed callbacks, explicit
  file actions, conflict-aware document edits, and one document typography
  baseline across editing, preview, and PDF rendering. Word editing and
  read-only preview retain the same canonical TipTap tree, while browser PDF
  export captures that same Worker/WASM page layout instead of rebuilding
  content from compatibility HTML. Caption numbering and cross-reference
  validity also update in that same transaction graph, including a truthful
  missing-target state after deletion.
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

### Embed a controlled React editor

Install the public package with its React peers:

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
[component documentation](https://a3s-lab.github.io/Office/docs/components/).

## Five format-native surfaces

Each surface keeps a canonical model that matches its file format and user
interaction model.

- **Document** — Pagination, sections, focused page-setup tabs, clean page
  margins with on-demand header/footer editing, outline navigation with
  contextual full-text results, keyboard-operated live physical-page raster
  thumbnails with viewport-bounded capture, and shared virtualized
  long-document windows for page buttons, heading rows, full-text results, and
  tracked-revision and anchored-comment review that preserve native scroll
  distance, sparse current/selection/draft pins, and Home/End reachability;
  paragraph styles and phone-sized paragraph-spacing and pagination controls
  with touch targets and exact invoker-focus restoration; stable compact-ribbon
  edge paging and viewport-bounded list galleries with touch-sized numbering
  actions, selection-preserving bullet and numbering commands, and explicit
  Escape-to-invoker focus restoration;
  grouped bundled, common system, monospace, imported, and host-provided font
  choices with live typeface previews, precision table sizing and autofit,
  viewport-safe comment drafting and focus-preserving citation drafts, tracked
  changes, notes,
  pointer- and Shift+F10-accessible host-defined selection menus, and shared
  edit/preview/PDF typography and page-chrome placement. Editing, read-only
  preview, and browser PDF export share one live pagination result.
  _DOCX import/export; PDF export._
- **Markdown** — GFM source, visual editing, synchronized and resizable split
  preview, source-native undo/redo with typing coalescing and selection restore,
  source-aware ribbon formatting and shortcuts, empty-source guidance, task
  lists, tables, links, images, code, and keyboard-accessible host-defined
  selection menus across both editing surfaces. _Markdown import/export._
- **Spreadsheet** — Multiple sheets, formulas, dependency-aware recalculation,
  live selection statistics, an A3S-owned worksheet bar with lifecycle and
  color controls, safe deletion confirmation, in-place rename validation, a
  grid-aligned single-row responsive workbook footer that keeps the active
  worksheet and its actions visible across window and phone-width changes,
  selection-preserving undo/redo, deterministic Arrow, Enter, Tab, Home,
  PageUp/PageDown, row, column, and all-cells selection shortcuts, Cmd/Ctrl
  formatting through root-scoped capture before vendor grid listeners, an
  editor-owned Cmd/Ctrl+F search bar that finds displayed values, raw values,
  formulas, and sparse cells without opening browser Find, with a viewport-safe
  phone layout, touch-sized controls, and exact grid-focus restoration,
  desktop workbook task panes that become focus-contained phone dialogs with
  an inert workbook background and exact ribbon-invoker restoration,
  Shift+F11 creation, Ctrl/Cmd+PageUp/PageDown switching in edit and read-only
  preview,
  one shared cell/worksheet context-menu surface with executable shortcut
  hints, keyboard-operated cells and sheet tabs, direct type-to-edit,
  permission-resilient multi-cell cut/copy/paste, clear, F2 editing, and
  focus-safe Escape behavior, direct
  font-family, vertical-alignment, text-wrap, number/percent, and decimal
  controls, charts, validation, protection, comments, and print settings.
  _XLSX, XLS, ODS, and CSV import; XLSX and PDF export._
- **Presentation** — Typed slide scene graph, multi-selection, groups, object
  transforms, keyboard-accessible table-size insertion, Shift+F10-accessible
  native slide and object context actions, guides, comments, transitions,
  editor-scoped formatting and clipboard shortcuts that restore the selected
  object, a dismissible phone slide navigator, a top-aligned phone canvas with
  a priority-aware status bar, one-step slideshow from the beginning or current
  slide, keyboard-complete playback, and a responsive presenter view with one
  navigation strip, current/next context, speaker notes, and a session timer.
  _PPTX import/export; PDF export._
- **PDF** — PDFium rendering, navigation, search, form filling, annotations,
  annotation color, opacity, compatible stroke-width controls, history,
  a scrollable page-thumbnail rail with current-page synchronization and
  focus-synchronized Arrow/Home/End navigation, bounded rendering for long
  files, and a dismissible page drawer whose phone trigger stays in the page
  controls instead of covering PDF content,
  focus-safe page and search drafts, responsive search-result, navigation, and
  zoom controls, editor-scoped shortcuts, and a keyboard-operated overflow menu
  that retains secondary annotation tools and appearance settings on phones.
  _PDF open/save._

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
instead of changing the wrong text. The Playground's open-ended question action
enters a focused draft before dispatch, so a host never receives an unfinished
“Question:” request; attached context stays available without dominating the
assistant surface.

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
validation, mutation, batch operations, exact natural-unit inventories,
sibling-isolated semantic previews, screenshots, file watching, and a
source-bound exact-layout raster boundary across Office packages. The first
layout-authoritative route supports image-only PPTX slides whose single opaque
PNG exactly covers the declared slide surface. The optional Rust `pdfium`
feature adds bounded, one-based PDF page inventory and exact page PNGs through
an explicit host-supplied PDFium 7881 library. It records media/crop boxes,
rotation, physical and pixel geometry, source and engine hashes, and never
downloads a runtime or introduces a Browser dependency. Consumers can inspect
selected pages from one previously validated complete inventory without
rescanning the full document; render still revalidates the immutable source and
actual page profile before publication. The same retained inventory now
authorizes bounded native PDF text-layer extraction with source-order Unicode,
exact UTF-8/UTF-16 ranges, optional glyph boxes in PDF coordinates, and bounded
document outlines with exact page targets. Text and outline calls revalidate
the immutable source, reuse inventory authority, and return typed limit or
unsupported failures without OCR or Browser access. Richer slides and formats
without an authoritative provider remain typed unsupported instead of being
relabeled semantic previews.
Native DOCX and PPTX table reads also normalize merged cells into one-based
logical row and column coordinates, row and column spans, and stable anchor
references for covered physical cells.

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

Read the [native engine design](docs/latest/en/native-office-engine.md), the
complete [CLI reference](docs/latest/en/cli-reference.md), or the published
[CLI and Skill guide](https://a3s-lab.github.io/Office/docs/automation/).

## Architecture

<p align="center">
  <a href="assets/readme/architecture.svg">
    <img src="assets/readme/architecture.svg" width="1200" alt="A3S Office architecture with browser editing and native automation planes controlled by the host product">
  </a>
</p>

The browser plane combines controlled editor surfaces, framework adapters,
Workers, Rust WebAssembly, and PDFium. The native plane keeps filesystem and
OOXML package concerns in a separate Rust core, with a host-injected optional
PDFium provider for read-only PDF page evidence. Both planes expose typed
contracts; neither requires an A3S backend.

For engine ownership, Worker/WASM boundaries, delivery stages, and performance
gates, see
[Browser editor architecture](docs/latest/en/browser-editor-architecture.md).

## Project status

A3S Office is pre-1.0. Content models and component props are public, but a
minor release may still include breaking model changes. Required migrations
will be called out in the [changelog](CHANGELOG.md).

The project targets predictable browser editing and file preservation, not
pixel parity with every Microsoft Office or WPS feature. Unsupported OOXML
semantics and fidelity gaps remain explicit compatibility boundaries instead
of being silently approximated. Track product depth and release evidence in
the [editor quality roadmap](docs/latest/en/editor-quality-roadmap.md).

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

The published website keeps the interactive Playground at its root and builds
the searchable Rspress documentation center under `/docs/`. Simplified Chinese
is the stable default, English remains available from the language menu, and
the version menu switches between `latest` and frozen release documentation.
Search is scoped to the active language and version. Both surfaces use the
same deployment base, so preview, Pages, and fork deployments keep working
without hard-coded return URLs.

- [Live Playground](https://a3s-lab.github.io/Office/)
- [Documentation center](https://a3s-lab.github.io/Office/docs/)
- [A3S Office 0.1.0 documentation](https://a3s-lab.github.io/Office/docs/0.1.0/)
- [React, Vue, and Web Component integration](https://a3s-lab.github.io/Office/docs/components/)
- [Office CLI and coding-agent Skill](https://a3s-lab.github.io/Office/docs/automation/)
- [Browser editor architecture](docs/latest/en/browser-editor-architecture.md)
- [Native Office engine](docs/latest/en/native-office-engine.md)
- [CLI reference](docs/latest/en/cli-reference.md)
- [Editor quality roadmap](docs/latest/en/editor-quality-roadmap.md)
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
