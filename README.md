# A3S Office

<p align="center">
  <strong>Embeddable Browser Office Editors for A3S</strong>
</p>

<p align="center">
  <em>Edit documents, Markdown, spreadsheets, presentations, and PDFs with React, Vue, or Web Components</em>
</p>

<p align="center">
  <a href="#overview">Overview</a> •
  <a href="#features">Features</a> •
  <a href="#quick-start">Quick Start</a> •
  <a href="#editor-model">Editor Model</a> •
  <a href="#native-automation">Native Automation</a> •
  <a href="#architecture">Architecture</a> •
  <a href="#development">Development</a>
</p>

<p align="center">
  <a href="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/A3S-Lab/Office/actions/workflows/ci.yml/badge.svg"></a>
  <a href="https://a3s-lab.github.io/Office/"><img alt="Playground" src="https://github.com/A3S-Lab/Office/actions/workflows/pages.yml/badge.svg"></a>
  <a href="LICENSE"><img alt="MIT License" src="https://img.shields.io/badge/License-MIT-blue.svg"></a>
</p>

---

## Overview

**A3S Office** is an open-source component library for browser-based Office
editing. It provides document, Markdown, spreadsheet, presentation, and PDF
surfaces through one npm package while keeping each editor's native interaction
model.

The package is controlled by the host application. A3S Office owns editing,
layout, import, export, and browser rendering; the host owns persistence,
authentication, collaboration, authorization, and AI requests. A server is
optional.

The same components power the
[A3S Office Playground](https://a3s-lab.github.io/Office/). The repository also
contains the native `a3s-office` CLI, a standard MCP server, and the downloadable
Office Skill for deterministic document automation.

### Basic usage

```tsx
import { useState } from 'react';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

const initialDocument: DocumentContent = {
  type: 'document',
  html: '<h1>Project brief</h1><p>Start editing here.</p>',
  pageSize: 'a4',
};

export function App() {
  const [content, setContent] = useState(initialDocument);

  return (
    <div style={{ height: '100vh', minHeight: 0 }}>
      <DocumentEditor content={content} onChange={setContent} />
    </div>
  );
}
```

## Features

- **Five Editor Surfaces**: Edit documents, Markdown, spreadsheets,
  presentations, and PDFs without leaving the host product
- **Framework Entry Points**: Use React components, Vue 3 adapters, custom
  elements, or the framework-neutral Core API
- **Controlled Content**: Persist typed content models in the host and receive
  every user change through explicit callbacks or events
- **Office File Workflows**: Import and export DOCX, XLSX, PPTX, Markdown, PDF,
  HTML, text, ODS, XLS, and CSV where supported
- **Document Review**: Add anchored comments, replies, resolved state, tracked
  changes, citations, notes, captions, and cross-references
- **Browser-Native Layout**: Combine TipTap, editor-specific scene models,
  Workers, Rust WebAssembly, and PDFium without a remote rendering service
- **Typed Commands**: Dispatch editor actions through typed controllers rather
  than interpreting visible labels or DOM text
- **Lazy Editor Engines**: Load only the selected editor and its large runtime
  assets
- **Native Automation**: Inspect and modify Office packages through the Rust
  CLI, JSON output, MCP, or the Office Skill

### Editor matrix

| Editor | Editing engine | Main capabilities | Native files |
| --- | --- | --- | --- |
| Document | TipTap/ProseMirror + Worker/Rust-WASM layout | Sections, page layout, headers and footers, tables, images, comments, tracked changes, citations, notes, captions, references | DOCX import/export, PDF export |
| Markdown | TipTap + Markdown source model | Source and preview split view, visual editing, headings, lists, links, images, code | MD import/export |
| Spreadsheet | Fortune Sheet + typed workbook model | Multiple sheets, formulas, formatting, charts, validation, protection, comments, print settings | XLSX/XLS/ODS/CSV import, XLSX/PDF export |
| Presentation | Scene graph + TipTap text editing | Slides, layouts, shapes, images, tables, charts, comments, transitions, presenter view | PPTX import/export, PDF export |
| PDF | PDFium WebAssembly | Rendering, navigation, search, form filling, annotations, history, save | PDF open/save |

### Package matrix

| Entry point | Purpose |
| --- | --- |
| `@a3s-lab/office` | React editors plus the Core API |
| `@a3s-lab/office/react` | React editor components |
| `@a3s-lab/office/vue` | Vue 3 component adapters |
| `@a3s-lab/office/web-component` | Custom element definitions |
| `@a3s-lab/office/core` | Content models, templates, file import, and export |
| `@a3s-lab/office/styles.css` | Editor and design-system styles |

## Quick Start

### Installation

React applications install the package and its peer dependencies:

```bash
bun add @a3s-lab/office react react-dom
```

Vue applications also install Vue. The current Vue adapter uses the same React
editor engine:

```bash
bun add @a3s-lab/office react react-dom vue
```

Import the shared stylesheet once at the application entry:

```ts
import '@a3s-lab/office/styles.css';
```

Every editor fills its host. Give that host an explicit height and allow nested
flex layouts to shrink:

```css
.office-host {
  width: 100%;
  height: min(900px, 100vh);
  min-height: 480px;
}
```

### React

React editors are controlled components. Persist the complete value emitted by
`onChange`; document content may contain both compatibility HTML and the
versioned structured model.

```tsx
import { useState } from 'react';
import { createArtifact, type SpreadsheetContent } from '@a3s-lab/office/core';
import { SpreadsheetEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

export function WorkbookPage() {
  const artifact = createArtifact('blank-spreadsheet');
  const [content, setContent] = useState(
    artifact.content as SpreadsheetContent,
  );

  return (
    <div className="office-host">
      <SpreadsheetEditor
        content={content}
        onChange={setContent}
        theme="light"
      />
    </div>
  );
}
```

Set `preview` for a read-only surface and use `theme="light"`,
`theme="dark"`, or `theme="system"`. Warm an editor chunk from an intent signal
when opening latency matters:

```tsx
import { preloadOfficeEditor } from '@a3s-lab/office/react';

<button
  onFocus={() => void preloadOfficeEditor('document')}
  onPointerEnter={() => void preloadOfficeEditor('document')}
>
  New document
</button>;
```

### Vue 3

Vue content editors support `v-model:content` and also emit `change`.

```vue
<script setup lang="ts">
import { ref } from 'vue';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/vue';
import '@a3s-lab/office/styles.css';

const content = ref<DocumentContent>({
  type: 'document',
  html: '<h1>Project brief</h1><p>Start editing here.</p>',
  pageSize: 'a4',
});
</script>

<template>
  <div class="office-host">
    <DocumentEditor v-model:content="content" theme="light" />
  </div>
</template>
```

Document, Spreadsheet, and Presentation emit `agent-request`. Presentation
also emits `start-slideshow`.

### Web Components

Register the custom elements once, then pass complex values through JavaScript
properties.

```ts
import {
  defineA3SOfficeElements,
} from '@a3s-lab/office/web-component';
import '@a3s-lab/office/styles.css';

defineA3SOfficeElements();

const editor = document.querySelector('a3s-document-editor');
if (editor) {
  editor.content = {
    type: 'document',
    html: '<h1>Project brief</h1><p>Start editing here.</p>',
    pageSize: 'a4',
  };
  editor.addEventListener('change', (event) => {
    console.log(event.detail);
  });
}
```

```html
<a3s-document-editor theme="system"></a3s-document-editor>
```

Available tags are `a3s-document-editor`, `a3s-markdown-editor`,
`a3s-spreadsheet-editor`, `a3s-presentation-editor`, and `a3s-pdf-viewer`.

### Core file workflows

The Core API has no component dependency. It creates typed artifacts and
performs browser file import or export.

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

`downloadArtifact` starts a browser download. `createArtifactBlob` returns a
`Blob` for application-managed persistence.

See the online
[component integration guide](https://a3s-lab.github.io/Office/#guide) for
copyable React, Vue, and Web Component examples.

## Editor Model

### Controlled state

The host owns the current content value. Editors do not select a database,
collaboration backend, identity provider, or model provider. File actions are
also optional host commands; do not duplicate them when the surrounding
product already provides open, save, or export controls.

`DocumentContent` stores compatibility HTML and may carry a structured,
versioned model. Persist the complete emitted value. Directly replacing HTML
invalidates an older model safely.

### Layout and runtime assets

Document pagination runs in a dedicated Worker backed by
`office-kernel.wasm`. The package includes deterministic Latin, Simplified
Chinese, Arabic, and Hebrew layout fonts. Presentation alignment uses the same
kernel. PDF rendering uses `pdfium.wasm`, while presentation export loads the
browser PptxGenJS runtime only when needed.

Applications serving package assets from a separate CDN can pass explicit
`kernelWasmUrl`, `layoutFonts`, `wasmUrl`, or `pptxRuntimeUrl` values. Static
servers must return WebAssembly as `application/wasm` and allow fonts through
the required CORS policy.

### Fidelity boundary

A3S Office aims for predictable browser editing and native file preservation;
it does not claim pixel parity with every Microsoft Office or WPS feature.
Unsupported OOXML semantics, arbitrary floating-object layout, complete font
substitution, modern threaded comments, advanced spreadsheet calculation, and
the remaining presentation scene features stay explicit fidelity gates.

See [Browser editor architecture](docs/browser-editor-architecture.md) for
engine ownership, Worker/WASM boundaries, delivery stages, and performance
gates.

## Native Automation

The repository includes the native Rust CLI and its standard MCP server:

```bash
cargo run -p a3s-office-cli -- --version
cargo run -p a3s-office-cli -- validate report.docx --json
cargo run -p a3s-office-cli -- mcp
```

The CLI inspects and modifies OOXML packages without launching desktop Office.
It supports typed reads, queries, validation, mutation, batch operations,
semantic previews, screenshots, and file watching.

Package the capability for A3S Use through its standard external-repository
layer:

```bash
./scripts/package-a3s-use-extension.sh /tmp/a3s-use-office

a3s-use component install a3s/office \
  --from /tmp/a3s-use-office \
  --allow-unsigned \
  --json
```

See [Native Office engine](docs/native-office-engine.md) and
[CLI reference](docs/cli-reference.md) for the complete automation contract.

## Architecture

Each editor keeps its own canonical model instead of forcing every file type
through one abstraction:

```text
React / Vue / Web Components
            │
 controlled editor components
            │
 typed commands + content models
            │
 ┌──────────┼───────────┬─────────────┐
 TipTap   Workbook    Slide scene   PDFium
            │              │
       Worker + Rust WebAssembly
            │
 DOCX / Markdown / XLSX / PPTX / PDF workflows
```

Documents use one logical TipTap/ProseMirror tree and non-history pagination
decorations. Markdown keeps source as its controlled value. Spreadsheet keeps
the workbook grid canonical. Presentation keeps a slide scene graph and mounts
TipTap only for selected rich text. PDF commands call typed PDFium
capabilities directly.

Public framework adapters converge on the same React editor engine. The
framework-neutral Core entry owns models and file workflows. The native Rust
workspace is a separate automation boundary and does not sit on the browser
editing path.

## Development

Requirements:

- Node.js 20 or newer
- Bun 1.3 or newer
- Rust 1.85 or newer

Run checks from the A3S Office repository:

```bash
bun install
bun run format:check
bun run lint
bun run typecheck
bun run test
bun run kernel:test
bun run build
bun run playground:typecheck
bun run playground:build
bun run playground:bundle-check
bun run playground:visual
cargo fmt --all -- --check
cargo check --workspace --all-targets
cargo test --workspace
```

Start the product Playground locally:

```bash
bun run playground
```

The production Playground is generated in `playground-dist/`. Pushes to `main`
deploy it with GitHub Actions and GitHub Pages.

- [Office Playground](https://a3s-lab.github.io/Office/#office)
- [Component integration](https://a3s-lab.github.io/Office/#guide)
- [Office CLI guide](https://a3s-lab.github.io/Office/#cli)
- [Office CLI Skill](https://a3s-lab.github.io/Office/#skill)

Visual tests cover all five editor shells at 1280 × 800 and 768 × 800. Update
committed baselines only for intentional UI changes:

```bash
bun run playground:visual:update
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

## Status

A3S Office is pre-1.0. Content models and component Props are public, but minor
releases may still contain breaking model changes. Release notes will identify
required migrations.

## License

A3S Office is available under the [MIT License](LICENSE). Bundled PDFium and
other third-party assets carry additional notices in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
