# A3S Office

Open-source browser editors for documents, Markdown, spreadsheets,
presentations, and PDFs. The same editor engine used by A3S Web is packaged with
[Rslib](https://rslib.rs/) for React, Vue 3, Web Components, and framework-free
file workflows.

[![CI](https://github.com/A3S-Lab/Office/actions/workflows/ci.yml/badge.svg)](https://github.com/A3S-Lab/Office/actions/workflows/ci.yml)
[![Playground](https://github.com/A3S-Lab/Office/actions/workflows/pages.yml/badge.svg)](https://a3s-lab.github.io/Office/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Try the editors in the
[online Playground](https://a3s-lab.github.io/Office/).

## What is included

| Editor | Main capabilities |
| --- | --- |
| Document | Rich text, sections, page layout, headers and footers, comments, tracked changes, citations, notes, captions, cross-references, DOCX import/export, and PDF export |
| Markdown | TipTap editing, source and split views, headings, lists, quotes, links, images, code, native Markdown import/export, and read-only preview |
| Spreadsheet | Multiple sheets, formulas, charts, conditional formatting, pivot tables, comments, validation, protection, print settings, XLS/XLSX/ODS/CSV import, XLSX export, and PDF export |
| Presentation | Slides, masters and layouts, rich text, shapes, images, tables, charts, notes, comments, transitions, presenter view, PPTX import/export, and PDF export |
| PDF | Rendering, search, annotations, forms, redaction tools, and saving an edited copy |

The editors run entirely in the browser. A server is optional and remains in
control of persistence, collaboration, authentication, and AI requests.

The browser architecture is editor-specific rather than forcing every product
through one abstraction. Documents use one logical TipTap/ProseMirror tree;
automatic page layout runs in a Worker backed by the Rust WebAssembly kernel
and returns non-history page decorations. Spreadsheet and Presentation keep
their grid and scene models, using TipTap only for rich text where appropriate.
PDF rendering remains on PDFium. See
[Browser editor architecture](docs/browser-editor-architecture.md) for the
current fidelity boundary and migration plan.

The repository also contains the native `a3s-office` Rust CLI, its standard
MCP server, the Office Skill, and the OOXML engine used for deterministic
document automation. These surfaces can run directly or as an external
capability installed into A3S Use.

The built-in editor interface currently uses Simplified Chinese. Public APIs,
documentation, and source identifiers are in English.

## Installation

```bash
npm install @a3s-lab/office react react-dom
```

Vue applications also install Vue:

```bash
npm install @a3s-lab/office react react-dom vue
```

Import the stylesheet once near the application entry:

```ts
import '@a3s-lab/office/styles.css';
```

## React

Each editor is controlled: the host owns the content model and stores every
change.

```tsx
import { useState } from 'react';
import {
  createArtifact,
  type DocumentContent,
} from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import '@a3s-lab/office/styles.css';

export function App() {
  const initial = createArtifact('blank-document');
  const [content, setContent] = useState(
    initial.content as DocumentContent,
  );

  return (
    <div style={{ height: '100vh' }}>
      <DocumentEditor content={content} onChange={setContent} />
    </div>
  );
}
```

The React entry exports:

- `DocumentEditor`
- `MarkdownEditor`
- `SpreadsheetEditor`
- `PresentationEditor`
- `PdfViewer`
- their Props types, `OfficeFileAction`, `OfficeTheme`, and
  `defaultOfficeKernelWasmUrl` and `defaultPdfiumWasmUrl`

Set `preview` for a read-only representation. Use `theme="light"`,
`theme="dark"`, or `theme="system"`.

### Document layout kernel

`DocumentEditor` resolves `office-kernel.wasm` beside `react.js` and runs page
layout in its dedicated Worker. If assets are hosted elsewhere, pass an
absolute `kernelWasmUrl`:

```tsx
<DocumentEditor
  content={content}
  kernelWasmUrl="https://cdn.example.com/a3s-office/office-kernel.wasm"
  onChange={setContent}
/>
```

The current kernel provides automatic block-boundary pagination. Exact Word or
WPS line layout, floating objects, and multi-column flow remain explicit
fidelity gates rather than implied compatibility.

### Markdown

`MarkdownEditor` is powered by TipTap and keeps Markdown source as its
controlled value. Visual, source, and split editing all update the same
`MarkdownContent` model:

```tsx
import { useState } from 'react';
import type { MarkdownContent } from '@a3s-lab/office/core';
import { MarkdownEditor } from '@a3s-lab/office/react';

export function Notes() {
  const [content, setContent] = useState<MarkdownContent>({
    type: 'markdown',
    markdown: '# Notes\n\nStart writing here.',
  });

  return <MarkdownEditor content={content} onChange={setContent} />;
}
```

### PDF

The package emits `pdfium.wasm` beside `react.js`, and `PdfViewer` resolves that
file automatically:

```tsx
import { PdfViewer } from '@a3s-lab/office/react';

<PdfViewer
  fileName='contract.pdf'
  loadSource={() => fetch('/contract.pdf').then((response) => response.blob())}
  onSave={async (pdf) => {
    await uploadPdf(pdf);
    return true;
  }}
/>;
```

When assets are served from another location, pass an absolute `wasmUrl`.

## Vue 3

Vue components bridge the same editor engine and support `v-model:content`:

```vue
<script setup lang="ts">
import { ref } from 'vue';
import {
  createArtifact,
  type SpreadsheetContent,
} from '@a3s-lab/office/core';
import { SpreadsheetEditor } from '@a3s-lab/office/vue';
import '@a3s-lab/office/styles.css';

const artifact = createArtifact('blank-spreadsheet');
const content = ref(artifact.content as SpreadsheetContent);
</script>

<template>
  <div style="height: 100vh">
    <SpreadsheetEditor v-model:content="content" />
  </div>
</template>
```

Vue content editors emit `change` and `update:content`. Document, spreadsheet,
and presentation editors also emit `agent-request`; presentations emit
`start-slideshow`.

## Web Components

Register five custom elements once:

```ts
import {
  defineA3SOfficeElements,
} from '@a3s-lab/office/web-component';
import { createArtifact } from '@a3s-lab/office/core';
import '@a3s-lab/office/styles.css';

defineA3SOfficeElements();

const editor = document.querySelector('a3s-document-editor');
const artifact = createArtifact('blank-document');
editor.content = artifact.content;
editor.addEventListener('change', (event) => {
  artifact.content = event.detail;
});
```

```html
<a3s-document-editor theme="system"></a3s-document-editor>
```

Available tags are:

- `a3s-document-editor`
- `a3s-markdown-editor`
- `a3s-spreadsheet-editor`
- `a3s-presentation-editor`
- `a3s-pdf-viewer`

Complex values such as `content`, `fileActions`, `loadSource`, and `onSave` are
JavaScript properties. Simple values such as `preview`, `theme`, `save-status`,
`kernel-wasm-url`, and `wasm-url` are attributes or properties.

## Core API

`@a3s-lab/office/core` has no React component API. It provides content models,
templates, source-file registration, and file import/export:

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

Supported imports:

- DOCX, HTML, and plain text
- Markdown (`.md` and `.markdown`)
- XLSX, XLS, ODS, and CSV
- PPTX
- PDF

Native exports are DOCX, Markdown, XLSX, PPTX, and PDF. `downloadArtifact`
starts a browser download; `createArtifactBlob` returns a Blob for
application-managed storage.

The package emits `pptxgen.bundle.js` beside `core.js` and loads it only when a
presentation is exported. If package assets are hosted elsewhere, provide an
absolute URL:

```ts
const output = await createArtifactBlob(presentation, {
  pptxRuntimeUrl: 'https://cdn.example.com/a3s-office/pptxgen.bundle.js',
});
```

The Core API also exposes typed AI proposal requests. Editors send structured
selection context through `onAgentRequest`; the host decides which model to
call and whether a proposed change may be applied.

## Native CLI and A3S Use

Build and run the native CLI from this repository:

```bash
cargo run -p a3s-office-cli -- --version
cargo run -p a3s-office-cli -- --help
```

The same binary exposes a standard MCP server:

```bash
cargo run -p a3s-office-cli -- mcp
```

To install Office as an external A3S Use capability, create a package and
install it by its stable package ID:

```bash
./scripts/package-a3s-use-extension.sh /tmp/a3s-use-office

a3s-use component install a3s/office \
  --from /tmp/a3s-use-office \
  --allow-unsigned \
  --json

a3s-use office --version
a3s-use doctor office --json
a3s-use mcp serve office
```

The package manifest binds route `office` to this HTTPS repository and declares
the compatible A3S Use SemVer range. A3S Use verifies package trust and
compatibility before projecting its CLI, MCP, and Skill surfaces; it does not
clone and execute the repository.

See [Native Office engine](docs/native-office-engine.md) and
[CLI reference](docs/cli-reference.md) for the automation surface.

## Package entry points

| Import | Purpose |
| --- | --- |
| `@a3s-lab/office` | React editors plus Core API |
| `@a3s-lab/office/react` | React editors |
| `@a3s-lab/office/vue` | Vue 3 adapters |
| `@a3s-lab/office/web-component` | Custom elements |
| `@a3s-lab/office/core` | Models and file workflows |
| `@a3s-lab/office/styles.css` | Editor and design-system styles |

## Layout and assets

Editors fill their container. Give the container an explicit height:

```css
.office-host {
  width: 100%;
  height: min(900px, 100vh);
  min-height: 480px;
}
```

The stylesheet uses `--a3s-*` custom properties. Override them on the editor
container to match the host product.

The published package contains the browser Office layout Worker and WebAssembly
kernel, a 4.6 MB PDFium WebAssembly file, and the browser build of PptxGenJS.
Configure the CDN or static server to return `.wasm` as `application/wasm`.

## Development

Requirements:

- Node.js 20 or newer
- Bun 1.3 or newer
- Rust 1.85 or newer

```bash
bun install
bun run typecheck
bun run test
bun run kernel:test
bun run build
cargo fmt --all -- --check
cargo check --workspace --all-targets
cargo test --workspace
bun run playground
```

`bun run playground` starts the A3S Office product Playground. It includes the
Office-style workspace, document, Markdown, Spreadsheet, Presentation, and PDF
editors, the Office CLI introduction, and the downloadable `a3s-office` Skill.
Production output is generated in `dist/`. `bun run playground:build` packages
the Skill and creates the static Playground in `playground-dist/`.

Pushes to `main` deploy that static build to GitHub Pages through
`.github/workflows/pages.yml`. The workflow uses the Pages-provided base path,
so assets also resolve correctly from the project-site URL.

- [Office Playground](https://a3s-lab.github.io/Office/#office)
- [Office CLI guide](https://a3s-lab.github.io/Office/#cli)
- [Office CLI Skill](https://a3s-lab.github.io/Office/#skill)

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

## Status

The package is pre-1.0. Content models and component Props are public, but
breaking changes may still occur between minor releases. Release notes will
identify model migrations.

## License

A3S Office is available under the [MIT License](LICENSE). The bundled PDFium
WebAssembly artifact carries additional third-party notices; see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
