# A3S Office

Open-source browser editors for documents, spreadsheets, presentations, and
PDFs. The same editor engine used by A3S Web is packaged with
[Rslib](https://rslib.rs/) for React, Vue 3, Web Components, and framework-free
file workflows.

[![CI](https://github.com/A3S-Lab/Office/actions/workflows/ci.yml/badge.svg)](https://github.com/A3S-Lab/Office/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## What is included

| Editor | Main capabilities |
| --- | --- |
| Document | Rich text, sections, page layout, headers and footers, comments, tracked changes, citations, notes, captions, cross-references, DOCX import/export, and PDF export |
| Spreadsheet | Multiple sheets, formulas, charts, conditional formatting, pivot tables, comments, validation, protection, print settings, XLS/XLSX/ODS/CSV import, XLSX export, and PDF export |
| Presentation | Slides, masters and layouts, rich text, shapes, images, tables, charts, notes, comments, transitions, presenter view, PPTX import/export, and PDF export |
| PDF | Rendering, search, annotations, forms, redaction tools, and saving an edited copy |

The editors run entirely in the browser. A server is optional and remains in
control of persistence, collaboration, authentication, and AI requests.

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
- `SpreadsheetEditor`
- `PresentationEditor`
- `PdfViewer`
- their Props types, `OfficeFileAction`, `OfficeTheme`, and
  `defaultPdfiumWasmUrl`

Set `preview` for a read-only representation. Use `theme="light"`,
`theme="dark"`, or `theme="system"`.

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

Vue adapters emit `change`, `agent-request`, and, for presentations,
`start-slideshow`.

## Web Components

Register four custom elements once:

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
- `a3s-spreadsheet-editor`
- `a3s-presentation-editor`
- `a3s-pdf-viewer`

Complex values such as `content`, `fileActions`, `loadSource`, and `onSave` are
JavaScript properties. Simple values such as `preview`, `theme`, `save-status`,
and `wasm-url` are attributes or properties.

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

- DOCX, HTML, plain text, and Markdown
- XLSX, XLS, ODS, and CSV
- PPTX
- PDF

Native exports are DOCX, XLSX, PPTX, and PDF. `downloadArtifact` starts a browser
download; `createArtifactBlob` returns a Blob for application-managed storage.

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

The published package contains a 4.6 MB PDFium WebAssembly file and the browser
build of PptxGenJS. Configure the CDN or static server to return `.wasm` as
`application/wasm`.

## Development

Requirements:

- Node.js 20 or newer
- Bun 1.3 or newer
- Rust 1.85 or newer

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

`bun run playground` starts an interactive React host that can create and import
all supported editor types. Production output is generated in `dist/`.

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
