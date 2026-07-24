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
| Document | Rich text, sections, page layout, headers and footers, contextual image wrapping, comments, tracked changes, citations, notes, captions, cross-references, DOCX import/export, and PDF export |
| Markdown | TipTap editing with a source-and-preview split view by default, optional visual or source-only views, headings, lists, quotes, links, images, code, native Markdown import/export, and read-only preview |
| Spreadsheet | Multiple sheets, formulas, charts, conditional formatting, pivot tables, comments, validation, protection, print settings, XLS/XLSX/ODS/CSV import, XLSX export, and PDF export |
| Presentation | Slides, masters and layouts, rich text, shapes, images, tables, charts, notes, comments, transitions, presenter view, PPTX import/export, and PDF export |
| PDF | PDFium rendering, navigation, search, form filling, basic annotations, undo and redo, and saving an edited copy |

The editors run entirely in the browser. A server is optional and remains in
control of persistence, collaboration, authentication, and AI requests.

The browser architecture is editor-specific rather than forcing every product
through one abstraction. Documents use one logical TipTap/ProseMirror tree;
automatic page layout runs in a Worker backed by the Rust WebAssembly kernel
and returns non-history page decorations. Spreadsheet and Presentation keep
their grid and scene models, using TipTap only for rich text where appropriate.
Their ribbons dispatch typed commands instead of interpreting labels or
rendered text. Presentation slide-relative alignment also runs through the
Worker/Rust-WASM kernel. PDF rendering remains on PDFium and its A3S toolbar
calls typed PDFium capabilities directly. See
[Browser editor architecture](docs/browser-editor-architecture.md) for the
current fidelity boundary, implementation status, delivery stages, and
performance gates.

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

`DocumentContent` may carry a versioned `model` beside its required `html`
compatibility representation. `DocumentEditor` emits both after the first edit
and treats the structured model as authoritative while its HTML fingerprint
still matches. Persist the complete value received by `onChange`. Changing
`html` directly automatically invalidates an older model, so existing
HTML-only hosts remain safe during the migration.

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
  `DocumentLayoutFont`
- `defaultOfficeKernelWasmUrl`, `defaultDocumentLayoutFontUrl`,
  `defaultDocumentLatinLayoutFontUrl`,
  `defaultDocumentArabicLayoutFontUrl`,
  `defaultDocumentHebrewLayoutFontUrl`, `defaultDocumentLayoutFonts`, and
  `defaultPdfiumWasmUrl`

Set `preview` for a read-only representation. Use `theme="light"`,
`theme="dark"`, or `theme="system"`.

Each editor engine is an independent asynchronous chunk. Importing the React
entry does not download TipTap, Fortune Sheet, PDF, or Presentation code until
that editor is rendered. Applications can warm the selected engine from a
hover or keyboard-focus intent without mounting it:

```tsx
import { preloadOfficeEditor } from '@a3s-lab/office/react';

<button
  onPointerEnter={() =>
    void preloadOfficeEditor('document').catch(() => undefined)
  }
  onFocus={() =>
    void preloadOfficeEditor('document').catch(() => undefined)
  }
>
  New document
</button>;
```

The component displays the shared Office loading state if the chunk is not
already available.

### Document layout kernel

`DocumentEditor` resolves `office-kernel.wasm` and bundled Noto Sans, Noto Sans
Hans, Noto Naskh Arabic, and Noto Sans Hebrew regular faces beside `react.js`.
A dedicated Worker registers the raw font bytes with the Rust/WASM kernel,
resolves Unicode bidirectional levels, shapes each directional and font segment
with Rustybuzz, applies Unicode line-breaking rules across run boundaries, and
returns UTF-16 line offsets for TipTap. Each run carries an ordered, explicitly
registered CSS font stack plus its size, line height, letter spacing, ligature,
and kerning behavior. The kernel selects one face per grapheme, coalesces
adjacent graphemes that use the same face, and includes fallback faces in line
ascent and descent. Pagination consumes those deterministic line fragments
without putting font parsing or shaping on the React thread.

If assets are hosted elsewhere, pass an absolute `kernelWasmUrl` and matching
font source:

```tsx
<DocumentEditor
  content={content}
  kernelWasmUrl="https://cdn.example.com/a3s-office/office-kernel.wasm"
  layoutFonts={[
    {
      id: 'host-latin-regular',
      family: 'Host Latin',
      url: 'https://cdn.example.com/fonts/host-latin-regular.ttf',
      weight: 400,
      style: 'normal',
    },
    {
      id: 'host-cjk-regular',
      family: 'Host CJK',
      url: 'https://cdn.example.com/fonts/host-cjk-regular.otf',
      weight: 400,
      style: 'normal',
    },
  ]}
  onChange={setContent}
/>
```

Set
`--a3s-document-font-family: "Host Latin", "Host CJK", sans-serif`
on the Office container when replacing the bundled default. Every explicit
family before the generic family must match a loaded `layoutFonts` asset in the
same order; otherwise the editor deliberately retains browser line
measurement. The bundled default stack is `"A3S Office Noto Sans", "A3S Office
Noto Sans Hans", "A3S Office Noto Naskh Arabic", "A3S Office Noto Sans Hebrew",
sans-serif`.

The deterministic path accepts text-flow paragraphs with multiple visible runs
when every explicit CSS family before the generic family resolves to an exact
registered face. Up to eight ordered faces may participate in one run.
Different stacks, sizes, line heights, letter spacing, ligature settings, and
kerning settings may coexist in one paragraph. Mixed left-to-right and
right-to-left text is segmented by resolved Unicode bidi levels before
Rustybuzz shaping. Glyphs missing from the complete registered stack,
unregistered or browser-synthesized faces, unsupported OpenType features,
inline objects, soft hyphens, tab paragraphs containing resolved
right-to-left runs, and unsupported paragraph structures use the existing DOM
measurement path. This explicit fallback keeps editing and pagination working
without presenting estimated metrics as exact.

Page view includes horizontal and vertical rulers linked to the active section.
The horizontal ruler controls left and right margins, paragraph left, right,
first-line, and hanging indents, plus typed paragraph tab stops. The vertical
ruler controls top and bottom margins. Every handle supports pointer and
keyboard input, and narrow layouts hide the vertical ruler without removing
the equivalent Page Layout controls. Clicking the horizontal ruler adds a tab
stop; accessible handles move, change, or remove it. The Tab key inserts an
inline ProseMirror atom outside tables and lists; table-cell navigation and
list-item indentation/outdentation retain their native editor behavior.
Browser layout supports 48 px default stops, custom left, center, right, or
decimal alignment, and visible leaders in edit, preview, and PDF rendering,
including non-100% document zoom. Worker/Rust-WASM pagination uses the same
stop-selection policy and Rustybuzz-shaped following-segment widths for
structured left-to-right tabs, so a tabbed paragraph no longer falls back
merely because it contains a Tab.

Direct paragraph pagination can keep a paragraph together, keep it with the
following paragraph, begin it on a new page, or prevent a single visual line
at the top or bottom of a page. These paragraph properties survive DOCX export
and DOCX paragraph-formatting imports. Import resolves paragraph indents,
spacing, line rules, pagination, direction, and tab stops from document
defaults, default paragraph styles, `basedOn` style chains, and direct
paragraph overrides. Explicit left-to-right and right-to-left direction is
stored on semantic TipTap paragraph or list-item nodes, sent to the text-layout
kernel, and mapped to DOCX `w:bidi`; exported right-to-left text runs carry
`w:rtl`. The Home ribbon exposes direct paragraph-direction controls.
`w:tab` content becomes a structured inline node rather than collapsed
whitespace; export writes both paragraph `w:tabs` and inline tabs. The Page
Layout ribbon exposes compact paragraph-spacing and pagination controls without
expanding the page-settings panel. Full Word style resolution is also used for
imported run font families, sizes, text colors, and backgrounds across
paragraph and character styles. Imported theme fonts, theme colors, tint, and
shade are resolved from the DOCX theme part.

Document images are typed TipTap nodes rather than unstructured toolbar
effects. The contextual Picture ribbon edits inline, square, or top-and-bottom
layout, left/center/right alignment, text distance, alternative text, size, and
deletion on the selected node. Preview and editing use the same attributes.
DOCX export writes supported square and top-and-bottom layouts as `wp:anchor`
drawings; DOCX import restores those layouts, alignment, and wrap distance.
Paragraphs after a left- or right-aligned square image deliberately use browser
line-box measurement because their available width changes over the image
height. The Worker/Rust-WASM paginator still consumes those measured fragments,
reserves the image height, and reflows after image load or resize.

Page headers and footers use controlled, purpose-built TipTap surfaces instead
of detached `contenteditable` state or the deprecated browser
`document.execCommand` API. Their compact toolbars dispatch typed commands for
undo, redo, bold, italic, underline, paragraph alignment, color, links, and
bounded raster-image insertion. Active formatting is derived from the editor
selection, unsafe links are rejected before mutation, and all edits continue to
flow through the existing first/default/even-page model and DOCX relationship
export. Pagination carries typed physical-page, section-page, displayed-page,
and section metadata, then resolves the first/default/even variant for each
paper. Page gaps repeat the previous page footer and next page header without
duplicating editable document content. The first paper header and final paper
footer expose accessible edit targets; entering either target mounts the same
controlled TipTap surface directly in the margin, dims the document body, and
opens a contextual `Header and Footer` ribbon for formatting, navigation,
page-number visibility, and explicit exit. Escape or switching to web view
returns focus to the document body. The Page Layout panel remains the place to
edit every first/default/even variant rather than becoming a second content
source.

Ordered and bullet lists remain structured instead of being flattened into one
paragraph. Every item enters pagination independently, eligible item paragraphs
use the same Worker/Rust-WASM shaping path, and an automatic page gap is mapped
inside the correct list item. Native DOCX output writes one numbered paragraph
per item and preserves positive ordered-list starts, nesting levels, and
explicit RTL items. The same numbering path is used in table cells, document
notes, and page headers and footers. DOCX import restores positive list starts
plus decimal, lower- and upper-letter, and lower- and upper-Roman numbering
after Mammoth conversion. Reversed lists, per-item value overrides, arbitrary
Word numbering restarts, and loss-preserving custom numbering formats remain
fidelity gates.

Language-complete fallback and font substitution, bidirectional tab layout,
variable font axes, arbitrary floating-object offsets and layering, image
cropping and contour wrapping, arbitrary inline splitting inside table cells,
nested tables, and multi-column layout remain explicit fidelity gates rather
than implied Word or WPS compatibility.

Tables are measured as structured row flows instead of one keep-together
block. The Worker/WASM kernel keeps leading header rows with the first body
row, reserves their measured height on continuation pages, and returns row
breaks that map back to non-editable ProseMirror decorations. The page view
repeats the real rendered header columns without duplicating editable content.
Row-level `cantSplit` and `tblHeader` properties round-trip through DOCX and
are available from the contextual table controls. An ordinary row may split at
synchronized direct-cell block boundaries when its full height fits on a
continuation page. Per-cell decorations align the paper gap and repeated header
without changing the canonical table content. `cantSplit` rows, single long
paragraphs, and rows taller than a continuation page remain atomic and may
overflow.

After a ProseMirror transaction, DOM measurements before the earliest changed
position are reused. Layout requests carry separate document and layout
revisions, so resizing can supersede a layout without pretending the document
changed. Complete stable pages before the affected region are also reused.
The incremental planner rewinds across paragraph flows and `keepWithNext`
boundaries, sends only the safe suffix to the kernel with absolute page
indices, and merges a freshly calculated boundary with the retained prefix.
Runtime diagnostics report measured and reused blocks, reused pages, text
layout candidates, submitted and accepted text runs, and the number of blocks
submitted to the kernel.

The structured model currently covers the TipTap/ProseMirror logical document,
schema version, and monotonic revision. DOCX, HTML, and text imports create it
through the same extension schema used by `DocumentEditor`. DOCX export
materializes its compatibility HTML from the structured model before OOXML
generation; preview still consumes synchronized HTML. Loss-preserving OOXML
package state and WASM-owned serialization remain later fidelity gates.

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

### Spreadsheet and Presentation

Spreadsheet keeps the workbook grid as its canonical interaction model. Its
ribbon emits typed commands for selected-range formatting, merge state,
formula recalculation, gridlines, and zoom. A dedicated Worker/WASM dependency
graph and calculation engine is still a later fidelity gate.

Presentation keeps each slide as a scene graph and mounts TipTap only for the
selected text box. Its ribbon uses one typed command dispatcher for slide,
clipboard, element, comment, design, transition, view, and slideshow actions.
Slide-relative element alignment is revisioned, cancellable, and computed in
the shared Worker/Rust-WASM kernel with a JavaScript fallback.

`PresentationEditor` resolves `office-kernel.wasm` in the same way as
`DocumentEditor`. Hosts that serve package assets from another location can
pass `kernelWasmUrl` explicitly:

```tsx
<PresentationEditor
  content={content}
  kernelWasmUrl="https://cdn.example.com/a3s-office/office-kernel.wasm"
  onChange={setContent}
/>;
```

This geometry slice currently covers alignment to the slide only. Snapping,
alignment guides, grouping, connectors, text fitting, theme/master resolution,
and loss-preserving PPTX serialization remain explicit fidelity gates.

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
The visible toolbar belongs to A3S Office rather than the embedded PDF engine.
Navigation, zoom, search, basic annotation tools, history, and export call
typed plugin capabilities directly. The integration does not inspect the
rendered toolbar, infer commands from labels, or depend on private shadow-DOM
selectors.

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
kernel, the four default Latin, Simplified Chinese, Arabic, and Hebrew layout
faces, a 4.6 MB PDFium WebAssembly file, and the browser build of PptxGenJS.
These large assets remain editor-specific and lazy-loaded. Configure the CDN or
static server to return `.wasm` as `application/wasm` and font assets with an
appropriate font MIME type and CORS policy.

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
bun run playground:build
bun run playground:visual
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
`bun run playground:visual` checks committed Chromium screenshots and shared
shell geometry for all five editors at 1280 × 800 and 768 × 800. Install the
pinned browser once with `bunx playwright install chromium`; update baselines
only for an intentional UI change with `bun run playground:visual:update`.

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
