# A3S Office

<p align="center">
  <strong>Embeddable Browser Office Editors for A3S</strong>
</p>

<p align="center">
  <em>Headless editing architecture, product-ready Office surfaces, and typed host integration for React, Vue, and Web Components</em>
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

A3S Office combines a headless, extension-driven editing architecture with
complete product surfaces. Headless does not mean that the package has no UI:
it means the editing models, commands, and host contracts do not prescribe the
surrounding product shell or backend. Applications can use the included
product-ready editors, integrate them into existing workspace chrome, and
connect them to product infrastructure without forking the editing engines.

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
  pageColor: '#ffffff',
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
- **Headless, Extension-Driven Architecture**: Keep editing models and typed
  commands independent from the host shell; extend Document and Markdown
  through public TipTap Extensions
- **Framework Entry Points**: Use React components, Vue 3 adapters, custom
  elements, or the framework-neutral Core API
- **Controlled Content**: Persist typed content models in the host and receive
  every user change through explicit callbacks or events
- **Office File Workflows**: Import and export DOCX, XLSX, PPTX, Markdown, PDF,
  HTML, text, ODS, XLS, and CSV where supported
- **Document Review**: Add anchored comments, replies, resolved state, tracked
  changes, citations, notes, captions, and cross-references
- **Document Interaction Surfaces**: Keep page settings, sources, revisions,
  comments, anchored controls, and modal decisions in responsive,
  keyboard-accessible panes, rails, popovers, and dialogs
- **Shared Office Interaction System**: Reuse the same tokens, ribbon patterns,
  task panes, dialogs, popovers, focus rules, and responsive behavior across
  every editor surface
- **Browser-Native Kernels**: Combine TipTap, editor-specific scene models,
  cancellable Workers, Rust WebAssembly layout and Spreadsheet calculation,
  and PDFium without a remote rendering service
- **Typed Commands**: Dispatch editor actions through typed controllers rather
  than interpreting visible labels or DOM text
- **Host-Owned Selection Menus**: Replace the Document selection context menu
  with typed host actions that receive the selected fragment, nearby text, the
  complete document, and conflict-aware editing commands
- **Long-Document Navigation**: Browse a live Word-style heading outline,
  filter it without changing the document, collapse sections, and jump with
  mouse or keyboard while the current heading follows the editor selection
- **Lazy Editor Engines**: Load only the selected editor and its large runtime
  assets
- **Native Automation**: Inspect and modify Office packages through the Rust
  CLI, JSON output, MCP, or the Office Skill

### Editor matrix

| Editor | Editing engine | Main capabilities | Native files |
| --- | --- | --- | --- |
| Document | TipTap/ProseMirror + Worker/Rust-WASM layout | Sections, page layout, clean page margins with double-click and Insert-ribbon header/footer editing, live heading navigation, responsive style and list galleries, bullet and numbering formats, restart/continue/start controls, host-defined selection menus with full document context, rich-text formatting with editable superscript and subscript, size-aware table insertion, Design/Layout table tabs with styles, fill, borders, alignment and structural editing, images, comments, tracked changes, citations, notes, captions, references | DOCX import/export, PDF export |
| Markdown | TipTap + GFM source model | Source and preview split view, coalesced preview updates, synchronized scrolling, task lists, tables, links, images, and code | MD import/export |
| Spreadsheet | Fortune Sheet + persistent Worker/Rust-WASM calculation sessions | Multiple sheets, operation-driven cell patches, bounded scalar formulas, incremental dirty dependency graphs, cross-sheet dependencies, formatting, charts, validation, protection, comments, print settings | XLSX/XLS/ODS/CSV import, XLSX/PDF export |
| Presentation | Typed multi-selection scene graph + on-demand TipTap text editing + Worker/Rust-WASM geometry | Slides, layouts, shapes, images, tables, charts, comments, transitions, presenter view, object/content mode separation, persistent nested browser groups, native PPTX group export, collective move/scale and keyboard commands, selection alignment/distribution, snapped move/resize previews, alignment guides, and virtualized thumbnail scenes | PPTX import/export, PDF export |
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

## Architecture

A3S Office places headless editing engines inside complete, reusable product
surfaces. In this project, **headless** means that editing behavior and host
contracts are independent from the surrounding application. It does not mean
that adopters must rebuild toolbars, dialogs, or document interactions.

### Design characteristics and advantages

The architecture shares contracts and interaction primitives where consistency
matters, while preserving the native model of each file format:

| Design characteristic                 | How it works                                                                                                                                                                           | Product and engineering advantage                                                                                            |
| ------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| Headless product boundary             | Editors do not own the host shell, storage, identity, collaboration, authorization, or AI provider; the package still includes complete editor UI                                      | Embed Office editing in an existing product without adopting an A3S backend or rebuilding standard interactions              |
| Format-native canonical state         | Document uses a TipTap/ProseMirror tree, Markdown keeps GFM source, Spreadsheet owns a workbook, Presentation owns a typed scene graph, and PDF delegates document semantics to PDFium | Preserve format-specific behavior instead of forcing every editor through a lowest-common-denominator model                  |
| Typed extension and command contracts | Document and Markdown accept public TipTap Extensions; the other editors expose typed command runtimes and stable host callbacks                                                       | Add shortcuts, plugins, file actions, and AI workflows without matching visible labels, simulating clicks, or forking source |
| Controlled host integration           | The host supplies content and receives typed changes; file operations and service requests use explicit callbacks                                                                      | Connect autosave, versions, permissions, collaboration, and external state replacement predictably                           |
| Shared Office interaction system      | All surfaces use the same design tokens, shell patterns, ribbons, dialogs, popovers, task panes, focus rules, and responsive contracts                                                 | Give users a consistent Office experience while allowing the host to theme and compose its workspace                         |
| Worker/WASM compute boundary          | Revisioned and cancellable Workers plus Rust WebAssembly handle bounded layout, formula, and geometry work                                                                             | Keep expensive computation deterministic and away from the primary interaction path without requiring a rendering server     |
| Framework convergence                 | React, Vue, and Web Components use the same editor engine; the Core entry point owns framework-neutral models and file workflows                                                       | Maintain one behavior and compatibility contract across frontend stacks                                                      |
| Browser and native execution planes   | Browser components handle interactive editing; the Rust CLI, MCP server, and Office Skill handle deterministic file automation                                                         | Support end-user editing, agents, and CI workflows without exposing filesystem concerns to the browser bundle                |
| Explicit fidelity boundary            | Compatibility reports and fixture gates make preservation, normalization, and rejection decisions visible                                                                              | Reduce silent file damage and make Microsoft Office/WPS interoperability measurable                                          |

### Layer model

```text
                          Host product
       persistence · identity · collaboration · authorization · AI
                               │
             ┌─────────────────┴──────────────────┐
             │                                    │
     Browser editing plane                Native automation plane
 React · Vue · Web Components             CLI · MCP · Office Skill
 framework-neutral Core API                       │
             │                              native Rust core
 shared Office shell + design system              │
             │                                    │
 controlled editor surfaces                       │
             │                                    │
 typed commands + Extensions                      │
             │                                    │
 ┌───────────┼──────────────┐                     │
TipTap    Workbook/scene   PDFium                  │
             │                                    │
 Workers + Rust WebAssembly                       │
             └─────────────────┬──────────────────┘
                               │
              DOCX · Markdown · XLSX · PPTX · PDF
```

### Engine ownership

Each editor owns one canonical state model and delegates only specialized work:

| Editor       | Canonical state                         | Specialized boundary                                                    | Architectural benefit                                                                 |
| ------------ | --------------------------------------- | ----------------------------------------------------------------------- | ------------------------------------------------------------------------------------- |
| Document     | One logical TipTap/ProseMirror document | Non-history pagination decorations and Worker/Rust-WASM layout          | Page calculation does not create duplicate editable trees or pollute undo history     |
| Markdown     | GFM source                              | Coalesced visual-tree rebuilds and synchronized preview                 | The saved value stays portable while source and visual workflows remain available     |
| Spreadsheet  | Workbook model                          | Persistent, revisioned Worker/Rust-WASM formula session                 | Incremental recalculation stays off the main interaction path                         |
| Presentation | Typed slide scene graph                 | On-demand TipTap text editing and cancellable Worker/Rust-WASM geometry | Object operations remain structured while rich text and alignment use focused engines |
| PDF          | PDFium document model                   | WebAssembly rendering and document commands                             | PDF semantics remain with a dedicated PDF engine instead of an HTML approximation     |

Public framework adapters converge on the same React editor engine, while the
Core entry point owns framework-neutral models and file workflows. The native
Rust core and CLI remain a separate automation boundary, so browser bundles do
not inherit filesystem or OOXML package state.

See [Browser editor architecture](docs/browser-editor-architecture.md) for
engine protocols, Worker/WASM ownership, fallbacks, and performance gates. See
[Editor quality roadmap](docs/editor-quality-roadmap.md) for product depth,
compatibility evidence, and release criteria.

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

See the online [integration guide](https://a3s-lab.github.io/Office/#guide) for
copyable component examples, Office CLI setup, and the optional Codex Skill.

## Editor Model

### Controlled state

The host owns the current content value. Editors do not select a database,
collaboration backend, identity provider, or model provider. File actions are
also optional host commands; do not duplicate them when the surrounding
product already provides open, save, or export controls.

`DocumentContent` stores compatibility HTML and may carry a structured,
versioned model. Persist the complete emitted value. Directly replacing HTML
invalidates an older model safely. Page layout fields include `pageSize`,
`pageColor`, orientation, margins, columns, and page chrome; page color is
preserved through DOCX import/export and browser PDF export.

### Selection context menus

`DocumentEditor.getSelectionMenuItems` lets the host completely replace the
selected-text context menu. A menu factory receives an immutable snapshot with
the selected text and structured fragment, up to 2,000 characters on either
side, the current `DocumentContent`, the complete plain text, and synchronized
HTML. A selected action receives the same snapshot plus conflict-aware
`replaceText`, `insertBefore`, `insertAfter`, and `copyText` commands.

```tsx
import type { GetDocumentSelectionMenuItems } from '@a3s-lab/office/core';

const getSelectionMenuItems: GetDocumentSelectionMenuItems = () => [
  {
    id: 'polish',
    label: 'Polish',
    icon: 'wand',
    onSelect: async (context) => {
      const response = await llm.rewrite({
        task: 'polish',
        selection: context.selection.text,
        before: context.selection.beforeText,
        after: context.selection.afterText,
        document: context.document.content,
        documentText: context.document.text,
      });
      const result = context.commands.replaceText(response.text);
      if (!result.applied) handleSelectionConflict(result.reason);
    },
  },
];

<DocumentEditor
  content={content}
  getSelectionMenuItems={getSelectionMenuItems}
  onChange={setContent}
/>;
```

Return or await every asynchronous action. While that Promise is pending, the
editor maps the original range through unrelated document transactions. If the
selected text itself changes, editing commands return `stale-selection`
instead of modifying the wrong range. Each successful command emits one
controlled update and one undo record, and text insertion or replacement
honors tracked-changes mode. Vue uses `:get-selection-menu-items`; custom
elements use the `.getSelectionMenuItems` property.

### Extensions

`DocumentEditor` and `MarkdownEditor` accept TipTap `extensions`. Additional
extensions are appended to the built-in schema before the editor mounts:

```tsx
import { Extension } from '@tiptap/core';
import type { DocumentContent } from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';

const hostShortcuts = Extension.create({
  name: 'hostShortcuts',
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-s': () => {
        window.dispatchEvent(new Event('office-save-request'));
        return true;
      },
    };
  },
});

<DocumentEditor
  content={content satisfies DocumentContent}
  extensions={[hostShortcuts]}
  onChange={setContent}
/>;
```

Keep the extension array referentially stable. Every extension needs a unique
`name`; a duplicate built-in or host extension name throws instead of silently
overriding editor behavior. A custom Document Node or Mark also needs a DOCX
import/export strategy. A custom Markdown Node or Mark needs matching Markdown
parse and serialize rules. Shortcut, storage, and ProseMirror Plugin extensions
that do not change persisted structure are the safest starting point.

Vue uses the `:extensions` prop. Custom elements use the `.extensions` property
because Extension instances cannot be represented as HTML attributes.
Spreadsheet, Presentation, and PDF do not yet expose their internal command
contexts as a public extension API. Use the stable host ports instead:
`fileActions`, `onAgentRequest`, `onStartSlideshow`, `loadSource`, and `onSave`.
Do not import implementation details from `@a3s-lab/office/internal`.

### Layout and runtime assets

Document pagination and Spreadsheet scalar calculation run in a dedicated
Worker backed by `office-kernel.wasm`. Spreadsheet jobs use sparse populated
cells and one persistent calculation session per editor. The first request
replaces the session workbook; later Fortune operations project only their
changed coordinates into bounded cell patches and recalculate only the dirty
dependency subgraph. Structural operations fall back to a workbook
replacement. Manual mode retains pending cell changes until explicit
recalculation. Formula-cache writes join the source edit's undo step instead of
creating a stale intermediate history state. External controlled values
remount the Fortune surface because Fortune does not apply later `data` props,
while values emitted by the live surface keep the current selection and mount.
Superseded requests retain revision order inside the Worker, known grouped
formulas are refreshed before their dependents, and unsupported dependencies
enter an ordered cell-scoped compatibility pass. The package includes
deterministic Latin, Simplified Chinese, Arabic, and Hebrew layout fonts.
Presentation alignment and snapped object transforms use the same kernel.
Pointer movement updates a frame-coalesced preview; releasing the pointer emits
one controlled value and therefore one undo step. Presentation group paths are
stored outermost-first in the scene model, selected as atomic logical units,
and remapped when objects, slides, or layouts are copied. PPTX export writes
those paths as nested native group nodes for slide, layout, and master-derived
objects. Generated groups use identity child coordinates computed from emitted
OOXML geometry, so non-rotated and non-reflected group transforms reimport with
their hierarchy and visual scale intact. PDF rendering uses `pdfium.wasm`,
while presentation export loads the browser PptxGenJS runtime only when needed.

Applications serving package assets from a separate CDN can pass explicit
`kernelWasmUrl`, `layoutFonts`, `wasmUrl`, or `pptxRuntimeUrl` values. Static
servers must return WebAssembly as `application/wasm` and allow fonts through
the required CORS policy.

`SpreadsheetEditor` accepts `kernelWasmUrl` in React and Vue. Its custom
element uses the `kernel-wasm-url` attribute:

```html
<a3s-spreadsheet-editor
  kernel-wasm-url="/office-assets/office-kernel.wasm"
></a3s-spreadsheet-editor>
```

### Fidelity boundary

A3S Office aims for predictable browser editing and native file preservation;
it does not claim pixel parity with every Microsoft Office or WPS feature.
Unsupported OOXML semantics, arbitrary floating-object layout, complete font
substitution, modern threaded comments, Spreadsheet arrays, spills, structured
references, external-workbook refresh, kernel-owned number formatting and print
pagination, the A3S-owned virtual grid, moving sparse projection work off the
main thread, arbitrary rotated or reflected PPTX group transforms, and the
remaining presentation scene features stay explicit fidelity gates. Rust/WASM
is the canonical Spreadsheet calculation path; if Worker or WebAssembly
loading fails, the Fortune-based JavaScript fallback keeps editing available
but may follow Fortune coercion and eager-branch semantics on formulas outside
the shared parity fixtures.

See [Browser editor architecture](docs/browser-editor-architecture.md) for
engine ownership, Worker/WASM boundaries, delivery stages, and performance
gates. See [Editor quality roadmap](docs/editor-quality-roadmap.md) for the
depth-first product priorities, release evidence, and documentation contract.

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
- [Integration guide](https://a3s-lab.github.io/Office/#guide)

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
