# Browser editor architecture

A3S Office uses an editor-specific interaction surface above a shared,
browser-safe Office kernel. TipTap is the right document engine for flowing
rich text, but it is not the canonical interaction model for a spreadsheet
grid, slide canvas, or PDF page.

## Product architecture

| Product | Interaction surface | TipTap responsibility | Kernel responsibility |
| --- | --- | --- | --- |
| Document | TipTap and ProseMirror | Complete logical document, selection, commands, history, comments, and collaboration boundary | Page layout, style and font resolution, OOXML semantics, and serialization |
| Spreadsheet | Virtualized grid and canvas | Rich text inside a cell or floating text object | Formula calculation, workbook semantics, print layout, OOXML, and serialization |
| Presentation | Scene graph and slide canvas | Text inside an individual text box | Masters, layouts, themes, object geometry, OOXML, and serialization |
| PDF | PDFium page surface | None | PDF parsing, rendering, annotation serialization, and document save |

This separation prevents an editor framework from becoming a false abstraction
over products with different selection, layout, and performance requirements.

## Document editing flow

```text
Controlled WorkDocumentContent
             |
      TipTap / ProseMirror
      one logical document
             |
   browser block measurement
             |
       dedicated Worker
             |
   Rust WebAssembly layout kernel
             |
 versioned page and break result
             |
 ProseMirror widget decorations
  no document or history mutation
```

The document remains one logical ProseMirror tree. Automatic page boundaries
are widget decorations, not page nodes. Reflow therefore does not split the
content model or corrupt undo, selection mapping, copy and paste, or a future
collaboration protocol.

The first implemented browser-kernel slice paginates measured block boundaries.
It supports explicit page and section breaks, keep-together blocks,
keep-with-next headings, cancellation, stale-revision rejection, and automatic
reflow after editing or element resizing. Page view uses the result for visual
paper gaps and status-bar page counts. Web view removes automatic pagination.

This slice is not yet a Microsoft Word or WPS line-layout fidelity claim. Exact
line breaking, widow and orphan control, floating objects, footnote balancing,
multi-column flow, font substitution, and mixed-size sections require the
later layout stages below.

## WebAssembly boundary

`crates/web-kernel` is deliberately independent of the DOM, filesystem,
network, and an async runtime. It accepts bounded JSON requests through a
small raw WebAssembly ABI:

- `office_kernel_abi_version`
- `office_kernel_alloc` and `office_kernel_dealloc`
- `office_kernel_layout`
- `office_kernel_result_pointer` and `office_kernel_result_length`

The protocol is versioned and carries request and document revisions. The
Worker ignores cancelled requests, and the React integration ignores stale
results. A matching JavaScript implementation preserves editing if Worker or
WebAssembly loading is unavailable.

The npm package emits `office-kernel.worker.js` and `office-kernel.wasm` beside
the public JavaScript entries. Hosts may override the WASM URL when assets are
served from a CDN or a nonstandard path.

## Migration boundary

The current public `WorkDocumentContent` remains a controlled HTML-backed model,
and the existing DOCX import path converts through Mammoth. The new layout
kernel does not silently change that persistence contract.

Migration to browser-native OOXML is staged:

1. Keep editing behavior stable while moving pagination off the React thread.
2. Add line boxes, font metrics, widow and orphan rules, notes, fields, and
   floating-object anchors to the layout protocol.
3. Compile the A3S OOXML, relationship, style, numbering, theme, and package
   layers for the browser with explicit memory and archive limits.
4. Introduce a versioned structured document model and loss-preserving OOXML
   round trips without making HTML the only source of truth.
5. Share style, theme, formula, geometry, and serialization primitives across
   Document, Spreadsheet, and Presentation while retaining their distinct
   editing surfaces.

Each stage needs compatibility fixtures, deterministic layout goldens, large
document performance budgets, and real Microsoft Office and WPS
interoperability evidence before it can be described as fidelity-complete.

## Performance and safety rules

- Editing and selection stay on the main thread; parsing and layout do not.
- One active layout request exists per editor. A newer revision cancels the
  previous request.
- Requests reject invalid dimensions and more than 10,000 layout blocks.
- WebAssembly performs no network or filesystem access.
- Worker failure is recoverable and does not block typing.
- Automatic pagination transactions are excluded from editor history.
- Multi-column and mixed-page-layout documents currently retain explicit-break
  behavior until their layout protocols are implemented.

## Verification

The browser kernel is covered at four boundaries:

1. Rust unit tests for deterministic pagination and validation.
2. JavaScript fallback tests for protocol parity and no-Worker operation.
3. A raw generated-WASM ABI smoke test.
4. Browser checks for real Worker/WASM loading, page-view reflow, web-view
   clearing, page counts, and undo behavior.

Run the focused kernel checks with:

```bash
bun run kernel:test
cargo test -p a3s-office-web-kernel
```
