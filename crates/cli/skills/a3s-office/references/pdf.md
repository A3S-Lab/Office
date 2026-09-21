# PDF editor reference

PDF is a browser viewer/editor boundary backed by PDFium. It owns page
navigation, thumbnails, search, zoom, annotations, page organization, save,
and download; it is not a native OOXML document and it has no file-level
find/replace occurrence. For durable collaboration replicas, locate editable
form values and FreeText annotation contents with `collab find` /
`office_collaboration_find` before mutating. Each hit is one match in form
collection order, then FreeText annotation order:

- form value → `fieldId` + `indexUtf16` → `pdf-set-form-value` with that
  `search` text, `indexUtf16`, and `expectedValue` equal to the current field
  value. `value` replaces only the matched span and fails closed if the span
  or `expectedValue` drifted. Omit `search` and `indexUtf16` to set the whole
  field value only when `expectedValue` still matches; an absent field uses
  an empty `expectedValue`. A mismatch writes nothing.
- FreeText (`type` 3) → `annotationId`, `pageIndex`, `annotationType`,
  `indexUtf16` → `pdf-update-annotation` with that `search` text and
  `indexUtf16`. `nextAnnotation.contents` replaces only the matched span and
  fails closed if the span drifted. Omit `search` and `indexUtf16` to set the
  whole FreeText contents.

Highlight and other markup types are not searchable via collab find. Do not
invent PDF body `replace-text`. Use the source-checkout UI operator for
user-level behavior:

```bash
bun run office:ops -- capabilities pdf --json
bun run office:ops -- check pdf --json
bun run office:ops -- visual pdf --project desktop-1280 --project compact-768
```

The focused A3S Test suite is `tests/e2e/pdf-page-organization.acl`; the
supplemental visual contracts are `visual-tests/pdf-page-organization.functional.spec.ts`
and `visual-tests/pdf-menu.functional.spec.ts`. Prefer semantic targets:

- `role=button|组织 PDF 页面` for the page organizer;
- `role=button|更多 PDF 工具` for compact overflow actions;
- `role=searchbox|在 PDF 中搜索` for search;
- `role=button|下载 PDF` for export;
- `css=.work-pdf-thumbnail-rail` and `css=.work-pdf-embed` for bounded page
  and viewer assertions.

The page organizer is transactional from the user's point of view. After a
move, delete, split, extract, or merge action, observe the thumbnail count,
selected page, focus return, and the viewer's ready state before continuing.
On phone widths, open the page drawer through its semantic toggle and assert
that the viewer is inert while the drawer owns focus.

Use the PDF collaboration mutation route for durable annotations and event
exchange. A browser download or semantic preview does not prove that a saved
PDF has the intended bytes; use the focused A3S Test evidence plus an explicit
artifact readback when the task depends on export content.
