# Real-time collaboration roadmap

This roadmap defines one versioned collaboration protocol for every A3S Office
editor and for native CLI/coding-agent clients. It complements `ROADMAP.md`,
which tracks editor capability parity with WPS.

## Design invariants

- Yjs is the browser protocol implementation. Yrs is the native Rust
  implementation. Both exchange standard Yjs v1 updates and state vectors.
- The host owns transport, authentication, authorization, persistence, rooms,
  and retention. Office consumes a synchronized document and never silently
  creates a network provider.
- Every shared document is bound to a version, artifact ID, and artifact kind.
  A client rejects mismatches before it edits canonical state.
- Exactly one bootstrap authority seeds a new artifact. Concurrent independent
  seeds are detected as ambiguous instead of being resolved with
  last-write-wins.
- Canonical content, durable review data, and ephemeral awareness use separate
  roots. Awareness is never persisted as document content.
- Human and agent transactions carry typed origins. Undo managers track only
  the local binding origin, so one participant cannot undo another
  participant's operation.
- `view`, `comment`, `suggest`, and `edit` are explicit session modes. A format
  enables a mode only after its durable model exists; until then, unsupported
  non-edit modes remain read-only.
- Editors never collaborate by repeatedly replacing one serialized OOXML file
  or one universal JSON blob. Each format gets a typed, conflict-local model.

Session modes are client-side capability guards, not a security boundary. The
host and synchronization service must enforce authorization for received and
persisted updates.

## Shared schema

Protocol v1 reserves the `a3s.office` namespace.

| Root | Y type | Purpose |
| --- | --- | --- |
| `metadata` | `Y.Map` | Protocol version, artifact identity, kind, and initialized state. |
| `bootstrap.initializers` | `Y.Array` | Detects more than one independent initial seed. |
| `markdown.source` | `Y.Text` | Canonical Markdown source. |
| `document.content` | `Y.XmlFragment` | ProseMirror document, including section layout, comment anchors, and tracked-change marks. |
| `document.options` | `Y.Map` | Document-level editable options such as page color and tracking mode. |
| `document.comments` / `document.comment-order` | `Y.Map` / `Y.Array` | ID-keyed durable comment/reply records and deterministic presentation order. |
| `document.bibliography*` | Typed maps/arrays | Bibliography settings and ID-keyed citation sources. |
| `spreadsheet.*` | Typed maps/arrays | Sheet order, field-addressed sparse cells, formulas, styles, objects, names, and print state. |
| `presentation.*` | Typed maps/arrays | Slide/master/layout order, ID-keyed scene objects, notes, transitions, and comments. |
| `pdf.*` | Typed maps/arrays | Planned source identity, annotations, forms, and reviewed page operations. |

Schema additions must be backward readable within protocol v1. A breaking root
or meaning change requires a new protocol version and an explicit migration.

## Delivery phases

### Phase 1: protocol core and Markdown

Status: implemented in the browser library.

- Versioned sessions, metadata validation, actor-aware origins, edit-mode
  enforcement, and host-owned `Y.Doc` lifecycle.
- Explicit single-authority initialization with concurrent-bootstrap
  detection.
- Canonical `Y.Text` Markdown binding with bounded UTF-16-safe replacements.
- Remote update projection into source and visual panes.
- Per-client undo/redo that preserves remote operations.
- React, Vue, and Web Component session plumbing.
- Convergence, permission, StrictMode, framework-parity, and ownership tests.

Remaining before this phase is production complete:

- Provider integration examples and reconnect/offline browser tests.
- Awareness-backed participant state and source selections.
- Durable update persistence and compaction reference implementation.
- Cross-language Yjs/Yrs fixture tests.

### Phase 2: Document

Status: browser collaboration foundation implemented; native parity and the
remaining review/presence matrix are pending.

- TipTap is bound to `document.content` through
  `@tiptap/extension-collaboration`; StarterKit undo/redo is disabled and the
  collaboration binding owns local history.
- Section/page OOXML state, comment anchors, and tracked-change marks converge
  in the ProseMirror fragment. Document-level page color, tracking mode,
  comment threads/replies, and bibliography sources use conflict-local typed
  roots instead of a serialized document blob.
- React, Vue, and Web Component adapters accept the same initialized session.
- Append-only record claims deduplicate identical offline retries and fail
  closed when disconnected clients assign different records to one stable ID.
- Browser tests cover initialization, bootstrap races, sidecar convergence,
  stable-ID retries/collisions, stale-snapshot delete conflicts, local-only
  undo, permission guards, real TipTap peer edits, remote projection,
  StrictMode, and framework parity.

Remaining:

- Add awareness-backed cursors and selections without persisting presence.
- Add suggestion-only authorization and durable accept/reject decision audit
  records; `comment` and `suggest` remain read-only until those models exist.
- Prove concurrent text, table, list, section, comment, and revision workflows,
  plus DOCX import/export after merged edits.
- Add cross-language Yjs/Yrs fixtures and native client convergence.

Exit criterion: two browsers and one native client converge on document
content and review state; local undo never removes a remote change; a DOCX
round trip preserves supported and untouched unsupported OOXML.

### Phase 3: Presentation

Status: browser collaboration foundation implemented; rich-text CRDT and PPTX
round-trip concurrency coverage are pending.

- Stable slide/master/layout order arrays and ID-keyed record maps avoid a
  serialized whole-deck root. Scene elements and slide comments are ID-keyed
  child records.
- Snapshot writes are translated into `previous -> next` record patches, so an
  unrelated stale snapshot does not remove remotely added slides or objects.
- Typed validation rejects duplicate identities and invalid references before
  bootstrap; concurrent independent bootstrap remains fail-closed.
- Per-binding undo/redo tracks only local transactions. React, Vue, and Web
  Component surfaces project remote snapshots and keep non-edit modes
  read-only.
- Tests cover bootstrap, identity validation, separate-record convergence,
  stale snapshots, concurrent comments, local-only undo, permissions, mounted
  projection, and framework adapters.

Remaining:

- Bind editable scene text to collaborative XML fragments instead of scalar
  text/run replacement.
- Add creation claims and explicit conflict handling for concurrent same-ID
  object creation, delete-vs-edit, z-order, grouping, and theme changes.
- Keep derived thumbnails and layout measurements local and prove this in UI
  tests.
- Add offline/reordered-update property tests plus merged PPTX export/reopen.

Exit criterion: concurrent object transforms, rich-text edits, slide reorder,
comments, and delete-vs-edit cases converge and survive PPTX round trips.

### Phase 4: Spreadsheet

Status: browser collaboration foundation implemented; structural operations,
native recalculation parity, and XLSX concurrency coverage are pending.

- Stable sheet/named-range order arrays, ID-keyed records, and append-only
  creation claims avoid a serialized workbook or dense worksheet root.
- Sparse cell presence and recursively field-addressed values let formula,
  cached value, number format, style, hyperlink, and note edits converge even
  when two disconnected clients first populate the same blank cell.
- Sheet configuration and formula metadata are recursively field-addressed,
  so independent merge geometry, row/column sizing, protection, source-formula,
  and dynamic-array metadata leaves do not replace their containing object.
- Images, charts, pivot tables, defined names, calculation settings, print
  areas/titles, page breaks, and page setup have conflict-local typed roots.
- Snapshot writes are translated into `previous -> next` patches. Unrelated
  stale snapshots preserve remote cells and records; a delete that would
  discard a remote edit fails closed and asks the caller to refresh.
- Selection, active-sheet status, zoom, calculation chains, dynamic-array
  caches, and derived chart preview images are excluded from canonical state.
- The public Core binding supports typed origins, edit permissions,
  subscriptions, and per-client undo/redo. Tests cover bootstrap races,
  malformed roots, identity collisions, stale snapshots, OOXML field
  convergence, and local-only undo.
- React, Vue, and Web Component editors project remote workbook snapshots and
  route canonical edits plus local-only undo/redo through the same binding.
  Active-sheet status, selections, zoom, and Fortune Sheet focus remain local
  view state and are reapplied after remote canonical updates.

Remaining:

- Translate one user gesture into one transaction; batch paste/fill/sort/filter
  without emitting a document-sized replacement.
- Add stable structural operations and reference transforms for row/column
  insertion/deletion, merged ranges, tables, sort, and named references.
- Treat calculated values as derived once browser and native recalculation are
  deterministic and source formulas remain the only canonical formula state.
- Add offline/reordered-update property tests plus merged XLSX export/reopen.

Exit criterion: multi-client structural edits converge, recalculation agrees
across browser and native clients, and XLSX round trips preserve supported plus
untouched unsupported package state.

### Phase 5: PDF

- Bind a shared artifact to an immutable source fingerprint; reject updates for
  another byte source even when page counts happen to match.
- Collaborate on annotations, form values, signatures, redaction proposals,
  approvals, and reviewed page operations as typed records.
- Require explicit review for destructive redaction or page mutations.
- Keep rendered bitmaps, search indexes, and viewport state local.

Exit criterion: annotations and forms converge and reopen from a saved PDF;
destructive actions remain attributable, reviewable, and non-retryable.

### Phase 6: CLI, MCP, and coding agents

- Add Yrs with Yjs-compatible client IDs and standard sync protocol support to
  the native Office workspace.
- Persist append-only updates plus periodic snapshots/state vectors with
  bounded compaction and crash recovery.
- Expose typed commands to create, join, inspect, watch, synchronize, mutate,
  checkpoint, and leave sessions. Commands emit machine-readable progress and
  never require a TTY.
- Give every agent a stable actor ID and require an operation ID, expected
  artifact identity, permission mode, and optional state-vector precondition
  for mutations.
- Return structured conflict, stale-state, and permission errors. Ambiguous
  Office mutations are never automatically retried.
- Project the same commands into MCP and `a3s code` tools; stream remote changes
  into the agent event loop without polling the whole file.

Exit criterion: a browser, `office` CLI, MCP client, and `a3s code` agent edit
the same artifact concurrently, converge at the shared-model level, reconnect
after interruption, and preserve actor/operation attribution.

## Required test matrix

Every format phase adds tests for two concurrent browser clients, browser plus
Yrs, offline edits followed by reconnect, duplicate/reordered updates,
single-authority bootstrap, artifact and kind mismatch, local-only undo,
permission denial, host-owned lifecycle, schema migration, import after merge,
and export/reopen. Fuzz and property tests compare final state after different
valid update delivery orders.

## Non-goals

- Shipping a hosted collaboration backend inside the component package.
- Treating awareness as authorization or durable audit history.
- Hiding provider synchronization behind editor mount.
- Claiming comments or suggestions are supported by a format until its durable
  review schema and accept/reject semantics are implemented.
