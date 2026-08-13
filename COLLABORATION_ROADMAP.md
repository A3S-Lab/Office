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
| `spreadsheet.*` | Typed maps/arrays | Planned sheet order, sparse cells, styles, objects, and review state. |
| `presentation.*` | Typed maps/arrays | Planned slide order, scene objects, text, notes, and comments. |
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

- Use a stable slide-order array and one map per scene object; never replace the
  whole deck for a text or geometry edit.
- Use collaborative XML fragments for rich text and typed scalar fields for
  transforms, styles, links, notes, comments, and transitions.
- Define deterministic conflict rules for delete-vs-edit, z-order, grouping,
  theme changes, and object identity.
- Keep derived thumbnails and layout measurements local.

Exit criterion: concurrent object transforms, rich-text edits, slide reorder,
comments, and delete-vs-edit cases converge and survive PPTX round trips.

### Phase 4: Spreadsheet

- Model sheet order, identities, sparse cells, formulas, styles, merges,
  validations, comments, tables, charts, and names as typed shared structures.
- Translate one user gesture into one transaction; batch paste/fill/sort/filter
  without emitting a document-sized replacement.
- Treat calculated values, viewport state, and selection as derived or
  awareness data. Formula source remains canonical and recalculation is
  deterministic.
- Define structural conflict behavior for row/column insertion, deletion,
  merged ranges, table resize, sort, and named references.

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
