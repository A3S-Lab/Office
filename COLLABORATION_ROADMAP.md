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
  enables a mode only after its durable model exists. Document enables durable
  review-only `comment` plus actor-attributed `suggest`; unsupported
  mode/format pairs remain read-only.
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
| `document.change-decisions` / `document.change-decision-order` | `Y.Map` / `Y.Array` | Immutable accept/reject audit records and deterministic decision order. |
| `document.bibliography*` | Typed maps/arrays | Bibliography settings and ID-keyed citation sources. |
| `spreadsheet.*` | Typed maps/arrays | Sheet order, field-addressed sparse cells, formulas, styles, objects, names, and print state. |
| `presentation.*` | Typed maps/arrays | Slide/master/layout order, ID-keyed scene objects, notes, transitions, and comments. |
| `pdf.*` | Typed maps/arrays | Immutable source identity, annotations, forms, signatures, redaction proposals, reviewed page operations, and final decisions. |

Optional additive roots such as `pdf.source-identities` let newer clients
validate sessions created by newer clients while continuing to accept legacy
protocol-v1 documents that do not contain the root. Older clients ignore the
unknown root.

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
- A bounded, identity-bound host transport adapter exchanges standard Yjs v1
  state vectors and updates, preserves typed incremental origins, suppresses
  update echoes, and exposes explicit reconnect synchronization.
- A provider-owned Awareness controller publishes validated actor, mode,
  activity, and format-specific locations for all five artifact kinds without
  persisting presence in the shared document.
- A shared, accessible participant roster projects local and remote humans,
  agents, modes, activity, and typed location summaries into editable and
  preview chrome for all five editor kinds. Pairing fails closed across
  artifact, kind, Y.Doc client, actor, namespace, or mode boundaries.
- Every browser editor publishes its local ephemeral location and projects
  remote locations into its native editing surface. Document and Markdown use
  text selections/carets (with explicit Markdown source/visual coordinates),
  Spreadsheet uses Fortune Sheet cell presence, Presentation uses stable
  object frames, and PDF uses page/annotation markers. Remote roster rows
  navigate only on explicit activation; passive Awareness updates preserve
  local focus, selection, and viewport.
- React, Vue, and Web Component session and presence plumbing.
- A runnable A3S Boot reference backend now provides signed room tickets,
  Origin validation, permission enforcement, WebSocket rooms, durable Yrs
  persistence, ephemeral Awareness relay, reconnect repair, and service-owned
  durable update polling. Version-2 tickets bind actor display name and the
  durable store semantically authorizes Document comment- and suggestion-mode
  updates under the room lock before persistence or broadcast. Its browser
  adapter converts bounded base64 wire payloads back to the package's typed
  `Uint8Array` transport envelope and verifies the ready identity against the
  local session.
- Convergence, offline/reconnect, transport-boundary, presence, permission,
  StrictMode, framework-parity, ownership, backend persistence, room broadcast,
  ticket tamper, and read-only authorization tests.

The reference backend is complete for one service process. Multi-replica
deployments still require host-selected sticky room routing or shared
Redis/NATS fan-out plus a distributed writer/lock policy.

### Phase 2: Document

Status: browser collaboration foundation, participant roster, remote
selection/caret projection, participant navigation, durable review-only
comment and suggestion modes, immutable suggestion-decision audit, and native
text/options/bounded structural paragraph/comment mutations implemented; rich
native mutation parity remains pending.

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
- Native Rust, CLI, and MCP clients can replace an exact expected match count
  inside ProseMirror `Y.XmlText`, including UTF-16 astral characters, without
  replacing the shared HTML/XML tree. The replacement preserves the first
  deleted character's rich-text attributes and rotates the affected Word
  `textId` once. They can also insert a plain paragraph before or after a
  uniquely identified paragraph, heading, or document caption in a bounded
  top-level section, nested list item, table cell/header, or blockquote path,
  and delete a plain paragraph only when its complete text and current
  `textId` still match. Required container blocks and list-leading paragraphs
  are preserved. Any text or structural edit inside a table rotates every
  identified ancestor row's `rowTextId` in the same transaction. Page color
  and track-changes settings mutate their conflict-local `document.options`
  entries.
- Deterministic browser-to-Yrs and Yrs-to-browser fixtures prove that native
  text/options/paragraph updates remain readable by the Office TipTap schema
  after durable restart and replay, including concurrent native and browser
  paragraph insertion across nested lists and nested tables. Real CLI and MCP
  subprocess tests exercise the same bounded structural mutations.
- Document `comment` sessions now create selection-anchored threads, append
  replies, resolve or reopen threads, and delete only their actor's own comment
  or reply records while canonical text, structure, options, bibliography, and
  non-comment formatting stay read-only. Records persist by stable ID and
  immutable creation claim; removing an anchor through an authorized content
  edit retains a detached thread. Remote review changes do not enter local undo
  history.
- Authenticated Document `suggest` sessions create attributed insertion,
  deletion, and replacement proposals as `documentChange` marks. Their
  canonical projection is unchanged until an `edit` participant decides the
  proposal. Suggesters may withdraw their own insertion proposals, but cannot
  change canonical text, structure, non-suggestion formatting, options,
  comments, bibliography, another actor's proposal, or a deletion proposal's
  targeted canonical text.
- Accepting or rejecting a suggestion applies the tracked change and appends
  one immutable actor-attributed record to `document.change-decisions` plus its
  deterministic order root in the same transaction. Identical offline retries
  converge, a second conflicting final decision fails closed, old Documents
  without the additive roots remain readable, and a final decision clears
  local history that could otherwise resurrect the decided mark.
- Rust, CLI, MCP, and A3S Code expose `document-comment-create`,
  `document-comment-reply`, `document-comment-set-resolved`, and
  `document-comment-delete`. Projection schema v2 returns attributable threads,
  replies, anchor text, paragraph/text identities, browser-compatible UTF-16
  offsets, resolution, and detached state. Native and browser fixtures cover
  restart, reordered delivery, stale anchor guards, idempotent retries,
  ownership failures, and cross-language convergence.
- The A3S Boot backend permits `edit` content/review updates, semantically
  validated Document `comment` review-only updates, and semantically validated
  Document `suggest` proposal updates. Signed actor identity and display name
  must match every new comment, reply, or suggestion. Forged roots, canonical
  content, structure, options, non-suggestion formatting, authorship, order,
  claims, anchors, foreign deletion, or foreign suggestion changes fail before
  the durable store changes.

Remaining:

- Add closed typed native projection and mutation variants for creating
  Document suggestions and deciding them through CLI/MCP/A3S Code. The current
  typed native mutation set can synchronize the resulting Yjs state but does
  not originate these two workflows.
- Prove concurrent full-table, list-restructure, section, comment, and revision
  workflows, plus DOCX import/export after merged edits.
- Expand native convergence from bounded paragraph edits to complete table,
  list, and section operations, comments, and tracked revisions.

Exit criterion: two browsers and one native client converge on document
content and review state; local undo never removes a remote change; a DOCX
round trip preserves supported and untouched unsupported OOXML.

### Phase 3: Presentation

Status: browser collaboration foundation and native conflict-local
scene-element content and z-order mutations implemented; rich-text CRDT,
remaining structural native operations, and PPTX round-trip concurrency
coverage are pending.

- Stable slide/master/layout order arrays and ID-keyed record maps avoid a
  serialized whole-deck root. Scene elements and slide comments are ID-keyed
  child records.
- Append-only element claims make identical same-ID creation idempotent and
  reject different records assigned to one ID. Deletion writes a durable
  tombstone, removes the element from visible order, and permanently reserves
  the identity; element ID and type cannot drift.
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
- Rust, CLI, MCP, and A3S Code can create, optimistically update, move, or
  tombstone one element in a slide, master, or layout. Creation can insert
  after a stable active element. Update compares complete
  expected/current/next records and writes only changed top-level fields, so
  unrelated concurrent changes merge while a stale same-field edit fails
  atomically. Move compares stable observed and requested predecessor IDs,
  treats an already-satisfied destination as idempotent, and changes only the
  moved element's order entry; `null` means the first order position. Delete
  requires an exact complete-element match. Browser/Yrs fixtures, native
  restart, duplicate and reordered delivery, real CLI/MCP subprocesses, and
  the Playground cover update/create/move/delete interoperability.

Remaining:

- Bind editable scene text to collaborative XML fragments instead of scalar
  text/run replacement.
- Add structural native operations and explicit conflict handling for slide
  order, grouping, and theme changes; extend z-order beyond stable scene-element
  moves where a concrete browser workflow requires it.
- Keep derived thumbnails and layout measurements local and prove this in UI
  tests.
- Add offline/reordered-update property tests plus merged PPTX export/reopen.

Exit criterion: concurrent object transforms, rich-text edits, slide reorder,
comments, and delete-vs-edit cases converge and survive PPTX round trips.

### Phase 4: Spreadsheet

Status: browser collaboration foundation and native conflict-local cell
mutations implemented; structural operations, native recalculation parity, and
XLSX concurrency coverage are pending.

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
- Rust, CLI, MCP, and A3S Code can create or recursively patch one cell with
  `spreadsheet-set-cell`, or delete one exact complete cell with
  `spreadsheet-delete-cell`. Recursive optimistic guards merge unrelated JSON
  leaves and reject stale same-leaf edits before writing. Dense projections
  retain and safely extend row lengths, sparse projections remain `celldata`,
  and an empty sheet's first write stays sparse. Native restart, real CLI/MCP
  subprocesses, and browser Yjs duplicate/reordered delivery cover the same
  update contract.

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

Status: browser Core plus PDF viewer/framework projection, native annotation
and form-value mutation, and native append-only redaction/page-operation
proposal and final-decision mutations implemented; save/reopen, authenticated
assets, and full native parity are pending.

- Shared artifacts are bound to a lowercase SHA-256 fingerprint, exact byte
  length, and page count. Source bytes never enter Yjs, and a binding rejects
  another source even when its page count matches.
- Annotations, form values, signature placements, redaction proposals, page
  operations, and final review decisions use typed, ID-keyed, recursively
  field-addressed roots instead of one PDF or JSON blob.
- Annotation deletion uses an irreversible durable tombstone. Signature
  placements reference host-owned appearance assets; private/signature bytes
  are not synchronized as canonical content.
- Signature, redaction, page-operation, and decision records are append-only
  audit data. Redaction and destructive page operations require an explicit,
  attributable final review decision before a host applies them to bytes.
- Creation claims deduplicate identical offline retries and fail closed on
  concurrent same-ID reuse. Snapshot writes preflight stale conflicts before
  changing any shared root, and local undo is limited to annotations and form
  values.
- Core tests cover source mismatch, validation, bootstrap races, field-level
  convergence, delete-vs-edit, offline creation collisions, append-only review
  records, stale-write atomicity, local-only undo, permissions, immutable
  annotation identities, and legacy protocol-v1 reads.
- The React viewer captures the immutable PDFium annotation/form baseline,
  projects Yjs overlays in both directions, verifies source bytes and page
  count, and replaces EmbedPDF history with local-origin Yjs history. Viewer
  history is purged after projection so one user gesture cannot create two
  competing undo stacks.
- Browser projection currently accepts the five public editor tools: free
  text, highlight, underline, strikeout, and ink. Annotation kinds that need
  unsynchronized binary/context payloads, including stamps and signatures,
  fail closed until their authenticated asset port is available. Shared form
  records must match an existing writable field name.
- Vue emits `collaborationChange`; the Web Component dispatches
  `collaboration-change` and also supports a callback property. Viewport,
  search, selection, render caches, and page bitmaps stay local.
- Projection tests cover base annotation edits/deletion, created annotations,
  forms, remote echo suppression (including renderer-added default authors),
  ISO date portability, local undo, and irreversible tombstones.
- Rust, CLI, MCP, and A3S Code can set a PDF form value by stable field name.
  Existing values update only their conflict-local leaf; new records use the
  browser-compatible presence/fields/order encoding. Browser-to-Yrs fixtures,
  native restart/idempotency tests, real subprocesses, and concurrent
  browser/native edits prove form-value convergence without synchronizing PDF
  source bytes.
- Rust, CLI, MCP, and A3S Code can create, optimistically update, and
  irreversibly tombstone portable FreeText, Highlight, Underline, StrikeOut,
  and Ink annotations. Creation writes a complete immutable claim and always
  uses `source: created`; updates recursively compare mutable JSON leaves so
  unrelated concurrent edits merge while same-leaf conflicts fail closed; ID,
  page, type, and source identity cannot drift. Native restart, idempotency,
  real CLI/MCP subprocesses, browser Yjs delivery reordering, and a real
  EmbedPDF Playground projection cover create/update/delete. The deterministic
  Playground path is retained as an `a3s-test` ACL regression with visual,
  accessibility, console, and page-error evidence.
- Rust, CLI, MCP, and A3S Code can append a bounded redaction proposal, rotate
  or delete a validated source-page subset, propose a complete page
  permutation, and append the single final decision for an existing redaction
  or page operation. Actor IDs come from the replica manifest, timestamps use
  canonical UTC, and each typed record plus its browser-compatible canonical
  claim commit together. Missing targets, duplicate final decisions,
  conflicting same-ID reuse, invalid geometry, invalid rotation degrees,
  deleting every page, incomplete reorders, and out-of-range pages fail
  closed. Native restart/idempotency, real subprocess, and browser Yjs fixtures
  cover concurrent browser records plus reordered and duplicate native
  delivery. An exhaustive native property test delivers causally related
  rotation, deletion, reorder, and final-decision updates in all 24 orders,
  injects a duplicate, restarts the replica, and compares the final state
  vector and document digest with the source replica.

Remaining:

- Resolve signature appearance assets through an explicit authenticated host
  port and keep asset hashes aligned with the audit record.
- Apply approved redaction/page operations in a non-retryable host workflow,
  then save and reopen a merged PDF fixture to verify durable results.
- Extend offline/reordered-update property coverage and native Yrs parity to
  signatures and additional annotation-type payloads.

Exit criterion: annotations and forms converge and reopen from a saved PDF;
destructive actions remain attributable, reviewable, and non-retryable.

### Phase 6: CLI, MCP, and coding agents

Status: native Yrs replica store, resumable CLI/MCP event streams, `a3s code`
projection, a host-injected live CLI transport session, authenticated browser
suggestion-update authorization, and typed Markdown, Document content/comment,
Spreadsheet cell, Presentation scene-element, PDF annotation/form-value, and
PDF redaction/page-operation review mutation surfaces are implemented; typed
native Document suggestion/decision mutations, the remaining format mutations,
and native presence projection are pending.

- Yrs `0.27.3` now uses 53-bit Yjs-compatible client IDs and exchanges standard
  Yjs v1 updates, state vectors, and y-sync `SyncStep1`/`Update` messages with
  the browser package.
- Native replicas persist immutable, checksummed update entries and periodic
  full-state checkpoints/state vectors. Atomic no-clobber publication,
  contiguous sequence validation, startup replay, bounded automatic/manual
  compaction, and durable operation receipts provide crash recovery. Commit
  receipts are derived from authoritative checkpoint-plus-log replay; bounded
  canonical replay preserves genuinely missing updates while resolving Yrs
  array items whose causal dependencies arrived out of order.
- `a3s-office collab` exposes non-interactive `create`, `join`, `inspect`,
  `read`, `diff`/`synchronize`, `apply`, `mutate`, `checkpoint`, `watch`, and `leave`
  commands
  with JSON output and optional no-clobber binary output. `sync-step1`,
  `encode-update`, and `handle-message` expose standard y-sync document
  handshakes; mutating SyncStep2/Update messages use the same identity and
  receipt path. Resumable event cursors report compaction gaps as explicit
  full-state resets.
- The transport-neutral Rust session accepts the browser host-channel envelope,
  emits a fresh `SyncStep1` on connect/reconnect, answers peer handshakes with
  `SyncStep2`, durably applies incremental updates, suppresses their room echo,
  and projects updates committed by other CLI/MCP processes back to the host.
  A compaction gap sends one complete update and restarts the state-vector
  handshake instead of replaying an unsafe partial history.
- `a3s-office collab session` exposes that state machine as bounded JSONL over
  stdin/stdout for a coding-agent or product host. The host still owns the
  WebSocket or IPC channel, room selection, authentication, authorization,
  buffering, and delivery identities; the CLI never opens a provider itself.
- Each replica binds one stable actor ID, actor kind, permission mode, artifact
  identity/kind, namespace, and client ID. Apply/checkpoint/leave require the
  same identity plus a stable operation ID and accept an optional state-vector
  precondition.
- `collab mutate` and `office_collaboration_mutate` accept a closed typed
  operation instead of caller-authored Yjs bytes. Markdown replace/splice
  writes canonical `Y.Text` with browser UTF-16 offsets. Document exact-match
  replacement edits ProseMirror `Y.XmlText` in place, preserves the first
  replaced character's formatting attributes, rotates Word `textId`, and fails
  if the declared match count is stale. Bounded section/list/table/blockquote
  paragraph insert/delete uses explicit stable identities and exact deletion
  guards, while table-contained edits rotate all ancestor row text identities;
  page color and track-changes write their typed option fields. Document comment
  create/reply/resolve/delete writes attributable browser-compatible review
  records and exact selection marks. All emit minimal incremental updates
  through the same durable receipt/checkpoint path. Canonical content mutations
  require `edit`; Document review mutations accept `edit` or `comment`, with
  ownership-restricted deletion in `comment`. Raw remote updates remain
  receivable in every mode so read-only peers still converge.
- Host-authenticated Document `suggest` transport updates are semantically
  checked in Yrs before commit: only attributed insertion/deletion/replacement
  proposals owned by the authenticated actor may be added or withdrawn, while
  canonical state and every non-suggestion root remain protected. This is a
  transport authorization path, not a `NativeOfficeCollaborationMutation`;
  `collab mutate`, `office_collaboration_mutate`, and projection v2 do not yet
  expose native typed suggestion creation or final-decision operations.
- Rust `project`, `collab read`, and `office_collaboration_read` interpret the
  Office-owned browser schema inside Office rather than in a product host.
  Markdown returns its exact canonical source. Document returns bounded
  traversal-order paragraph records, stable `paragraphId`/`textId` pairs,
  structural ancestry, option fields, subordinate plain text, and projection-v2
  comment/reply/anchor/detached records together with the exact state vector.
  `document-replace-paragraph` uses those stable
  identities plus complete expected text to reject a stale same-paragraph
  browser/agent edit before writing, while unrelated changes can proceed
  without replacing the full document.
- Spreadsheet cell mutations create or recursively patch one zero-based
  coordinate after matching the caller's observed cell, or delete it only after
  an exact complete-cell match. They preserve the browser's field-addressed
  representation, dense/sparse projection mode, and atomic fail-closed
  semantics without replacing a worksheet or workbook.
- Presentation scene-element mutations create, update, move, or tombstone one
  stable object inside a slide, master, or layout. Canonical creation claims
  prevent conflicting same-ID reuse, optimistic top-level field guards merge
  unrelated concurrent edits, stable predecessor guards move one order entry
  without array indexes, and exact deletion guards prevent a stale destructive
  write without replacing the deck or container.
- PDF form-value mutations write the browser-compatible conflict-local record
  roots by fully-qualified field name, retain idempotent receipts and optional
  state-vector preconditions, and never synchronize source bytes.
- PDF annotation mutations create browser-compatible portable records, merge
  unrelated mutable JSON leaves under recursive optimistic guards, preserve
  immutable creation claims, and use irreversible deletion tombstones. The
  accepted annotation types are FreeText, Highlight, Underline, StrikeOut, and
  Ink; record identity and source-page bounds fail closed before mutation.
- PDF redaction and rotate/delete/reorder page-operation proposals plus final
  review decisions write append-only typed records and canonical creation
  claims, derive attribution from the replica, and reject stale identity,
  range, permutation, and target conflicts before producing an update.
- Durable operation receipts make identical retries idempotent and reject an
  operation ID reused for another payload. Artifact, kind, actor, mode,
  stale-state, invalid update, ambiguous bootstrap, corrupt log, and incomplete
  log failures are structured and fail closed.
- Cross-language tests import a fixture emitted by Yjs, validate browser
  metadata/bootstrap roots in Yrs, and apply the exported update to another
  Yrs peer. Core and process tests cover restart, checkpoint compaction,
  duplicate operations, stale preconditions, identity mismatches, and missing
  log entries.
- Validated browser update origins are retained in immutable operation receipts
  and resumable events, survive restart, and are re-emitted unchanged by native
  live transports. The host delivery operation remains distinct from the
  source browser operation, and source attribution is audit data rather than
  an authorization token.
- Cross-language tests apply deterministic native UTF-16 Markdown, ProseMirror
  Document text/options/paragraph/comment, Spreadsheet cell, Presentation
  scene-element content/z-order, and PDF annotation/form-value updates in browser Yjs,
  including concurrent browser/native paragraph, cell-leaf, element-field,
  annotation-leaf, and PDF form edits, in addition to importing browser Yjs
  fixtures in Yrs.
- Standard MCP exposes the durable replica lifecycle and bounded event stream;
  the same seven collaboration tools are explicitly available to the dedicated
  `a3s code` Use worker.

Remaining:

- Extend typed format-model mutations to Spreadsheet and remaining Presentation
  structural/rich-text operations plus PDF signatures; deepen Document
  mutations to additional nested structures and add typed suggestion creation,
  suggestion projection, and tracked final-decision operations for native
  clients.
- Project editor-visible presence and selection state while keeping Awareness
  ephemeral and outside native replica persistence.

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
- Claiming comments or suggestions are supported by a format that lacks a
  durable review schema and, for suggestions, explicit accept/reject semantics.
