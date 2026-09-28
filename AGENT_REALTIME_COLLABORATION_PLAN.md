# Agent realtime collaboration plan

Last reviewed: 2026-09-27

The primary goal of A3S Office is that an agent and a human share one live
document. Each agent edit is a transaction on that document. The caret the
human sees is the insertion point of that transaction after layout.

`COLLABORATION_ROADMAP.md` remains the protocol. `ROADMAP.md` remains the
Traditional Office parity inventory. This plan sets the order of work. Parity
that does not make the live caret true waits.

## Target

A remote observer, watching the Office editor, sees the agent's glyphs appear
at the agent's caret. The caret is the editor's own insertion point for that
actor. A concurrent human edit before that point moves it. A stale address
fails closed and does not rewrite the rest of the document. Undo removes only
the local actor's transactions.

Pixel position is a function of that insertion point after layout. Layout
does not advance the document and does not move the caret on its own.

## One clock

There is one sequence: the ordered session stream of the live replica.

A frame on that stream is one of two things:

- a document update, whose caret is the end position of that update;
- a caret move that changes no document bytes.

The editor applies frame N and places the caret from frame N. It does not
wait for a later packet to learn where the caret went. A document update and
a separate location message are two clocks, even if a comment says they are
"the same logical step."

These are not clocks:

- Layout, pagination, zoom, and line wrap. They compute pixels from the
  position frame N already carried. They never emit a caret of their own.
- OOXML, XLSX, PPTX, and PDF bytes. They are an import into the replica or an
  export from it. While the replica is live, `office_open` does not hold a
  second mutable session for that artifact.
- A plain-text projection used to guess a ProseMirror position. The position
  is the one in the update.

Awareness stays ephemeral and stays out of the replica, the checkpoint, and
the operation log. For a document update it is not sent at all: the end
position is inside the update. For a caret move with no edit, the move is the
next frame on this same stream, ordered before the next document update. It
is not a side channel that can arrive early, late, or alone.

## First principles

1. One replica (`Y.Doc` / Yrs) is the document. One session stream is its
   clock.
2. An edit is one transaction with a typed origin. The caret after the
   transaction is part of that transaction's result.
3. The caret's coordinate system is the editor's: ProseMirror position for
   Document, UTF-16 offset for Markdown source, cell for Spreadsheet until an
   in-cell editor exists, element-local text position for Presentation once
   text is a fragment, page or annotation for PDF.
4. The overlay paints `coordsAtPos` (or the surface's equivalent) of the
   position the frame carried, after layout of that same frame. A stale
   measure waits for the next frame. It does not invent a position.
5. A stable address plus the observed slice is the conflict check. If that
   slice moved, the frame writes nothing. The agent reads the replica and
   continues from the caret the replica now has.
6. Local undo tracks only the local origin.
7. A surface that cannot place a glyph caret does not invent one. A
   spreadsheet cell and a PDF page publish the cell, page, or annotation that
   the same frame wrote.

## What is in the way

The replica stream is the clock. These are the pieces that are still not that
clock:

- Spreadsheet row and column insertion rewrites cell formulas, named-range
  references, print-area references, merged-range anchors, chart reference
  strings, and table ranges in the same transaction, and moves the cells
  with them. A three-dimensional reference, a pivot table, or a corrupt
  address fails the frame and writes nothing. Chart pixel boxes stay put.
  Sorting a rectangle permutes its rows in one frame. Formula text moves
  with the cell, and references outside the rectangle stay as written. A
  stale observed cell, or a merge, table, or pivot that meets the rectangle,
  writes nothing. Native table create, update, and delete write the same
  `tables` record, `tableOrder`, and creation claim the browser reads. A stale
  observed record, an overlapping range, or a repeated name writes nothing.
  Deleting the record leaves the claim, so that ID cannot name a different
  table. A worksheet cell stays put. Renaming the table or a column, or
  changing its range, header, or totals row, rewrites structured references
  in that same transaction. A string literal and an external-workbook
  reference stay as written. A geometry change, or a table-local reference
  whose owning cell cannot be proven, writes nothing. Deleting a table that
  a structured reference still names writes nothing and leaves the hash
  unchanged. Deleting an unreferenced table with a built-in style writes
  that style onto cells that already exist in the same transaction: header
  and stripe fills, text color, bold, and a thin border around the table
  range. Empty positions stay empty. Cells outside the range, formula text,
  and the creation claim stay put. A range larger than 100,000 cells writes
  nothing.
  A plain-text splice already paints its `indexUtf16` caret inside the cell.
  Formula cells, numeric cells, and styles stay field writes and do not
  grow a glyph caret. Two offline edits of different cells, or of different
  fields of the same cell, converge to one document hash in every delivery
  order. Delivering one of those updates again leaves the hash unchanged.
  Exporting that merged replica and reading the worksheet back yields the
  same number, formula, and cached value. A cell neither edit touched stays
  as it was. Calculated values stay cached until browser and native
  recalculation are deterministic and the source formula is the only
  canonical formula state.
- Presentation slide order, grouping, and background are conflict-local frames.
  Moving a slide rewrites only `slide-order`; a predecessor that is not the
  observed one writes nothing, and shape text stays put. Grouping writes
  `groupIds` on every listed member in one transaction; one member whose
  path does not match writes nothing, and unrelated shapes stay put.
  A slide, master, or layout background is one field in that frame; a stale
  background writes nothing. The browser content replace checks the live
  replica before it writes those frames: a stale slide order, group path, or
  background leaves that frame unwritten, and a slide background frame also
  leaves `useLayoutBackground` unwritten. A text splice already paints its `indexUtf16`
  caret inside the shape. A browser element-order replace checks the live
  replica: a permutation of the same elements writes nothing when that order
  already drifted, and writes when the live order is still the observed one.
  Shape text stays put. A slide thumbnail or a measured layout box stays on
  the client. A replace that only carries `thumbnail`, `thumbnailUrl`,
  `measuredLayout`, `measuredBox`, or `lineBoxes` writes nothing and leaves
  the state vector unchanged. When that replace also changes shape text, the
  text and the authored geometry are written and those derived fields are
  absent from the replica. Two offline edits of a shape's text and fill,
  together with another shape's fill, converge to one document hash in every
  delivery order. Delivering one of those updates again leaves the hash
  unchanged. Exporting that merged replica and reading each shape back yields
  the same text. A shape neither edit changed stays as it was.
- A host save of a spreadsheet walks the live replica and writes every
  populated cell into the package in one export: formula text with its cached
  value, a plain number, inline text, and a shared string owned by exactly
  one cell. A cell whose replica value does not match the package cell, or a
  shared string used by more than one cell, fails that export and returns no
  package. A document export walks the live replica and writes every
  paragraph into `word/document.xml` by `paraId`. A paragraph with no
  `paraId`, a package that has no such paragraph, or a paragraph that is not
  a single text run fails that export and returns no package. A presentation
  export walks the live replica and writes every shape text into its slide
  by element id. A container that has text and no slide part, a missing
  shape, or a shape that is not a single text run fails that export and
  returns no package. A PDF export walks the live replica and writes every
  form field and every FreeText annotation literal into the uncompressed PDF.
  A field or FreeText annotation missing from the file, a value that cannot
  be represented as a literal, or bytes that are not UTF-8 fails that export
  and returns no bytes.
  Compressed PDF values fail closed. Parts the export does not rewrite stay
  byte-identical. While that replica is live, `office_open` does not open a
  second session.
- Two browsers and one native client share one document frame. The native
  agent inserts CJK one grapheme at a time, and that frame's caret is the
  end of the grapheme. Each browser paints it after pagination, on the
  wrapped line, at a zoom other than 100%. A drifted slice writes nothing.
  A human insert before the caret shifts the next splice.

`document-replace-text` still rotates `textId`. It is a finished
substitution. Typing uses `document-splice` and does not rotate `textId`.
Presentation typing uses `presentation-splice` on the element's text
fragment. An unrelated shape edit does not rewrite that text.

## What stays

Do not rebuild the protocol.

- Yjs in the browser, Yrs in Rust, standard v1 updates and state vectors.
- The host owns transport, auth, rooms, and retention. Office does not open a
  provider.
- Separate roots for canonical content and review data. Caret state does not
  join those roots.
- Typed origins and local-only undo.
- Projection schema v4 addresses.
- Fail-closed stale guards. No last-write-wins of an unrelated region.
- The browser overlay that paints a validated position with `coordsAtPos`.

Drop the second sequence. A native Awareness peer that emits
`outbound-awareness` after a mutate is a second clock. The mutate result
carries the caret. The session writes one frame.

## Target agent edit

One realtime agent edit is one frame:

1. Read the address from the replica (`collab read` /
   `office_collaboration_read`, or `office_collaboration_find` when locating
   text).
2. Submit one splice at that address with the observed slice.
3. The frame result is the document update plus the next caret: same identity,
   index advanced by the inserted UTF-16 length, or a hard failure and no
   write.
4. The editor applies that update, maps the caret that came with it, lays
   out, and paints that caret.

A caret move that inserts nothing is also one frame, with empty document
bytes, ordered on the same stream. A burst of glyphs is a sequence of frames.
It is not a finished file played back on a timer.

While a replica is the live document, the agent does not `office_open` it.
Import builds the replica. Export writes bytes when the host saves.

## Plan

Work in this order. Each phase ends when its exit test is green in this
repository. Do not start a later phase to avoid an earlier one.

### 1. Markdown is the reference caret

Markdown already splices `Y.Text`.

- `markdown-splice` returns the caret in the same result:
  `indexUtf16 + insertUtf16` on success. No follow-up location message.
- The browser paints that returned offset with the editor caret path. Do not
  convert it through a plain-text walk that adds a newline per block.
- A human insert before the offset moves the agent's following splice. A
  stale `expectedSlice` writes nothing.

Exit: two browsers and one native client. The agent inserts a CJK string one
grapheme cluster at a time. The remote caret sits on the insertion edge after
each update. A concurrent insert before the caret shifts it. A drifted slice
fails closed.

### 2. Document splice that keeps its address

Add `document-splice` on one `Y.XmlText`:

- Identity: `paragraphId`, `textId`, `indexUtf16`.
- Guard: `expectedSlice` for the deleted UTF-16 range, empty when
  `deleteUtf16` is 0.
- Insert is plain text at that index. Do not split a surrogate pair.
- Success returns the next index in the same text node.
- Do not rotate `textId` on an insert or delete that stays inside the node.
  Rotation stays on `document-replace-text` and on edits that change the
  identity of existing text. A caret stream that rotates `textId` per glyph
  cannot take the next step.
- Preserve marks at the insertion boundary the same way a local typing
  transaction does.
- The splice result includes the post-splice position. The browser binding
  places `anchor` / `head` from that result while applying the update.

`document-replace-text` stays for substitutions the agent intends as one
replacement. It is not the typing primitive.

Exit: the same test as phase 1 on a paginated Document. The caret box is the
editor caret after pagination settles, including a line wrap and a CJK
fullwidth character. A human edit before the point maps the agent's next
splice. A stale `textId` or slice writes nothing. Local undo does not remove
the remote glyphs.

### 3. One projection, after layout

- The overlay measures the position carried by the frame just applied, after
  the pagination pass for that same update.
- Measuring before that pass, or from a location message that is not the
  frame, is a second clock and is rejected.
- A stale measure skips the paint. The next frame supplies the position. Do
  not substitute the document end.
- The pixel check is "inside the caret's line box, on the insertion edge",
  not a fixed distance guessed from a replay buffer.

Exit: a playground or `a3s-test` case with pagination on, zoom not 100%, and
a wrapped paragraph. The painted caret tracks the native agent's last
successful splice.

### 4. Spreadsheet tells the truth

A field write from `spreadsheet-set-cell` or `spreadsheet-batch-cells` is the
cell. One cell write is one transaction. A plain-text splice paints the
returned index inside that cell. Do not reveal the new display string glyph
by glyph on the canvas.

The plain-text in-cell caret is that splice index inside the cell box. Formula
text, cached values, and styles stay field-addressed leaves. Calculated values
stay derived.

Exit: a remote cell highlight lands on the cell the agent just wrote, and
does not blink on an unchanged presence update. A plain-text splice paints
its returned `indexUtf16` on the insertion edge inside that cell. Formula
text, cached values, and styles stay field-addressed and do not paint a
glyph caret.

### 5. Presentation text becomes a fragment

Scene text already splices a collaborative fragment and the slide paints that
index inside the shape. Slide order, grouping, and background are separate
frames: the order array, every member's `groupIds`, or one container
`background` field. A stale predecessor, group path, or background writes
nothing.

Exit: two clients type in one text box. Carets are element-local positions.
An unrelated shape edit does not move them. Scalar run replacement is no
longer the agent write path for that text.

### 6. PDF has no flowing caret

Agent presence stays on page, annotation, or form field. Annotation and form
mutations stay the closed operations in the protocol. Do not put a typewriter
on a PDFium bitmap.

Exit: unchanged from collaboration phase 5, plus an explicit test that the
frame result of an agent form or FreeText edit is that annotation or field,
with no text offset into page pixels and no follow-up location message.

### 7. Snapshot on request

Export DOCX / XLSX / PPTX / PDF from the replica when the host saves. The
agent does not `office_save` after every glyph. Import creates or updates a
replica once; further agent edits use phases 1–6.

Exit: a save after a remote splice round-trips the spliced text and leaves
unsupported OOXML untouched. A second client still converges after reopen.

## Order against the parity roadmap

Keep the permanent gates in `ROADMAP.md`: no silent clobber, one intent one
undo record, fail closed.

Then:

1. Phases 1–3 of this plan (Markdown caret, Document splice, layout
   projection).
2. Phase 5's text-fragment binding, before more slide animation parity.
3. Phase 4's honest cell presence, then the in-cell caret. Formula and
   workbook semantics continue only when they do not pretend the canvas is
   typing.
4. Phase 6 and 7.
5. Remaining Traditional Office rows in `ROADMAP.md`.

## Still open

Each clock below is one replica. A frame lands completely or writes nothing.
Office does not open a network provider. The host owns rooms, auth, and
transport.

1. Spreadsheet calculated values. A comment and an insertion suggestion
   created offline converge to one document hash in every delivery order.
   Delivering one of those updates again leaves the hash unchanged. The
   comment body and the insertion both remain, and the text outside the
   comment stays. Creating a different body for a comment ID that already
   exists writes nothing. Accepting one offline insertion and rejecting
   another converges in every delivery order: the accepted characters stay,
   the rejected characters do not, and the text outside both revisions stays.
   Delivering one of those decisions again leaves the hash unchanged. A later
   opposite decision for a suggestion that already has a decision writes
   nothing. Accepting one offline deletion and rejecting another converges
   in every delivery order: the accepted characters stay gone, the rejected
   deletion stays, and the text outside both revisions stays. Accepting the
   deletion half of a replacement and rejecting its insertion half converges
   in every delivery order: the deleted characters stay gone, the rejected
   insertion does not remain, and the text outside that replacement stays.
   Delivering one of those decisions again leaves the hash unchanged.
   Cached spreadsheet values stay in the replica until browser and
   native recalculation are deterministic and the source formula is the only
   canonical formula state.
2. PDF signature appearance. The host authenticates the asset port. Office
   does not open that port. Asset hashes stay aligned with the audit record.
3. Approved PDF redaction and page operations. A non-retryable host workflow
   applies one approved decision, then saving and reopening the merged PDF
   shows that result. Applying the same decision again writes nothing.
4. PDF signatures and annotation types other than form values and FreeText.
   Offline updates converge in every delivery order. A repeated update leaves
   the document hash unchanged.
5. One Boot process is not a splittable room. Sticky routing or shared
   fan-out plus a writer lock stays with the host. Office does not open
   Redis or NATS.

Structural Document work that the caret needs still happens: the splice must
be mapped through lists, tables, and pagination. Full table, list, and
section convergence from collaboration phase 2 stays required for the
phase 2 exit when the caret sits in those containers. It is not a reason to
ship more ribbon before the caret is real.

## Non-goals

- A second sequence beside the replica stream: an OOXML session open on the
  same artifact, an awareness packet after a mutate, or a playback timer.
- A host-side player that diffs a finished document and reveals glyphs on a
  timer.
- A plain-text index used as a ProseMirror caret.
- Persisting the caret into the replica.
- Office opening its own network provider.
- Whole-package replace as the realtime write.
- Per-glyph `textId` rotation.
- A cross-product "universal document" so one caret widget can serve the grid,
  the slide, and the page.
- Macro execution.

## Proof

Each phase lands with a failing test first, in the Office crate that owns the
behavior:

- Rust: splice success address, stale slice, surrogate boundary, no `textId`
  rotation on pure insert, caret present in the frame result, caret bytes
  absent from the replica log.
- Browser: applying the update moves the editor caret with no second message;
  pagination and zoom do not change the model position; the painted rect is
  that caret after layout of the same frame.
- CLI/MCP subprocess: mutate then read returns the next address; a second
  client sees the same bytes.

A green ribbon test is not evidence for this plan.
