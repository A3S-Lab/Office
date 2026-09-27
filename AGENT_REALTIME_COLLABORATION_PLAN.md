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

- A spreadsheet cell write still names the cell. Plain-text typing uses
  `spreadsheet-splice` and returns `indexUtf16` on that same cell. The canvas
  highlight stays the cell; it does not reveal glyphs one by one. Formula
  cells, cached numbers, and styles stay on `spreadsheet-set-cell`.
- A host save writes one addressed value back into the package. DOCX
  paragraph text, an inline spreadsheet cell, one presentation shape run,
  and an uncompressed PDF field or FreeText literal are the snapshot.
  Shared strings, formulas, numeric cells, and compressed PDF values fail
  closed. Other parts stay byte-identical. While that replica is live,
  `office_open` does not open a second session.
- Two browsers and one native client share the document frame. The native
  agent inserts CJK one grapheme at a time. Each browser paints that frame's
  caret after pagination, on the wrapped line, at a zoom other than 100%.
  A drifted slice writes nothing. A human insert before the caret shifts
  the next splice.

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

Until the in-cell editor owns a text caret, agent presence on a spreadsheet
is the cell from `spreadsheet-set-cell` / `spreadsheet-batch-cells`. One cell
write is one transaction. Do not reveal the new display string glyph by glyph
on the canvas.

The in-cell caret is the next spreadsheet collaboration feature, and it is
the same frame as phase 2 inside one cell's rich text: splice, expected
slice, returned index inside that result. Formula text, cached values, and
styles stay field-addressed leaves. Calculated values stay derived.

Exit: a remote cell highlight lands on the cell the agent just wrote, and
does not blink on an unchanged presence update. No glyph animation is
claimed. The in-cell splice is a separate exit, copied from phase 2, and is
not started by faking glyphs first.

### 5. Presentation text becomes a fragment

Phase 3 of `COLLABORATION_ROADMAP.md` already says scene text must bind to
collaborative XML fragments instead of scalar run replacement. That binding
is now on the critical path, because a text-box caret cannot exist before it.

After the fragment exists, a text box uses `document-splice` rules
(phase 2) with `containerKind`, `containerId`, and `elementId` as the address
prefix. Until then, an element text update is atomic and presence is the
element frame.

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
