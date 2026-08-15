# Changelog

All notable changes to A3S Office will be documented in this file.

## Unreleased

## 0.5.0 - 2026-08-15

- Added native Spreadsheet collaboration cell mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. `spreadsheet-set-cell` creates
  or recursively patches one browser-compatible cell with optimistic leaf
  guards, while `spreadsheet-delete-cell` requires an exact complete-cell
  match. Unrelated concurrent formula, value, style, hyperlink, note, and
  metadata leaves merge; same-leaf conflicts fail without a durable update.
  Dense sheets retain and safely extend their matrix projection, sparse sheets
  remain sparse, and an empty sheet's first write uses `celldata`. Coordinates,
  JSON depth and size, unsafe object keys, malformed shared roots, and orphaned
  fields are bounded and validated before mutation. Native restart, real CLI
  and MCP subprocesses, browser Yjs duplicate/reordered delivery, and a
  Playground `a3s-test` regression cover set, create, and delete behavior.

## 0.4.0 - 2026-08-15

- Added a versioned, deterministic Document Snapshot codec for lossless
  controlled-value persistence and an agent-readable Markdown Source
  projection for single-section documents. Snapshot decoding validates the
  schema, version, size, structured model, and synchronized HTML fingerprint;
  Source revisions retain Office-owned section layout and reattach only
  unambiguous surviving comment anchors. Public Core exports, README guidance,
  and focused round-trip and fail-closed tests cover both contracts.
- Hardened controlled Document editing and pagination under concurrent host,
  agent, observer, and font updates. External snapshots now apply the smallest
  history-neutral ProseMirror transaction, while a single-flight pagination
  coordinator coalesces observer churn and aborts only invalidated work. The
  Playground now lazy-loads editors, file import, and PDF evidence support from
  lightweight shared contracts, reducing the initial entry bundle from 716.8
  KiB to 71.4 KiB gzip while keeping the full editor available on interaction.
  Adaptive ribbons now derive density from the stable outer viewport so
  overflow navigation cannot trigger a resize oscillation.
- Added native PDF annotation create, optimistic leaf update, and irreversible
  delete mutations through Rust, `collab mutate`, standard MCP, and A3S Code.
  The closed surface accepts portable FreeText, Highlight, Underline,
  StrikeOut, and Ink records; creation writes a browser-compatible `created`
  record plus immutable claim, concurrent updates merge unrelated JSON leaves,
  same-leaf conflicts fail without an update, and deletion writes a durable
  tombstone. Identical retries are no-ops, immutable ID/page/type changes are
  rejected, and created claims remain valid after later mutable edits. Native
  restart, duplicate/reordered-delivery, real CLI/MCP subprocess, and browser
  Yjs interoperability tests cover the lifecycle. Exact native updates now
  project through the EmbedPDF harness with real nested Highlight geometry.
  The stable logical document digest canonicalizes JSON object arrays across
  Yrs restart and delivery order while raw commit detection continues to retain
  causally pending structs. The Playground exposes a deterministic remote
  create/update/delete fixture, with an `a3s-test` ACL regression that captures
  screenshots, accessibility, console, and page-error evidence.
- Added native append-only PDF redaction and page-operation review mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. `pdf-propose-redaction` writes
  bounded page geometry with the replica actor and a canonical UTC timestamp;
  `pdf-propose-page-rotation`, `pdf-propose-page-deletion`, and
  `pdf-propose-page-reorder` write validated source-page subsets or a complete
  permutation without changing source bytes; and
  `pdf-decide-review` writes the single attributable final decision for an
  existing redaction or page operation. Both records and their canonical
  creation claims are committed atomically, identical stable-ID retries are
  no-ops, and conflicting ID reuse, missing targets, duplicate final
  decisions, invalid geometry, unsupported rotations, page-range violations,
  deleting every page, and incomplete reorders fail without a durable update.
  Native restart tests, real CLI/MCP subprocesses, and browser Yjs fixtures
  cover concurrent edits plus duplicate and reordered delivery. An exhaustive
  native convergence test now checks all 24 delivery orders for causally
  related rotation, deletion, reorder, and final-decision updates, including a
  duplicate delivery and durable restart. Native replay treats the immutable
  checkpoint and raw update log as authoritative and canonically rebuilds Yrs
  pending structures before reporting the committed state vector.
- Added the first typed native PDF collaboration mutation through Rust,
  `collab mutate`, standard MCP, and A3S Code. `pdf-set-form-value` updates an
  existing conflict-local form-value leaf or deterministically creates its
  typed presence/fields/order record without synchronizing source bytes.
  Browser-generated Yjs fixtures, native restart/idempotency checks, real CLI
  and MCP subprocesses, and concurrent browser/native replay prove the update
  remains readable and convergent across Yjs and Yrs. Invalid or oversized
  field identities fail before any durable state changes.
- Extended typed native collaboration mutations to Document. Coding agents can
  replace an exact, fail-closed match count inside ProseMirror `Y.XmlText`,
  rotate the affected Word `textId` plus every identified ancestor table row's
  `rowTextId`, and insert or guarded-delete plain paragraphs in bounded section,
  nested list-item, table-cell/header, and blockquote containers without
  replacing the shared XML tree. Page-color and track-changes remain
  conflict-local options. Rust, CLI, MCP, and browser Yjs fixtures cover
  restart, idempotency, stale identity/text rollback, emoji offsets, nested
  tables, concurrent structural edits, and cross-language replay.
- Added idempotent typed native collaboration mutations through Rust,
  `collab mutate`, standard MCP, and A3S Code. The initial Markdown
  replace/splice surface writes canonical `Y.Text` with browser-compatible
  UTF-16 offsets, rejects surrogate-splitting or stale ranges, requires edit
  mode, and publishes its incremental update through the existing durable live
  session path.
- Native collaboration receipts now preserve validated browser source
  actor/operation origins separately from host delivery IDs. Attribution
  survives restart, appears in resumable CLI/MCP events, participates in
  idempotency conflicts, and is re-emitted unchanged to other live peers.
- Added a transport-neutral native collaboration session and the machine-only
  `collab session` JSONL bridge. Coding agents can now join a host-owned live
  room with browser-compatible `SyncStep1`/`SyncStep2` reconnect handshakes,
  durable delivery receipts, external CLI/MCP update projection, compaction
  recovery, typed outbound origins, and remote-echo suppression while the host
  retains connectivity, room, authentication, and authorization ownership.
- Added a bounded, identity-bound host transport adapter for Yjs v1
  state-vector/update synchronization with explicit reconnect, echo
  suppression, and typed incremental origins. Added a provider-owned Awareness
  controller with validated participant state and format-specific locations for
  Document, Markdown, Spreadsheet, Presentation, and PDF sessions.
- Live Word pagination now preserves each section's exact page size,
  orientation, margins, and page gap across mixed-layout documents. Kernel
  protocol 16 carries deduplicated page styles into both JS and Rust/WASM,
  returns metrics on every physical page, and drives variable-size page
  sheets, borders, navigation thumbnails, and exact per-page PDF capture.
- Added bounded structured OMML equations for Word documents. Inline and
  display equations now round-trip as native `m:oMath` and `m:oMathPara`
  across the document body, headers, footers, footnotes, and endnotes, with an
  accessible MathML preview and atomic insert/update commands. The structured
  display subset preserves `left`, `right`, `center`, and `centerGroup`
  paragraph justification, canonicalizes absent or empty justification to the
  `centerGroup` default, and requires one optional `m:oMathParaPr` before one
  `m:oMath` child. The structured
  subset covers Unicode runs with literal/normal-text semantics, six math
  scripts, four styles, manual breaks, and alignment points,
  plus ordered `m:rPr -> w:rPr -> m:t/w:t` runs with bounded Word fonts and
  theme references, Latin/complex-script bold and italic flags, all-caps and
  small-caps presentation, single and double strike, outline, shadow, emboss,
  imprint, proofing/grid flags, hidden and web-hidden states, direct/theme
  colors, signed character spacing through
  31,680 twips, 1–600% horizontal scaling, half-point kerning thresholds and
  signed baseline positions, half-point sizes, colored underline styles, all
  seven legacy text-animation values, all 27 line-border styles with
  direct/theme colors, 2–96 eighth-point widths, 0–31 point spacing, and
  explicit shadow/frame flags, every named highlight color, complete patterned
  run shading with direct or theme foreground/background colors, manual run
  widths from 0 through 31,680 twips with optional signed 32-bit grouping IDs,
  explicit baseline/superscript/subscript run alignment, all five Word
  emphasis-mark values (`none`, `dot`, `comma`, `circle`, and `underDot`),
  RTL/complex-script flags, language tags, and East Asian typography metadata
  with optional signed 32-bit run IDs, two-lines-in-one flags, all five
  enclosing-bracket styles, horizontal-in-vertical rotation, and rotated-text
  compression, explicit paragraph-mark always-hidden/reset flags, and Office
  2010 text glow, shadow, reflection, text-outline, and text-fill effects. Glow
  retains an optional 0 through 2,147,483,647 EMU radius, one RGB or 17-slot
  theme color source, and up to 64 ordered, repeatable tint, shade, alpha,
  hue-modulation, saturation, and luminance transforms. The distinct Office 2010
  shadow effect retains the same
  color model plus optional 0 through 2,147,483,647 EMU blur and offset
  coordinates, a direction from 0 inclusive to 360 degrees exclusive, signed
  horizontal and vertical scales, skew angles strictly between -90 and 90
  degrees, and all ten rectangle alignments. Angles retain exact
  1/60,000-degree units and scales retain exact 1/1,000-percent units.
  The leaf Office 2010 reflection effect retains optional blur and distance
  coordinates, start/end opacity and position from 0 through 100 percent,
  direction and fade direction from 0 inclusive to 360 degrees exclusive,
  signed horizontal and vertical scales, skew angles strictly between -90 and
  90 degrees, and the same ten rectangle alignments. Angles retain exact
  1/60,000-degree units, while opacity, positions, and scales retain exact
  1/1,000-percent units.
  The structured Office 2010 text-outline effect retains an optional width from
  0 through 20,116,800 EMUs, three line caps, five compound-line styles, two pen
  alignments, distinct none/solid/gradient fills, optional lists of 2 through 10
  gradient stops with RGB or theme colors and ordered transforms, exact linear
  or path shading with optional signed 32-bit relative fill rectangles, all 11
  preset dashes, and round, bevel, or miter joins with optional exact
  nonnegative limits.
  The Office 2010 text-fill effect reuses the same strict no/solid/gradient fill
  grammar, bounded colors and transforms, 2 through 10-stop lists, linear/path
  shades, signed relative rectangles, and exact units without outline geometry.
  Explicit zero/default geometry values reset inherited formatting instead of
  canonicalizing away. Strict universal font-size and position measures enter
  the model only when they convert exactly to the bounded half-point form.
  Strict universal manual widths are accepted only when they convert exactly
  to bounded whole twips; omitted grouping IDs remain distinct from explicit
  zero. Explicit baseline alignment remains present so inherited superscript or
  subscript formatting can be reset, and explicit `none` emphasis remains
  present so inherited emphasis marks can be removed.
  Empty `w:eastAsianLayout` elements canonicalize away, while omitted flags
  stay distinct from explicit `false` resets and signed run IDs retain explicit
  zero. An empty `w:specVanish` canonicalizes to `true`; omission remains
  distinct from an explicit `false` inheritance reset.
  Omitted `w14:glow/@w14:rad` keeps its zero schema default while explicit zero
  remains present. Export declares `w14` and merges it into `mc:Ignorable`.
  Omitted `w14:shadow` geometry keeps its zero/`none` defaults while explicit
  zero and `none` remain present; the effect remains distinct from legacy
  `w:shadow` on/off formatting.
  A present empty `w14:reflection` remains distinct from omission. Its omitted
  geometry keeps zero/`none` defaults while explicit zero and `none` remain
  present.
  A present empty `w14:textOutline` remains distinct from omission and keeps
  the schema's bevel default. Omitted fill, dash, and join choices retain their
  defaults, while explicit zero/default attributes and empty child choices
  remain present.
  A missing `w14:textFill` continues to use `w:color`; a present empty wrapper,
  empty solid fill, or gradient without a stop list retains the distinct black
  schema default.
  All-caps and small-caps presentation, character spacing, width scaling,
  effective kerning, baseline shifts, baseline/superscript/subscript alignment,
  and exact transform-free Office 2010 RGB text fills or black fill defaults use
  safe MathML/CSS projections without changing source Unicode text. Word
  emphasis marks project through CSS as filled dots,
  a literal comma, or an open circle above the text, or a filled dot below it.
  Superscript and subscript also project the smaller rendered size required by
  Word. When `w:position` and `w:vertAlign` coexist, both remain in native
  schema order and the later explicit alignment controls the CSS vertical
  position. `w:em` remains after `w:rtl`/`w:cs` and before `w:lang`, while
  `w:eastAsianLayout` remains after `w:lang`, `w:specVanish` follows it,
  `w14:glow` follows `w:specVanish`, `w14:shadow` follows `w14:glow`, and
  `w14:reflection` follows `w14:shadow`, followed by `w14:textOutline` and then
  `w14:textFill`. Simple
  explicitly sized solid, double, dotted, dashed, inset, and outset line
  borders project through CSS with direct or automatic color and point
  padding; explicit `nil`/`none` resets also project. Relief effects, legacy
  text animations, complex
  multi-line, wavy, or 3D line borders, border shadow/frame, theme-only border
  colors, and hidden or web-hidden states remain native metadata because Word
  view and rendering settings govern them. Manual run widths also remain
  native-only because Word ignores `w:fitText` inside Office Math, so the
  MathML preview deliberately does not emulate them. East Asian
  two-lines-in-one, enclosing brackets, horizontal-in-vertical rotation, and
  rotated-text compression also remain native-only because CSS writing modes,
  text combination, and transforms cannot reproduce Word's inline line-box
  semantics without layout drift. `w:specVanish` also stays native-only and
  never hides equation previews because its display semantics apply only to
  paragraph marks; Word additionally ignores it unless `w:vanish` is set.
  Schema-valid values remain preserved without inventing that dependency.
  Office 2010 glow, shadow, reflection, and text-outline effects also remain
  native-only because CSS `text-shadow`, reflection, opacity, transform,
  text-stroke, paint-order, and border approximations cannot preserve theme
  colors, ordered transforms, exact blur and offset coordinates, reflection
  opacity/position/fade geometry, signed scale/skew, rectangle alignment,
  gradient or compound strokes, preset dashes, caps, joins, or pen alignment.
  Text `noFill`, theme or transformed text-fill colors, and nonempty gradients
  remain native-only; previews keep readable fallback color instead of using
  fragile transparent-text or background-clipped-gradient approximations.
  Highlight precedence over shading is
  retained. Named highlights, explicit highlight removal, clear direct fills,
  solid direct foregrounds, and nil shading project through MathML
  `mathbackground`; pattern masks and theme-only colors remain native metadata.
  Explicit on/off values survive native regeneration. Enabled mutually
  exclusive casing, strike, or relief combinations, invalid animation values,
  art-border styles, out-of-range border width/spacing, malformed border
  colors/flags, missing, malformed, fractional, or out-of-range manual widths
  and grouping IDs, missing or unknown vertical-alignment or emphasis-mark
  values, malformed or out-of-range East Asian layout IDs, flags, or bracket
  styles, malformed paragraph-mark visibility flags, malformed glow radii,
  shadow or reflection geometry, text-outline fill/gradient/dash/join
  structure, text-fill wrapper/fill/gradient structure, color choices, or
  transform chains, and
  unknown, reordered,
  duplicated, spoofed, or
  relationship-bound Word run properties fail closed.
  Supported object property containers also preserve one optional ordered
  `m:ctrlPr` control format through that bounded property model. The control
  may contain a direct `w:rPr` or tracked `w:ins`, `w:del`, `w:moveFrom`, or
  `w:moveTo` provenance with bounded IDs/authors, optional validated core dates
  and Microsoft 365 `w16du:dateUtc` values with a UTC `Z` suffix,
  Word-legal `moveFrom/moveTo -> ins/del` and `ins -> del` nesting, and an
  optional deepest `w:rPr`. Empty direct control properties canonicalize away;
  empty revisions remain native provenance. Every supported `deg`, `den`, `e`,
  `fName`, `lim`, `num`, `sub`, and `sup` argument slot now retains the same
  direct or revision-wrapped control format after its expressions. Fixed slots
  use named metadata; matrix cells, equation-array rows, and delimiter arguments
  use strictly dimension-aligned metadata. Argument formatting, revision
  provenance, and document-level move-range pairing remain native-only because
  they are absent from professional MathML. Unknown, malformed, illegally
  nested, spoofed, or relationship-bound control markup fails closed. Safe
  object-control values project only onto separable MathML control/operator
  nodes while all supported values remain in native metadata.
  Matrix properties now preserve the ordered
  `baseJc -> plcHide -> rSpRule -> cGpRule -> rSp -> cSp -> cGp -> mcs -> ctrlPr`
  grammar. The five spacing rules, unsigned-short row/gap values, and minimum
  column widths through 31,680 twips round-trip with attribute-free Word
  defaults. Row and column gaps project to safe MathML table spacing, while
  minimum width remains native-only for layout because MathML exposes fixed
  column width.
  N-ary `grow` and delimiter `grow`/`shp` properties now round-trip with their
  distinct object defaults and attribute-free enabled values. Growing n-ary
  operators project to MathML `stretchy`; fixed delimiters project with
  `stretchy=false`, and content-matched growing delimiters use
  `symmetric=false`. Native export retains schema order and canonicalizes
  default non-growing n-aries and growing centered delimiters.
  bar/no-bar/skewed/linear fractions, super- and subscripts, aligned right-side
  sub-superscripts, left-side pre-sub/superscripts with empty script slots,
  radicals with optional degrees and canonical hidden empty degree slots,
  functions, supported n-ary operators, combining accents, overbars and
  underbars, group characters with explicit character
  position and baseline justification, phantoms with visible or hidden bases,
  independently zeroed width, ascent, or descent, and transparent spacing,
  border boxes with independently visible edges and four strike directions,
  semantic boxes with operator-emulation,
  no-break, differential-spacing, manual-break, and alignment properties,
  bounded rectangular matrices with base and column alignment, ordered
  row-spacing and column-gap rules, and minimum column widths, equation arrays
  with 1–64 rows, vertical base alignment, maximum/object distribution,
  row-spacing rules, and `&` alignment/spacer markers, lower and upper limit
  objects, and custom delimiters in strict or transitional UTF-8/UTF-16
  packages. Fraction properties enforce `fPr -> num -> den` ordering and
  canonicalize an absent or attribute-free `type` to the `bar` default.
  N-ary operators enforce optional `naryPr` before required `sub`, `sup`, and
  `e` slots. An absent `chr` defaults to U+222B, an attribute-free `limLoc`
  defaults to `undOvr`, disabled growth normalizes away, enabled growth
  round-trips, and hidden limits use canonical empty script slots.
  Attribute-free operator characters and contradictory hidden nonempty limits
  fail closed. Delimiters
  enforce optional `dPr` before 1–32 `e` arguments and the ordered
  `begChr -> sepChr -> endChr -> grow -> shp -> ctrlPr` property grammar.
  Omitted characters normalize to `(`, U+2502, and `)`, while attribute-free
  character properties remain explicitly empty. Empty argument slots and the
  growing, centered defaults canonicalize away; non-growing and shape-matched
  delimiters round-trip. Functions enforce optional `funcPr` before required
  `fName` and `e` slots and preserve empty name or argument slots. Every
  supported `CT_OMathArg` slot now preserves an empty argument and enforces
  `argPr -> expressions -> ctrlPr` ordering. Its optional trailing `ctrlPr`
  retains one bounded direct or revision-wrapped Word control. Absent or empty argument/control
  properties and absent, empty, or zero `argSz` values normalize to the
  default. Relative sizes from -2 through 2 round-trip in every argument slot.
  The 13 Word-effective parent/child pairs project through inverse-sign relative
  MathML `scriptlevel`; valid sizes elsewhere remain native-only. Out-of-range
  or malformed sizes, duplicate or misplaced properties, malformed control
  revisions, and semantic properties fail closed. Depth, node, text, model-size,
  matrix-dimension, cell-count,
  equation-array row/alignment-marker, and equation-count budgets are enforced.
  Invalid or non-combining accent characters, malformed math-run or function
  structures, invalid or contradictory fraction, radical, n-ary, delimiter, bar,
  group-character, phantom, border-box, box, or equation-array properties,
  malformed
  lower/upper limit structures,
  malformed, duplicated, reordered, or out-of-range matrix spacing/gap
  properties, ragged or over-limit matrices, over-limit
  equation arrays, malformed script-property, pre-script, or math-paragraph
  structures, malformed placement, namespace spoofing, nested math, and
  relationship-bound properties
  fail closed to bounded text with explicit diagnostics.
- Added source-backed DOCX package preservation. Browser import now registers
  the original package, and export retains safe source-only OPC parts,
  content-type declarations, and relationships while generated core parts stay
  authoritative. Invalidated signatures, VBA, ActiveX, and custom-ribbon parts
  are deliberately omitted, and a missing registered source fails closed
  instead of silently producing a lossy export. Imported source metadata also
  carries a SHA-256 fingerprint so a different re-registered DOCX is rejected.
- Preserved passive OOXML extensions inside regenerated `word/settings.xml`.
  Ignorable extension attributes and elements plus structurally valid,
  non-conflicting `mc:AlternateContent` blocks now survive strict or
  transitional UTF-8/UTF-16 source packages. Generated Word settings remain
  authoritative; malformed, relationship-bound, protection, template,
  mail-merge, field-update, and duplicate-setting markup is not restored.
- Preserved passive ignorable extension trees on regenerated DOCX styles and
  numbering definitions. Styles match by type and style ID, while imported
  abstract-numbering, concrete-numbering, and level metadata follows rewritten
  numbering IDs. Generated Word semantics remain authoritative; source-only,
  duplicate, relationship-bound, malformed, and ambiguous one-to-many
  extension mappings are dropped instead of attaching to the wrong identity.
- Preserved passive non-OOXML vendor extensions on uniquely matched picture
  drawings in regenerated document, header, and footer parts. Drawing identity
  uses the normalized anchor plus drawing-property ID across strict/transitional
  UTF-8/UTF-16 sources. Header and footer image identities now survive editable
  page-chrome HTML. Relationship-bound, source-only, duplicate, Microsoft/OOXML
  semantic, and ambiguous drawing branches remain disconnected.
- Preserved passive non-OOXML vendor extensions on stable paragraphs and their
  paragraph properties in regenerated document, header, and footer parts.
  Native `w14:paraId` plus `w14:textId` identities survive sanitized body and
  page-chrome HTML; text edits rotate the version ID, while formatting-only
  edits and moves retain it and copies or splits receive independent paragraph
  IDs. Changed text versions, duplicate identities, relationship-bound content,
  and Microsoft/OOXML semantic branches fail closed.
- Preserved passive non-OOXML vendor extensions on stable `w:tbl`, `w:tr`, and
  `w:tc` scopes plus their property nodes. Native row `w14:paraId` and
  `w14:textId` identities now survive body and page-chrome HTML; table and cell
  identity is conservatively derived from directly owned row and paragraph
  IDs. Row text or structure edits rotate the row version, copied identities
  are repaired, nested rows and cells no longer leak into outer-table export,
  and duplicate, cross-kind, relationship-bound, or semantic branches fail
  closed while generated table geometry and formatting remain authoritative.
- Preserved stable native footnote, endnote, comment, and reply identities
  across reorderings while assigning fresh IDs to copies and collisions.
  Resolved comments now emit valid `commentsExtended.xml` even without replies;
  passive extensions on uniquely matched note, comment, and `commentEx` roots
  survive regeneration, and valid `commentsIds` durable IDs are rebound to the
  final comment paragraph IDs. Deleted records, duplicate or namespace-spoofed
  identities, relationship-bound branches, and unsupported modern
  reaction/people sidecars fail closed instead of reviving stale metadata.
- Added native editable DrawingML pictures inside footnotes and endnotes.
  Public import and artifact export retain picture identity, layout, wrapping,
  crop, and layer metadata across the body and both note parts. Export repairs
  missing note-part image relationships from the OOXML writer, assigns
  collision-free relationship IDs, and validates every embedded media target.
  Passive non-OOXML extensions follow only uniquely matched note drawings
  across strict/transitional UTF-8/UTF-16 sources. Changed, duplicate,
  namespace-spoofed, relationship-bound, or semantic branches fail closed;
  generated geometry and media remain authoritative, while legacy VML,
  shapes, SmartArt, and drawing-bearing content-control wrappers normalize.
- Preserved text-stable direct runs inside uniquely matched footnotes,
  endnotes, comments, and replies. Passive extensions on paragraph, run, and
  run-property scopes now follow exact text and structural ancestry; safe
  unmodeled note properties survive, and unchanged plain-text comments regain
  relationship-free source run segmentation and formatting. Supported
  regenerated semantics remain authoritative, while edits, duplicate
  paragraphs or properties, wrapped or mixed semantic content, relationship
  references, and ambiguous mappings fail closed.
- Preserved supported hyperlink wrappers and their stable runs inside
  text-stable footnotes, endnotes, comments, and replies. Generated note links
  remain authoritative, while unchanged plain-text comments recover safe
  HTTP(S), `mailto`, or internal-anchor links, eligible tooltips, passive
  wrapper metadata, and relationship-free formatting. External relationships
  are validated against the owning part, deduplicated or assigned a
  collision-free ID, and rewritten in final XML. Text edits, missing or
  duplicate relationships, wrong types or target modes, relative or unsafe
  targets, combined external-plus-anchor destinations, namespace spoofing,
  unsupported wrappers, and ambiguous spans fail closed.
- Preserved text-stable static rich-text and plain-text content controls inside
  footnotes, endnotes, comments, and replies. Eligible inline controls and
  contiguous block controls recover their wrappers, aliases, tags, locks,
  signed native IDs, Word 2013 appearance and color, end-character formatting,
  passive extensions, and stable runs. Footnote and endnote tables now export
  as native editable OOXML blocks instead of flattened row text; structurally
  stable and nested tables can participate in rich-text block controls while
  generated geometry remains authoritative. ID collisions are rewritten
  without disturbing unconflicted source IDs. Text or table-structure edits,
  duplicate or ambiguous mappings, active bindings or placeholder state, form
  or nested controls, relationship-bound content, hyperlinks, math, and
  drawings fail closed before any wrapper is emitted.
- Preserved source DOCX font-table metadata and eligible internal obfuscated
  font payloads through relationship-ID collisions. Font references are
  rewritten to their final package IDs, while external references, wrong
  relationship or content types, duplicate identities, and payload-path
  collisions fail closed. Strict and transitional UTF-8/UTF-16 package XML is
  decoded consistently. Embedded fonts remain available to native DOCX
  consumers; browser editing, preview, and PDF export still use registered A3S
  fonts or substitution.
- Defined authoritative controlled-update behavior for reviewed Word ranges.
  Comment and tracked-change mutations now produce typed React, Vue, and Web
  Component conflict events plus an accessible warning, while harmless range
  movement and document switches remain quiet and orphaned comment records are
  retained for host recovery.
- Added editable Word tight and through image wrapping, including wrap-side
  controls, browser `shape-outside` contour presentation, and exact DrawingML
  `wrapPolygon` import/export for supported floating pictures without leaking
  internal export markers.
- Added editable four-edge Word image cropping with percentage validation,
  matching edit/preview presentation, and exact DrawingML `a:srcRect`
  preservation for inline and floating pictures. Export patches the owning
  picture deterministically without leaking internal markers into the DOCX.
- Preserved and authored precise Word floating-image anchors with signed
  horizontal and vertical offsets relative to the column, paragraph, margin,
  or page. Picture Properties validates the complete placement atomically,
  edit/preview apply matching offsets, and DOCX `positionH`/`positionV`
  round-trip without converting aligned anchors into offset anchors.
- Continued row-spanning Word table cells across every covered physical row
  during pagination. Combined `rowspan`/`colspan` cells now receive contiguous
  selection ranges and in-cell page-break widgets, while DOCX `vMerge` and
  `gridSpan` round-trip together.
- Completed editable nested Word tables across insertion, targeted inner-table
  sizing, DOCX import/export, and pagination. Outer rows can now split at
  nested-row boundaries instead of forcing a tall inner table to overflow as
  one atomic block.
- Added percentage-based table-column authoring in the Layout ribbon and Table
  Properties dialog. Percentage preferences survive merged cells and DOCX
  `tcW` round-trips while pixel `tblGrid` widths remain browser fallbacks.
- Preserved semantic Word theme color references for run text, run shading,
  table-cell fills, and independent cell borders, including tint and shade.
  Untouched formatting now writes the original theme attributes with a correct
  RGB fallback, while explicit color edits discard stale theme semantics.
- Added bounded PDFium-native text runs with stable indices, character and
  UTF-8/UTF-16 ranges, exact PDF-space bounds, deterministic validation, and an
  independent hard run limit to the native PDF text-layer receipt. PDFium 7881
  segment look-ahead indices are normalized as exclusive ends so valid final
  text runs cannot be rejected as out of range.
- Recorded the completed native exact-unit source-layout contracts: bounded,
  content-addressed PPTX slide rasters and PDFium-backed PDF page inventory,
  geometry, rendering, typed failures, and deterministic receipts now satisfy
  the Office-side requirements tracked in #1 and #4.
- Preserved imported Word numbering identities, abstract-numbering identities,
  and levels in the controlled document model. Separated list runs that belong
  to one native Word list now reuse one DOCX numbering instance on export
  instead of silently restarting under unrelated generated identities.
- Preserved native multilevel `numFmt` and compound `lvlText` patterns across
  controlled edits and DOCX export, including non-Latin numbering families.
  Continue Numbering now adopts the preceding native identity, while an
  explicit style change clears stale imported formatting metadata.
- Preserved native numbering suffix, level alignment, physical and logical
  indentation, hanging or first-line offsets, and `lvlRestart` rules. RTL list
  definitions keep `start`/`end` semantics instead of being flattened to
  `left`/`right` during browser editing and DOCX export.
- Added conditional table-style support for paragraph contextual spacing and
  outline levels across style precedence, controlled editor attributes, format
  copy and clearing, and DOCX export.
- Added a Spreadsheet command catalog and adopted the shared WPS-oriented
  quick-access, adaptive, and collapsible ribbon. Conditional Formatting now
  lives under Home and Styles, Data exposes executable ascending and descending
  sort commands, and workbook recalculation is visible and executable with F9.
- Added focused component and controller coverage, desktop and compact
  Playwright regression, and a schema-validated deterministic A3S Test workflow
  for the aligned Spreadsheet ribbon.
- Added executable Paste, Cut, and Copy commands to the Spreadsheet Home
  clipboard group. Ribbon clicks and WPS `Cmd/Ctrl+V`, `Cmd/Ctrl+X`, and
  `Cmd/Ctrl+C` shortcuts now share one typed command port, permission-resilient
  browser/local clipboard fallback, and grid-focus restoration.
- Added a WPS-style Spreadsheet Format Painter to the Home clipboard group.
  Single-click one-shot and double-click locked sessions copy native cell-style
  patterns across ranges and sheets without changing values, formulas,
  comments, links, or merges; another click or Escape exits cleanly.
- Added bounded format capture and target guards, duplicate-target suppression,
  one controlled workbook batch per application, accessible pressed/live
  state, and desktop plus compact Web regression coverage.
- Added WPS-style Spreadsheet AutoFilter under Data and Sort and Filter.
  `Cmd/Ctrl+Shift+L` toggles filtering, `Alt+ArrowDown` opens the selected
  header menu, and arrows, Space, Enter, and Escape operate it without leaving
  the grid.
- Added finite current-region discovery for single-cell selections, exact
  explicit-range filtering, safe empty/merge/pivot rejection, controlled
  selection and hidden-row preservation, accessible vendor filter controls,
  XLSX round-trip coverage, and desktop plus compact Web regression.
- Added WPS-style Spreadsheet Freeze Panes under View and Window. The current
  cell freezes the rows above and columns to its left, with separate top-row,
  first-column, and unfreeze commands behind one controlled workbook update.
- Added Arrow, Home, End, Enter, and Escape menu operation, pressed and live
  state, selection and grid-focus restoration, XLSX round-trip coverage, and
  desktop plus compact Web regression. Delayed grid-focus recovery now yields
  to deliberate pointer and Tab navigation so repeated ribbon actions remain
  usable immediately after a controlled workbook remount.
- Added a WPS-familiar Rows and Columns menu to Spreadsheet Home and Cells.
  The existing typed workbook commands now expose row insertion above or below,
  column insertion left or right, and selected-row or selected-column deletion
  without duplicating the structure-editing model.
- Added independent command availability, Arrow/Home/End/Enter/Escape menu
  behavior, exact grid-focus restoration, desktop and compact Web workflow
  regression, and schema-validated deterministic A3S Test coverage.
- Added a WPS-familiar Merge and Center split control to Spreadsheet Home and
  Alignment. Its menu executes Merge and Center, Merge Cells, Merge Across,
  Unmerge Cells, and Unmerge and Fill, while `Ctrl+M` shares the primary path.
- Kept every merge intent within one controlled Fortune workbook batch, used
  the native merge model for availability and unmerge ranges, restored grid or
  invoker focus exactly, and added focused, XLSX round-trip, desktop, and
  compact Web regression coverage.
- Added the WPS Clear menu to Spreadsheet Home and Editing with Clear All,
  Clear Formats, Clear Contents, Clear Comments, and Clear Hyperlinks. Delete
  and Backspace now share the typed Clear Contents path.
- Preserved content, formats, comments, hyperlinks, and merge geometry according
  to each Clear mode, including bounded range subtraction for borders,
  conditional formats, and alternating formats; each intent stays within one
  controlled workbook batch and restores grid focus.

## 0.3.0 - 2026-08-07

- Added a Writer command catalog for stable ribbon grouping and WPS-compatible
  shortcut metadata, moved undo and redo into a compact quick-access toolbar,
  added persistent plus temporary ribbon collapse behavior, and made lower
  priority groups compact before the ribbon falls back to horizontal paging.
- Aligned Writer superscript and subscript with the WPS `Ctrl+Shift+=` and
  `Ctrl+=` shortcuts and added deterministic desktop browser coverage for the
  expanded, collapsed, and temporary ribbon states.
- Made the displayed WPS Writer shortcuts executable inside the document for
  font sizing, paragraph alignment and line spacing, heading styles, spelling,
  field refresh, comments, and track changes without capturing host inputs.
- Added a permission-free Writer formatting clipboard with WPS
  `Ctrl+Shift+C` / `Ctrl+Shift+V`, a one-shot format painter, semantic-mark
  preservation, and single-transaction formatting paste.
- Extended WPS alignment and format-copy shortcuts into page headers and
  footers, corrected their superscript and subscript shortcut descriptions,
  and added schema-safe body-format projection for page-chrome editors.
- Reordered the Writer Insert ribbon into WPS-familiar Pages, Table,
  Illustrations, Links, Header and Footer, and Text groups, with page-number
  visibility beside header and footer commands.
- Added direct Writer Page Layout presets for margins, orientation, paper size,
  and one-to-three-column layouts, with custom margins and advanced columns
  routed to the matching Page Setup tab. Deterministic browser coverage proves
  live landscape and two-column rendering, Escape close, accessibility, and
  empty console and page-error diagnostics.
- Aligned Writer References, Review, and View grouping with WPS terminology and
  order, added direct previous/next plus accept/reject revision commands, and
  marked Picture, Table, and Header and Footer tabs as contextual tools.
- Replaced arbitrary Writer zoom presets with WPS-style 100%, One Page, and Page
  Width commands. Fit zoom is calculated from the live page and editor viewport;
  deterministic browser coverage resolves two tracked changes, verifies both
  fit modes, and captures accessible, error-free evidence.
- Made Writer's status-bar word count actionable with live page, word,
  character, and paragraph details plus the WPS `Ctrl+Shift+G` shortcut. The
  labelled view-and-zoom toolbar supports arrow-key traversal, while compact
  Web layouts retain page and zoom controls before lower-priority status items.

## 0.2.2 - 2026-08-07

- Selected the Word `ascii`, `hAnsi`, `eastAsia`, or complex-script font slot
  from each run's actual text while preserving `bCs`, `iCs`, `szCs`, `cs`,
  `rtl`, and font-hint behavior for multilingual DOCX content.
- Added a deterministic 30-row Latin, Chinese, Arabic, Hebrew, and mixed-format
  fixture with A3S Test coverage and a real WPS Writer PDF layout gate.
- Added a calibrated Chromium native-PDF fallback for WPS reference captures
  when the embedded PDF renderer cannot initialize the exported document.

## 0.2.1 - 2026-08-06

- Matched WPS Writer automatic line layout across common Latin and Chinese
  system fonts with measured per-font advances while retaining the original
  OOXML line-spacing multiple as the DOCX round-trip authority.
- Preserved section-level Word document-grid type and line pitch plus run-level
  `snapToGrid` overrides across DOCX import and export, and stopped exporting a
  generated document grid when the source document does not define one.
- Added deterministic 30-row common-font, 36-row CJK-font, and 18-row document-
  grid fixtures with A3S Test browser coverage and real WPS PDF layout gates.

## 0.2.0 - 2026-08-06

- Added a real WPS Writer page-layout gate that exports a deterministic A4 DOCX
  through WPS, captures normalized A3S and WPS pages, and rejects page-size,
  semantic-landmark, browser-error, or bounded pixel regressions.
- Matched WPS automatic Word line spacing without changing the original OOXML
  multiple used for DOCX export, removed editor-only spacing around imported
  tables, and removed the transparent paginated-page border from content
  geometry.
- Reworked the default Markdown split view into a flat writing-and-reading
  workspace with a bounded text measure, clearer typography for headings,
  quotations, code, tables, and task lists, and no nested preview card chrome.
- Replaced the unusable stacked Markdown phone split with a full-workspace
  Source/Preview switch whose controls remain touch-sized while preserving the
  controlled Markdown value and synchronized visual tree.
- Promoted shared editor context menus to a viewport-bound phone action sheet
  with 44 px rows, bounded internal scrolling, safe-area spacing, and the same
  keyboard dismissal and focus restoration as the desktop menu.
- Expanded shared Office select menus to 44 px option rows on phones, with a
  taller viewport-bounded scroll region and preserved End-key selection,
  editor-focus recovery, and Escape-to-trigger restoration.
- Separated persistent desktop sidebar state from the temporary phone drawer,
  so resizing across the compact breakpoint closes the overlay immediately,
  preserves workspace focus, and restores the prior desktop preference.
- Kept the documentation language and version selectors visible at ordinary
  laptop widths; the site opens in Simplified Chinese on `latest` and retains
  the frozen `0.1.0` documentation for version switching.
- Separated Word Table Properties updates from preceding edits in TipTap
  history, so undoing a property change no longer removes a newly inserted
  table.
- Replaced the isolated Word alternative-text prompt with one responsive
  Picture Properties workflow for centimeter width and height, a per-image
  aspect-ratio lock, wrapping, alignment, text distance, and alternative text.
- Applied each Picture Properties draft as one separated TipTap history entry,
  while cancel and Escape leave the document unchanged, retain the selected
  image, and restore the exact ribbon invoker.
- Preserved untouched imported image dimensions instead of materializing their
  rounded centimeter display values, and added unit, Playwright, and phone A3S
  Test coverage for size coupling, compact controls, focus, and diagnostics.
- Expanded Word Table Properties into Table, Row, Column, and Cell tabs with
  one validated TipTap transaction for preferred table geometry, selected-row
  sizing and pagination, current-column width, and selected-cell alignment and
  margins.
- Preserved untouched imported table dimensions and partial cell-margin
  inheritance at their exact source values instead of quantizing them through
  the centimeter display fields.
- Kept custom select menus inside the active modal focus scope and routed
  document undo and redo shortcuts from non-text ribbon controls without
  intercepting native input history.
- Extended the deterministic styled-DOCX A3S Test workflow across the complete
  responsive Table Properties dialog, preview fidelity, accessibility, and
  empty console and page-error diagnostics.

## 0.1.0 - 2026-08-02

- Extracted the complete A3S Web document, spreadsheet, presentation, and PDF
  editor engine into an independent package.
- Added React, Vue 3, Web Component, and framework-free Core entry points.
- Added DOCX, XLSX, PPTX, PDF, HTML, Markdown, text, CSV, XLS, and ODS file
  workflows.
- Added a colocated PDFium WebAssembly asset with an overridable URL.
- Added a browser-neutral exact-unit Office layout-renderer contract and a
  source-bound, no-resampling implementation for opaque PNGs that completely
  cover one PPTX slide, including deterministic profile receipts, source
  mutation checks, bounded no-clobber output, and typed unsupported outcomes
  for every richer layout.
- Added constant-scope inspection of pages from a previously validated native
  PDF inventory, allowing large-document consumers to reuse one complete page
  scan while render-time source and profile checks remain authoritative.
- Added an interactive React playground, type checks, integration tests, and
  Rslib packaging.
- Added GFM task lists, coalesced source-to-visual Markdown updates,
  synchronized split-pane scrolling, and dedicated compatibility tests.
- Added source-aware Markdown ribbon commands and shortcuts plus host-defined
  selected-text menus for both source and visual editing surfaces, with typed
  React, Vue, and Web Component APIs and stale-selection protection.
- Added controlled Markdown source history with typing coalescing, toolbar and
  keyboard undo/redo, selection restoration, and safe rebasing after host
  content replacement without polluting visual-editor history.
- Added a shared bounded Rust Spreadsheet formula parser plus a cancellable
  Worker/WebAssembly scalar calculation kernel with sparse requests,
  deterministic dependency order, cross-sheet references, target-only
  recalculation, bounded dependency depth, JavaScript parity fallback,
  dependency-failure propagation, and ordered cell-scoped Fortune fallback.
- Added `kernelWasmUrl` support for Spreadsheet in React and Vue and the
  matching `kernel-wasm-url` Web Component attribute.
- Added persistent nested Presentation groups with atomic selection and
  collective transforms, plus native PPTX group-node export for slides,
  layouts, and master-derived artwork. Supported group scale is normalized
  across geometry, typography, rich-text runs, and border weights on import;
  rotated or reflected source groups remain an explicit compatibility warning.
- Virtualized Presentation thumbnail scenes in both the slide strip and sorter.
  Every slide keeps a stable keyboard target and scroll footprint, while only
  the selected slide and the viewport overscan mount full scene content.
- Added primary-ribbon Spreadsheet number and percent formats, decimal-place
  controls, and readable percentage defaults in the quarterly-plan template.
- Added direct Presentation playback from the beginning or current slide with
  F5/Shift+F5 shortcuts, automatic fullscreen, and an in-page fallback with an
  explicit exit path.
- Shared the keyboard-accessible table-dimension picker across Document and
  Presentation, with exact row and column creation in one controlled update.
  Desktop keeps the fast 8 × 10 matrix without duplicate table-cell semantics;
  phones use focused row and column controls with 44 px targets and an
  editor-accented, white-text insertion action.
- Kept Word caption numbers and inline cross-references synchronized in the
  live TipTap transaction graph. Deleting or reordering a caption now
  renumbers surviving targets, updates their references, and exposes a visible
  `Missing reference` state for dangling fields instead of leaving a valid-
  looking stale number.
- Kept PDF page navigation, zoom, and history reachable through the compact
  overflow menu while preserving page status in the primary toolbar.
- Moved the phone PDF page-drawer trigger into the toolbar page controls so it
  no longer overlays document content, while retaining modal focus isolation,
  current-page synchronization, and focus restoration after selection.
- Removed internal implementation terminology from the Spreadsheet
  conditional-format manager.
- Added a visible prompt to the empty Markdown source pane without changing its
  controlled content.
- Added Spreadsheet font-family, vertical-alignment, and text-wrap commands to
  the primary ribbon through the native Fortune cell-format model.
- Added an editor-owned Spreadsheet Find bar with Cmd/Ctrl+F interception,
  displayed-value, raw-value, formula, and sparse-cell matching, deterministic
  cell navigation, repeated-shortcut refocus, and grid-focus restoration.
- Kept the phone Spreadsheet Find bar inside the viewport with a 40 px input
  and 40 px previous, next, and close actions, while preserving exact result
  navigation and grid-focus restoration after Escape.
- Made Spreadsheet workbook task panes modal at phone widths, isolated the
  ribbon, grid, and worksheet footer while open, contained forward and reverse
  Tab navigation, and restored the exact ribbon invoker after Escape.
- Made the Presentation chart inspector a modal surface whenever its responsive
  layout overlays the canvas. The close action receives initial focus, the
  ribbon, slide workspace, and status bar remain inert, Tab stays contained,
  dirty fields consume the first Escape, and closing restores the selected
  chart without changing the desktop docked inspector.
- Made Presentation comment review a full-editor modal on phones while keeping
  the desktop review strip docked. The phone surface uses readable review
  typography and touch-sized actions, isolates the ribbon, slide workspace,
  and status bar, contains forward and reverse Tab navigation, lets a dirty
  comment consume Escape before close, and restores the exact New Comment,
  View Comments, or comment-marker invoker.
- Kept common Presentation arrangement commands visible at desktop width by
  compacting group, ungroup, and layer actions without removing their labels
  from accessible names and tooltips.
- Added PDF annotation opacity and compatible stroke-width editing through a
  compact keyboard-accessible style popover, typed capability commands, and
  native PDF annotation defaults and selection updates.
- Rebuilt Presentation transition controls as standard ribbon groups, paged
  compact Office ribbons by complete command groups, and reset stale ribbon
  scroll state when the available width grows.
- Reserved both compact-ribbon navigation edges while tools overflow, bounded
  Word list galleries to the phone viewport, enlarged their numbering controls
  for touch, and restored the TipTap selection after list commands so active
  style and start-value settings remain available when a gallery is reopened.
- Made open popovers explicit editor-shortcut boundaries so Escape closes the
  PDF annotation-style popover without cancelling the selected annotation
  tool.
- Split Word page setup into keyboard-accessible Page, Columns and Sections,
  and Header and Footer tabs so paper controls stay focused and heavyweight
  header/footer editors mount only when requested.
- Stacked the file command bar above non-PDF ribbon tabs at phone widths so
  filenames and actions no longer compress or overlap the keyboard-accessible,
  horizontally scrollable tab row; PDF retains its single compact toolbar.
- Replaced the phone Presentation thumbnail rail with a dismissible,
  focus-managed slide drawer so the editing canvas keeps the primary width.
- Unified modal focus boundaries across the compact Office sidebar,
  Presentation slide drawer, and AI assistant. Focus now enters the visible
  surface, wraps on Tab and Shift+Tab, keeps the background inert, lets only
  the top overlay consume Escape, and returns to the invoking control.
- Kept the phone PDF toolbar clear of host file actions and moved secondary
  annotation tools, opacity, stroke width, and deletion into its scrollable,
  keyboard-operated overflow menu.
- Matched registered document font weights using the CSS Fonts search order so
  common 680/730 heading weights and browser-synthesized bold metrics stay on
  the deterministic Rust/WASM text-layout path instead of falling back to DOM
  line measurement.
- Unified Word body typography, headings, paragraphs, lists, quotations, and
  image wrapping across editing, read-only preview, and PDF composition.
- Positioned preview and PDF headers and footers inside the configured page
  margins without shifting body content; empty headers are no longer rendered
  and PDF composition no longer inserts the filename as an implicit header.
- Preserved physical Word page width and margins in compact preview so a
  narrow viewport scrolls the page instead of changing line wrapping.
- Kept the canonical TipTap document and its Worker/WASM page decorations
  mounted when switching between Word editing and read-only preview, preserving
  page count, automatic breaks, font shaping, and table pagination instead of
  rebuilding a separate HTML preview.
- Corrected page-layout geometry so headers and footers overlay their physical
  top and bottom margins without reducing the body height a second time; the
  kernel protocol is now version 15.
- Made browser Word PDF export consume the mounted TipTap and Worker/WASM
  pagination surface through a stable `artifactId` across React, Vue, and Web
  Components. Export now preserves automatic breaks, shaped runs, table
  continuations, page geometry, and page chrome while capturing long documents
  in bounded batches; the Playground exposes both DOCX and PDF from one compact
  header export menu.
- Kept Word task-pane headers stretched to the pane edges at compact and phone
  widths so titles and close actions retain the shared Office alignment.
- Kept a Word picture selected while its alternative text is edited and
  returned keyboard focus to the exact Picture ribbon command after save or
  cancel, preventing the next key press from changing document content.
- Positioned the Word comment composer against its selected text before the
  browser's first paint, removing the initial jump from the top of the review
  rail.
- Kept Word revision review keyboard focus on the matching action for the next
  change, then returned it to the document after the final individual decision.
- Made the empty Word revision pane reflect whether new changes are actually
  being recorded and added in-pane start/stop controls, keeping the phone modal
  workflow actionable and focus-stable after the last revision is resolved.
- Highlighted all Word Find matches persistently, made initial forward and
  backward navigation select the expected match, and kept Find and Replace
  controls focused across repeated keyboard actions.
- Extended the Word navigation pane from heading-only filtering to contextual
  full-text results with persistent match highlights and exact selection
  jumps. Compact navigation now closes before restoring body focus and the
  selected range.
- Kept new Word comment composers inside the visible review rail for long
  selections and limited discard prompts to comment drafts or replies with
  written content.
- Returned keyboard focus to the unfinished Word comment, reply, or citation
  field when a user cancels closing or switching its task pane, preserving the
  draft and the user's editing position across desktop and compact layouts.
- Preserved unfinished Word replies and edited citation fields when users
  cancel comment deletion, citation deletion, or an internal citation switch,
  and returned focus to the exact field instead of the destructive action.
- Moved Word citation validation beside the invalid tag or title field, added
  accessible error relationships and invalid styling, and focused the field so
  phone users can repair and complete the citation workflow immediately.
- Guarded worksheet deletion with the shared safe-default Office confirmation
  dialog, and kept invalid inline renames open with concise visible and
  accessible validation instead of silently discarding the entered name.
- Enlarged the shared Office color palette into an eight-column phone layout,
  kept the panel inside the viewport, and made vertical keyboard navigation
  follow the rendered grid across theme and standard colors.
- Stopped delayed Spreadsheet grid-focus recovery from stealing focus after a
  letter starts cell editing, preserving direct text entry after F2 and Escape.
- Kept multi-cell paste selection mutable across Fortune Sheet and React state
  replays, preventing a frozen-range crash while preserving the pasted range
  for subsequent copy, cut, and undo commands.
- Released the Presentation slide drawer's modal focus isolation as soon as a
  responsive viewport returns to desktop width, while preserving the open rail
  and moving focus to the active slide.
- Stabilized Presentation focus across React commits after cutting the focused
  object, keeping the active slide inside the editor shortcut scope so an
  immediate paste restores and focuses the clipboard object.
- Moved phone PDF search and page navigation to a dedicated second toolbar row
  so clearing a query cannot collide with the host download action.
- Replaced Playground AI implementation snippets with concise, file-specific
  guidance for documents, Markdown, spreadsheets, presentations, and PDFs.
