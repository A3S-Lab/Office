# Markdown editor reference

Markdown is a browser editor surface, not an OOXML package. Use the native
collaboration mutation route for durable Markdown state and the source-checkout
UI operator for interaction evidence:

```bash
bun run office:ops -- capabilities markdown --json
bun run office:ops -- check markdown --json
bun run office:ops -- visual markdown --project compact-768
```

The focused regression suite is
`tests/e2e/markdown-phone-modes.acl`. It covers source, visual, and split
modes, compact-pane switching, source editing, and focus restoration. Prefer
semantic targets when exploring:

- `role=textbox|Markdown 源码` for source editing;
- `role=textbox|Markdown 编辑区` for visual editing;
- `role=button|显示源码窗格`, `role=button|显示预览窗格`, and
  `role=button|切换 Markdown 视图` for compact mode changes;
- `css=.work-markdown-workspace` and
  `css=.work-markdown-pane.<source|visual>` for bounded layout assertions.

The visual pane is read-only. After a source edit, observe the rendered
heading or paragraph and assert that the active pane and keyboard focus remain
correct. On phone widths, set the viewport explicitly before the action and
capture both the selected compact pane and the final accessibility tree.

For collaboration, use `a3s use office collab mutate` with the Markdown
operation contract from the MCP reference. Prefer locate-then-replace:

1. `collab find <store> --find ...` or MCP `office_collaboration_find`
2. Pass that `matchCount` as `expectedMatches` on `markdown-replace-text`
3. Add optional `occurrence` (1-based from find) to change one match
4. Under concurrent peers, also pass the hit's `indexUtf16` so a drifted span
   fails closed even when the match count is unchanged

Keep `markdown-splice` for a known UTF-16 range. `expectedSlice` must equal the
text in `[indexUtf16, indexUtf16 + deleteUtf16)`, or be empty when `deleteUtf16`
is 0. A drifted slice returns `office.collaboration.mutation_match_conflict` and
writes nothing, so a stale whole-source splice cannot overwrite a peer edit.
Use `markdown-replace` only
as a compare-and-swap: `expectedMarkdown` must equal the source just read from
the replica. A drifted base returns `office.collaboration.mutation_match_conflict`
and writes nothing, so a stale whole-source rewrite cannot overwrite a peer
edit. Prefer `markdown-replace-text` for incremental edits. Inspect the replica
state first and keep the operation ID stable on retry. Do not treat a rendered
preview as proof that the persisted Markdown text is correct: read the replica
back after the mutation.
