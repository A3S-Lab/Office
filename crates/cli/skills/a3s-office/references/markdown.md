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
operation contract from the MCP reference. Markdown splice offsets are UTF-16
code units; inspect the replica state first and keep the operation ID stable on
retry. Do not treat a rendered preview as proof that the persisted Markdown
text is correct: read the replica back after the mutation.
