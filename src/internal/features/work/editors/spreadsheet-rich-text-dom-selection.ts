export interface SpreadsheetRichTextDomSelection {
  editor: HTMLElement;
  end: number;
  start: number;
}

export function captureSpreadsheetRichTextDomSelection({
  expandedOnly = false,
  root,
  target,
}: {
  expandedOnly?: boolean;
  root: HTMLElement | null;
  target?: EventTarget | null;
}): SpreadsheetRichTextDomSelection | null {
  const selection = window.getSelection();
  if (!root || !selection || selection.rangeCount !== 1) return null;
  const range = selection.getRangeAt(0);
  if (expandedOnly && range.collapsed) return null;
  const targetNode = target instanceof Node ? target : null;
  const editor = spreadsheetRichTextEditor(
    targetNode ?? range.commonAncestorContainer,
  );
  if (
    !editor ||
    !root.contains(editor) ||
    !editorContainsRange(editor, range)
  ) {
    return null;
  }
  const start = spreadsheetTextOffsetAtPoint(
    editor,
    range.startContainer,
    range.startOffset,
  );
  const end = spreadsheetTextOffsetAtPoint(
    editor,
    range.endContainer,
    range.endOffset,
  );
  return start === null || end === null || start > end
    ? null
    : { editor, end, start };
}

export function spreadsheetRichTextEditor(node: Node): HTMLElement | null {
  const element = node instanceof Element ? node : node.parentElement;
  const editor = element?.closest<HTMLElement>(
    '.luckysheet-cell-input, .fortune-fx-input',
  );
  return editor?.isContentEditable ||
    editor?.getAttribute('contenteditable') === 'true'
    ? editor
    : null;
}

function editorContainsRange(editor: HTMLElement, range: Range): boolean {
  return (
    (range.startContainer === editor ||
      editor.contains(range.startContainer)) &&
    (range.endContainer === editor || editor.contains(range.endContainer))
  );
}

function spreadsheetTextOffsetAtPoint(
  editor: HTMLElement,
  node: Node,
  offset: number,
): number | null {
  try {
    const prefix = document.createRange();
    prefix.selectNodeContents(editor);
    prefix.setEnd(node, offset);
    return prefix.toString().length;
  } catch {
    return null;
  }
}
