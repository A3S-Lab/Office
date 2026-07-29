export interface WorkspaceContextMenuAnchorBounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

const TEXT_CONTROL_MIRROR_PROPERTIES = [
  'box-sizing',
  'width',
  'height',
  'padding-top',
  'padding-right',
  'padding-bottom',
  'padding-left',
  'border-top-width',
  'border-right-width',
  'border-bottom-width',
  'border-left-width',
  'border-top-style',
  'border-right-style',
  'border-bottom-style',
  'border-left-style',
  'font-family',
  'font-size',
  'font-style',
  'font-weight',
  'font-variant',
  'font-stretch',
  'line-height',
  'letter-spacing',
  'text-align',
  'text-indent',
  'text-transform',
  'word-spacing',
  'tab-size',
  'direction',
  'writing-mode',
  'white-space',
  'overflow-wrap',
  'word-break',
  'scrollbar-gutter',
] as const;

export function workspaceTextControlSelectionBounds(
  control: HTMLTextAreaElement,
): WorkspaceContextMenuAnchorBounds | null {
  const { selectionEnd, selectionStart } = control;
  if (selectionEnd <= selectionStart) return null;

  const ownerDocument = control.ownerDocument;
  const view = ownerDocument.defaultView;
  if (!view || !ownerDocument.body) return null;

  const controlBounds = control.getBoundingClientRect();
  if (controlBounds.width <= 0 || controlBounds.height <= 0) return null;

  const computedStyle = view.getComputedStyle(control);
  const mirror = ownerDocument.createElement('div');
  mirror.setAttribute('aria-hidden', 'true');
  mirror.style.position = 'fixed';
  mirror.style.left = `${controlBounds.left}px`;
  mirror.style.top = `${controlBounds.top}px`;
  mirror.style.margin = '0';
  mirror.style.overflow = 'hidden';
  mirror.style.pointerEvents = 'none';
  mirror.style.visibility = 'hidden';
  for (const property of TEXT_CONTROL_MIRROR_PROPERTIES) {
    mirror.style.setProperty(
      property,
      computedStyle.getPropertyValue(property),
    );
  }

  const text = ownerDocument.createTextNode(`${control.value}\u200b`);
  mirror.append(text);
  ownerDocument.body.append(mirror);
  mirror.scrollLeft = control.scrollLeft;
  mirror.scrollTop = control.scrollTop;

  const range = ownerDocument.createRange();
  range.setStart(text, selectionStart);
  range.setEnd(text, selectionEnd);
  const selectionBounds = range.getBoundingClientRect();
  mirror.remove();

  return visibleSelectionBounds(selectionBounds, controlBounds);
}

function visibleSelectionBounds(
  selection: DOMRect,
  control: DOMRect,
): WorkspaceContextMenuAnchorBounds | null {
  const left = Math.max(selection.left, control.left);
  const top = Math.max(selection.top, control.top);
  const right = Math.min(selection.right, control.right);
  const bottom = Math.min(selection.bottom, control.bottom);
  if (right <= left || bottom <= top) return null;
  return {
    left,
    top,
    width: right - left,
    height: bottom - top,
  };
}
