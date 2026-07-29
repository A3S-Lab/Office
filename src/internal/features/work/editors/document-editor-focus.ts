const DOCUMENT_EDITOR_FOCUS_RETRY_FRAMES = 4;

export function restoreDocumentEditorFocus(
  getTarget: () => HTMLElement | null,
): void {
  const commandTrigger = document.activeElement;
  let remainingFrames = DOCUMENT_EDITOR_FOCUS_RETRY_FRAMES;

  const restore = () => {
    if (remainingFrames <= 0) return;
    remainingFrames -= 1;
    const target = getTarget();
    if (!target?.isConnected) {
      requestAnimationFrame(restore);
      return;
    }
    const activeElement = document.activeElement;
    const canRestore =
      activeElement === commandTrigger ||
      activeElement === target ||
      activeElement === document.body ||
      activeElement === document.documentElement ||
      !activeElement?.isConnected;
    if (canRestore) target.focus({ preventScroll: true });
  };

  requestAnimationFrame(restore);
}
