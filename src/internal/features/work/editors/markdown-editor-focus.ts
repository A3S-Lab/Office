export type MarkdownEditingSurface = 'source' | 'visual';

const MARKDOWN_SURFACE_FOCUS_RETRY_FRAMES = 4;

export function restoreMarkdownEditingSurfaceFocus(
  getTarget: () => HTMLElement | null,
  focusOrigin: Element | null = document.activeElement,
): void {
  const commandTrigger = focusOrigin;
  let remainingFrames = MARKDOWN_SURFACE_FOCUS_RETRY_FRAMES;

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
