import type { Editor } from '@tiptap/core';

const COMPOSITION_SETTLE_RETRY_MS = 20;
const MAX_COMPOSITION_SETTLE_RETRIES = 4;

/**
 * Keeps controlled TipTap editors from publishing or reconciling content while
 * the browser and ProseMirror still own an IME composition. The zero-delay
 * settlement lets ProseMirror flush WebKit's pending DOM records first; the
 * bounded retries cover runtimes that keep `view.composing` true briefly after
 * `compositionend` without adding work to ordinary keyboard input.
 */
export class ControlledEditorComposition {
  private active = false;
  private generation = 0;
  private settling = false;
  private timer: ReturnType<typeof setTimeout> | null = null;

  start(): void {
    this.generation += 1;
    this.cancelTimer();
    this.active = true;
    this.settling = false;
  }

  end(editor: Editor, onSettled: (editor: Editor) => void): void {
    const generation = ++this.generation;
    this.cancelTimer();
    this.active = false;
    this.settling = true;
    this.scheduleSettlement(editor, onSettled, generation, 0);
  }

  isBlocking(editor?: Editor | null): boolean {
    return (
      this.active ||
      this.settling ||
      Boolean(editor && !editor.isDestroyed && editor.view.composing)
    );
  }

  destroy(): void {
    this.generation += 1;
    this.cancelTimer();
    this.active = false;
    this.settling = false;
  }

  private scheduleSettlement(
    editor: Editor,
    onSettled: (editor: Editor) => void,
    generation: number,
    retry: number,
  ): void {
    const delay = retry === 0 ? 0 : COMPOSITION_SETTLE_RETRY_MS;
    this.timer = setTimeout(() => {
      this.timer = null;
      if (generation !== this.generation) return;
      if (editor.isDestroyed) {
        this.settling = false;
        return;
      }
      if (editor.view.composing) {
        if (retry < MAX_COMPOSITION_SETTLE_RETRIES) {
          this.scheduleSettlement(editor, onSettled, generation, retry + 1);
        } else {
          this.settling = false;
        }
        return;
      }
      this.settling = false;
      onSettled(editor);
    }, delay);
  }

  private cancelTimer(): void {
    if (this.timer === null) return;
    clearTimeout(this.timer);
    this.timer = null;
  }
}
