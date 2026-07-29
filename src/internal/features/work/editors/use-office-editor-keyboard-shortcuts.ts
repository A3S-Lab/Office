import { type RefObject, useEffect } from 'react';
import type { OfficeEditorRuntime } from './office-editor-extension';

export interface OfficeEditorKeyboardShortcutOptions {
  capture?: boolean;
  enabled?: boolean;
  onHandled?: (event: KeyboardEvent) => void;
  scopeRef?: RefObject<HTMLElement | null>;
}

export function useOfficeEditorKeyboardShortcuts<Context, Commands>(
  editor: OfficeEditorRuntime<Context, Commands>,
  options: OfficeEditorKeyboardShortcutOptions = {},
): void {
  const { capture = false, enabled = true, onHandled, scopeRef } = options;

  useEffect(() => {
    if (!enabled) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (scopeRef && !keyboardEventBelongsToScope(event, scopeRef.current)) {
        return;
      }
      if (editor.handleKeyDown(event)) onHandled?.(event);
    };
    window.addEventListener('keydown', handleKeyDown, { capture });
    return () => {
      window.removeEventListener('keydown', handleKeyDown, { capture });
    };
  }, [capture, editor, enabled, onHandled, scopeRef]);
}

function keyboardEventBelongsToScope(
  event: KeyboardEvent,
  scope: HTMLElement | null,
): boolean {
  if (!scope) return false;
  if (event.target instanceof Node && scope.contains(event.target)) return true;
  const activeElement = document.activeElement;
  return activeElement instanceof Node && scope.contains(activeElement);
}
