import { useEffect } from 'react';
import type { OfficeEditorRuntime } from './office-editor-extension';

export interface OfficeEditorKeyboardShortcutOptions {
  capture?: boolean;
  enabled?: boolean;
}

export function useOfficeEditorKeyboardShortcuts<Context, Commands>(
  editor: OfficeEditorRuntime<Context, Commands>,
  options: OfficeEditorKeyboardShortcutOptions = {},
): void {
  const { capture = false, enabled = true } = options;

  useEffect(() => {
    if (!enabled) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      editor.handleKeyDown(event);
    };
    window.addEventListener('keydown', handleKeyDown, { capture });
    return () => {
      window.removeEventListener('keydown', handleKeyDown, { capture });
    };
  }, [capture, editor, enabled]);
}
