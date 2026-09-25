import { readFileSync } from 'node:fs';
import { describe, expect, test } from '@rstest/core';

describe('document host editor hooks', () => {
  test('gives the host the live editor and an initially open comment pane', () => {
    const source = readFileSync(
      'src/internal/features/work/editors/document-editor.tsx',
      'utf8',
    );
    expect(source).toContain('onEditorReady?: (editor: Editor) => void');
    expect(source).toContain('onEditorReady?.(current)');
    expect(source).toContain('defaultCommentsOpen?: boolean');
    expect(source).toContain('defaultOpen: defaultCommentsOpen');
  });
});
