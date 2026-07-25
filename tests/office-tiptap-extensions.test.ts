import { Extension } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { mergeOfficeTiptapExtensions } from '../src/internal/features/work/editors/office-tiptap-extensions';

describe('TipTap editor extension composition', () => {
  test('keeps built-in and host extensions in deterministic order', () => {
    const builtIn = Extension.create({ name: 'builtInBehavior' });
    const host = Extension.create({ name: 'hostBehavior' });

    expect(
      mergeOfficeTiptapExtensions('DocumentEditor', [builtIn], [host]).map(
        ({ name }) => name,
      ),
    ).toEqual(['builtInBehavior', 'hostBehavior']);
  });

  test('rejects a host extension that overrides a built-in name', () => {
    const builtIn = Extension.create({ name: 'paragraphBehavior' });
    const host = Extension.create({ name: 'paragraphBehavior' });

    expect(() =>
      mergeOfficeTiptapExtensions('DocumentEditor', [builtIn], [host]),
    ).toThrow(
      'DocumentEditor extension "paragraphBehavior" is already registered.',
    );
  });

  test('rejects duplicate names inside the host extension list', () => {
    const first = Extension.create({ name: 'hostBehavior' });
    const second = Extension.create({ name: 'hostBehavior' });

    expect(() =>
      mergeOfficeTiptapExtensions('MarkdownEditor', [], [first, second]),
    ).toThrow('MarkdownEditor extension "hostBehavior" is already registered.');
  });
});
