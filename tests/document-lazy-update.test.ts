import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { documentTextStatistics } from '../src/internal/features/work/editors/document-editor-support';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { prepareLazyDocumentEditorSource } from '../src/internal/features/work/work-document-lazy-model';
import {
  materializeLazyDocumentUpdate,
  transferLazyDocumentTextStatistics,
} from '../src/internal/features/work/work-document-lazy-update';
import { createSchemaValidatedWorkDocumentModel } from '../src/internal/features/work/work-document-model';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

test('patches controlled HTML from a text-only lazy chunk edit', () => {
  const html = `<section data-document-section="true" data-section-id="document-section-1">${Array.from(
    { length: 12 },
    (_, index) => `<p>paragraph ${index + 1}</p>`,
  ).join('')}</section>`;
  const root = windowDocumentModel(
    {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: { id: 'document-section-1' },
          content: Array.from({ length: 12 }, (_, index) => ({
            type: 'paragraph',
            content: [{ type: 'text', text: `paragraph ${index + 1}` }],
          })),
        },
      ],
    },
    {
      blockSize: 2,
      blockThreshold: 4,
      trustedIntegrityFeatures: 0,
    },
  );
  const model = createSchemaValidatedWorkDocumentModel(html, root, {
    initialIntegrityFeatures: 0,
  });
  const content: WorkDocumentContent = {
    type: 'document',
    html,
    model,
    pageSize: 'a4',
  };
  const prepared = prepareLazyDocumentEditorSource(model, true, html);
  if (!prepared) throw new Error('Expected a lazy editor source.');
  const editor = new Editor({
    extensions: createWorkDocumentExtensions({ getContent: () => content }),
    content: prepared.root,
  });
  editor.commands.setTextSelection(editor.state.doc.content.size - 1);
  const before = editor.state.doc;
  const statisticsBefore = documentTextStatistics(editor);

  editor.commands.insertContent(' up');
  editor.commands.insertContent('dat');
  editor.commands.insertContent('ed');
  expect(transferLazyDocumentTextStatistics(before, editor.state.doc)).toBe(
    true,
  );
  expect(documentTextStatistics(editor)).toEqual({
    characterCountWithSpaces: statisticsBefore.characterCountWithSpaces + 8,
    characterCountWithoutSpaces:
      statisticsBefore.characterCountWithoutSpaces + 7,
    paragraphCount: statisticsBefore.paragraphCount,
    wordCount: statisticsBefore.wordCount + 1,
  });
  const update = materializeLazyDocumentUpdate(
    before,
    editor.state.doc,
    editor.getJSON(),
    model,
  );

  expect(update.html).toContain('paragraph 12 updated');
  expect(update.html).not.toContain('data-section-page-size');
  editor.destroy();
});
