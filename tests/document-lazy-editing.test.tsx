import { type Editor, Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { type DocumentContent, importOfficeFile } from '../src/core';
import { DocumentEditor } from '../src/react';
import {
  createSchemaValidatedWorkDocumentModel,
  documentModelForContent,
  documentModelHasTrustedInitialIntegrityFeatures,
} from '../src/internal/features/work/work-document-model';
import { documentLazyHtmlProjection } from '../src/internal/features/work/work-document-lazy-model';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

test('edits and round-trips the hydrated end of a lazy document', async () => {
  let editor: Editor | null = null;
  let changed: DocumentContent | null = null;
  let changeCalls = 0;
  const captureEditor = Extension.create({
    name: 'captureLazyDocumentEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const root = lazyEditingRoot();
  const html = `<section data-document-section="true" data-section-id="document-section-1">${Array.from(
    { length: 12 },
    (_, index) => `<p>paragraph ${index + 1}</p>`,
  ).join('')}</section>`;
  const model = createSchemaValidatedWorkDocumentModel(html, root, {
    initialIntegrityFeatures: 0,
  });

  render(
    <DocumentEditor
      content={{
        type: 'document',
        html,
        model,
        pageSize: 'a4',
      }}
      extensions={[captureEditor]}
      onChange={(content) => {
        changeCalls += 1;
        changed = content;
      }}
      theme="light"
    />,
  );

  const surface = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(editor).not.toBeNull());
  fireEvent.keyDown(surface, { ctrlKey: true, key: 'End' });
  await waitFor(() =>
    expect(editor?.state.doc.textContent).toContain('paragraph 12'),
  );
  expect(documentLazyHtmlProjection(model)).not.toBeNull();

  editor?.commands.insertContent(' verified');
  editor?.commands.insertContent('-lazy');
  editor?.commands.insertContent('-edit');
  await waitFor(() => expect(changed?.html).toContain('verified-lazy-edit'));
  if (!changed) throw new Error('Expected a controlled document update.');

  expect(changed.html).not.toContain('documentLazyBlock');
  expect(changed.html).not.toContain('data-section-page-size');
  expect(changeCalls).toBe(1);
  expect(
    documentNodeContainsType(changed.model?.root, 'documentLazyBlock'),
  ).toBe(false);
  expect(documentModelForContent(changed)).toBe(changed.model);
  expect(
    documentModelHasTrustedInitialIntegrityFeatures(changed.model ?? null),
  ).toBe(true);

  editor?.commands.insertContent(' second-edit');
  await waitFor(() => expect(changed?.html).toContain('second-edit'));
  expect(changeCalls).toBe(2);

  const blob = await createDocxBlob(changed);
  const reopened = await importOfficeFile(
    new File([blob], 'lazy-document-edited.docx', { type: blob.type }),
  );
  expect(reopened.content).toMatchObject({
    type: 'document',
    html: expect.stringContaining('verified-lazy-edit'),
  });
});

function lazyEditingRoot(): WorkDocumentNode {
  return windowDocumentModel(
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
      tableRowSize: 2,
      tableRowThreshold: 4,
      trustedIntegrityFeatures: 0,
    },
  );
}

function documentNodeContainsType(
  node: WorkDocumentNode | null | undefined,
  type: string,
): boolean {
  if (!node) return false;
  if (node.type === type) return true;
  return (node.content ?? []).some((child) =>
    documentNodeContainsType(child, type),
  );
}
