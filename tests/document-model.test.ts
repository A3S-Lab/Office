import { describe, expect, test } from '@rstest/core';
import {
  createSchemaValidatedWorkDocumentModel,
  createWorkDocumentModel,
  documentModelForContent,
  documentModelHasTrustedInitialIntegrityFeatures,
  resolveWorkDocumentEditorInput,
} from '../src/internal/features/work/work-document-model';
import { syncDocumentContentFromHtml } from '../src/internal/features/work/work-document-section';
import type {
  WorkDocumentContent,
  WorkDocumentNode,
} from '../src/internal/features/work/work-types';

const root = {
  type: 'doc',
  content: [
    {
      type: 'paragraph',
      content: [{ type: 'text', text: 'Structured content' }],
    },
  ],
} satisfies WorkDocumentNode;

describe('structured document model', () => {
  test('creates and resolves a versioned document model', () => {
    const html = '<p>Structured content</p>';
    const model = createWorkDocumentModel(html, root);
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model,
    } satisfies WorkDocumentContent;

    expect(model).toMatchObject({
      schema: 'a3s.office.document',
      version: 1,
      revision: 1,
    });
    expect(documentModelForContent(content)).toEqual(model);
    expect(resolveWorkDocumentEditorInput(content, html)).toEqual({
      model,
      source: root,
      sourceKey: `model:${model.revision}:${model.htmlFingerprint}`,
      revision: model.revision,
    });
  });

  test('increments the model revision without mutating the previous model', () => {
    const html = '<p>Structured content</p>';
    const first = createWorkDocumentModel(html, root);
    const second = createWorkDocumentModel(html, root, first);

    expect(first.revision).toBe(1);
    expect(second.revision).toBe(2);
  });

  test('keeps schema-validation trust process-local', () => {
    const html = '<p>Structured content</p>';
    const model = createSchemaValidatedWorkDocumentModel(html, root, {
      initialIntegrityFeatures: 0,
    });
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model,
    } satisfies WorkDocumentContent;

    expect(documentModelHasTrustedInitialIntegrityFeatures(model)).toBe(true);
    expect(documentModelForContent(content)).toBe(model);

    const clonedModel = structuredClone(model);
    expect(documentModelHasTrustedInitialIntegrityFeatures(clonedModel)).toBe(
      false,
    );
    expect(documentModelForContent({ ...content, model: clonedModel })).toEqual(
      clonedModel,
    );
  });

  test('accepts a persisted legacy HTML fingerprint', () => {
    const html = '<p>Structured content</p>';
    const model = structuredClone(createWorkDocumentModel(html, root));
    model.htmlFingerprint = 'p:1ht88x0';

    expect(
      documentModelForContent({
        type: 'document',
        pageSize: 'a4',
        html,
        model,
      }),
    ).toEqual(model);
  });

  test('keeps the trusted HTML shortcut bound to the exact model snapshot', () => {
    const html = '<p>Structured content</p>';
    const model = createSchemaValidatedWorkDocumentModel(html, root, {
      initialIntegrityFeatures: 0,
    });
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model,
    } satisfies WorkDocumentContent;

    expect(documentModelForContent(content)).toBe(model);
    expect(
      documentModelForContent({ ...content, html: '<p>Host edit</p>' }),
    ).toBeNull();

    model.htmlFingerprint = 'mutated';
    expect(documentModelForContent(content)).toBeNull();
  });

  test('ignores a stale model when a host changes legacy HTML', () => {
    const model = createWorkDocumentModel('<p>Original</p>', root);
    const content = {
      type: 'document',
      pageSize: 'a4',
      html: '<p>Changed by the host</p>',
      model,
    } satisfies WorkDocumentContent;

    expect(documentModelForContent(content)).toBeNull();
    expect(
      resolveWorkDocumentEditorInput(content, '<p>Normalized host HTML</p>'),
    ).toMatchObject({
      model: null,
      source: '<p>Normalized host HTML</p>',
      sourceKey: expect.stringContaining('html:'),
      revision: 0,
    });
  });

  test('keeps a synchronized model when the legacy HTML fallback normalizes differently', () => {
    const html =
      '<section data-page-chrome="{&quot;headerHtml&quot;:&quot;<p>Header</p>&quot;}"><p>Structured content</p></section>';
    const model = createWorkDocumentModel(html, root);
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model,
    } satisfies WorkDocumentContent;

    expect(
      resolveWorkDocumentEditorInput(
        content,
        '<section data-page-chrome="{&quot;headerHtml&quot;:&quot;&lt;p&gt;Header&lt;/p&gt;&quot;}"><p>Structured content</p></section>',
      ),
    ).toEqual({
      model,
      source: root,
      sourceKey: `model:${model.revision}:${model.htmlFingerprint}`,
      revision: model.revision,
    });
  });

  test('drops a stale model when an HTML-only command changes the document', () => {
    const html = '<p>Original</p>';
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model: createWorkDocumentModel(html, root),
    } satisfies WorkDocumentContent;

    const next = syncDocumentContentFromHtml(content, '<p>Changed</p>');

    expect(next.html).toContain('Changed');
    expect(next.model).toBeUndefined();
  });

  test('rejects malformed or excessively deep model trees', () => {
    let deep: WorkDocumentNode = { type: 'paragraph' };
    for (let index = 0; index < 300; index += 1) {
      deep = { type: 'blockquote', content: [deep] };
    }
    const html = '<p>Fallback</p>';
    const content = {
      type: 'document',
      pageSize: 'a4',
      html,
      model: createWorkDocumentModel(html, deep),
    } satisfies WorkDocumentContent;

    expect(documentModelForContent(content)).toBeNull();
  });
});
