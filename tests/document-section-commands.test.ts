import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { updateDocumentMirrorMargins } from '../src/internal/features/work/work-document-page-margins';
import {
  measureDocumentLayoutBlocks,
  measureDocumentLayoutBlocksIncrementally,
} from '../src/internal/features/work/work-document-pagination';
import {
  activeDocumentSection,
  documentSectionById,
} from '../src/internal/features/work/work-document-section-editor';
import { layoutOfficeDocumentInJavaScript } from '../src/internal/kernel/office-kernel-fallback';
import { OFFICE_KERNEL_PROTOCOL_VERSION } from '../src/internal/kernel/office-kernel-protocol';

describe('document section commands', () => {
  test('yields between large-document measurement blocks without changing layout input', async () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>First paragraph</p><p>Second paragraph</p><p>Third paragraph</p>',
        '</section>',
        '<section data-document-section="true" data-section-id="section-2" data-section-page-size="legal">',
        '<p>Fourth paragraph</p>',
        '</section>',
      ].join(''),
    });
    let checkpoints = 0;

    const incremental = await measureDocumentLayoutBlocksIncrementally(
      editor,
      null,
      0,
      new Map(),
      1_000_000,
      {
        checkpoint: async () => {
          checkpoints += 1;
        },
      },
    );
    const synchronous = measureDocumentLayoutBlocks(editor);

    expect(checkpoints).toBe(4);
    expect(
      incremental.blocks.map(({ block, from, section, to }) => ({
        block,
        from,
        section,
        to,
      })),
    ).toEqual(
      synchronous.blocks.map(({ block, from, section, to }) => ({
        block,
        from,
        section,
        to,
      })),
    );
    expect(incremental.pageStyles).toEqual(synchronous.pageStyles);
    expect(incremental.measuredBlockCount).toBe(synchronous.measuredBlockCount);
    expect(incremental.reusedBlockCount).toBe(synchronous.reusedBlockCount);
    expect(incremental.unsupportedLayout).toBe(synchronous.unsupportedLayout);

    editor.destroy();
  });

  test('stops incremental measurement when the active pagination run is aborted', async () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>First paragraph</p><p>Second paragraph</p><p>Third paragraph</p>',
        '</section>',
      ].join(''),
    });
    const controller = new AbortController();
    let checkpoints = 0;

    await expect(
      measureDocumentLayoutBlocksIncrementally(
        editor,
        null,
        0,
        new Map(),
        1_000_000,
        {
          checkpoint: async () => {
            checkpoints += 1;
            if (checkpoints === 2) controller.abort();
          },
          signal: controller.signal,
        },
      ),
    ).rejects.toMatchObject({ name: 'AbortError' });
    expect(checkpoints).toBe(2);

    editor.destroy();
  });

  test('paginates mixed section page geometry instead of rejecting it', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1" data-section-break-after="continuous" data-section-page-size="a4"><p>First</p></section>',
        '<section data-document-section="true" data-section-id="section-2" data-section-page-size="legal"><p>Second</p></section>',
      ].join(''),
    });

    const snapshot = measureDocumentLayoutBlocks(editor);

    expect(snapshot.unsupportedLayout).toBe(false);
    expect(snapshot.pageStyles).toHaveLength(2);
    expect(snapshot.blocks.map((block) => block.block.pageStyleId)).toEqual([
      'document-page-style-1',
      'document-page-style-2',
    ]);
    expect(snapshot.blocks[1]?.section?.layout.pageSize).toBe('legal');

    const layout = layoutOfficeDocumentInJavaScript({
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'layout',
      requestId: 1,
      revision: 1,
      documentRevision: 1,
      startPageIndex: 0,
      page: snapshot.pageStyles[0].page,
      pageStyles: snapshot.pageStyles,
      blocks: snapshot.blocks.map(({ block }) => block),
    });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.page.height)).toEqual(
      snapshot.pageStyles.map(({ page }) => page.height),
    );

    editor.destroy();
  });

  test('inserts, updates, and merges sections through TipTap commands', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>First</p><p>Second</p>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(9);

    expect(editor.commands.insertDocumentSection('continuous')).toBe(true);
    expect(editor.state.doc.childCount).toBe(2);
    expect(activeDocumentSection(editor)?.index).toBe(1);

    const active = activeDocumentSection(editor);
    expect(active).not.toBeNull();
    if (!active) throw new Error('Expected an active document section.');
    expect(
      editor.commands.updateActiveDocumentSection({
        ...active.layout,
        orientation: 'landscape',
      }),
    ).toBe(true);
    expect(activeDocumentSection(editor)?.layout.orientation).toBe('landscape');

    expect(editor.commands.mergeDocumentSectionWithPrevious()).toBe(true);
    expect(editor.state.doc.childCount).toBe(1);
    expect(editor.getText()).toContain('First');
    expect(editor.getText()).toContain('Second');

    editor.destroy();
  });

  test('updates document-wide margin settings across sections in one undo step', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1"><p>First</p></section>',
        '<section data-document-section="true" data-section-id="section-2"><p>Second</p></section>',
      ].join(''),
    });
    editor.commands.setTextSelection(editor.state.doc.content.size - 2);
    const active = activeDocumentSection(editor);
    expect(active?.id).toBe('section-2');
    if (!active) throw new Error('Expected the second document section.');

    expect(
      editor.commands.updateActiveDocumentSection(
        updateDocumentMirrorMargins(active.layout, true),
      ),
    ).toBe(true);
    expect(
      documentSectionById(editor, 'section-1')?.layout.pageMargins,
    ).toMatchObject({ mirrorMargins: true, gutterAtTop: false });
    expect(
      documentSectionById(editor, 'section-2')?.layout.pageMargins,
    ).toMatchObject({ mirrorMargins: true, gutterAtTop: false });

    expect(editor.commands.undo()).toBe(true);
    expect(
      documentSectionById(editor, 'section-1')?.layout.pageMargins,
    ).toBeUndefined();
    expect(
      documentSectionById(editor, 'section-2')?.layout.pageMargins,
    ).toBeUndefined();
    expect(editor.commands.undo()).toBe(false);

    editor.destroy();
  });
});
