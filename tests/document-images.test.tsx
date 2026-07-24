import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { DocumentPictureRibbon } from '../src/internal/features/work/editors/document-picture-ribbon';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentImageAlternativeText,
  documentImageLayoutOptions,
  normalizeDocumentImageLayoutOptions,
} from '../src/internal/features/work/work-document-image-layout';
import { measureDocumentLayoutBlocks } from '../src/internal/features/work/work-document-pagination';

const pixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC' +
  'AAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=';

test('keeps typed image layout in the TipTap model and live node view', () => {
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="Plan" width="120" height="80"`,
      ' data-office-image-layout="square"',
      ' data-office-image-alignment="right"',
      ' data-office-image-wrap-distance="5">',
    ].join(''),
  );
  selectFirstImage(editor);

  expect(documentImageLayoutOptions(editor)).toEqual({
    layout: 'square',
    alignment: 'right',
    wrapDistance: 5,
  });
  expect(editor.getJSON()).toMatchObject({
    type: 'doc',
    content: [
      {
        type: 'image',
        attrs: {
          layout: 'square',
          alignment: 'right',
          wrapDistance: 5,
        },
      },
      {
        type: 'paragraph',
      },
    ],
  });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  const image = container?.querySelector('img');
  expect(container?.dataset.officeImageLayout).toBe('square');
  expect(container?.dataset.officeImageAlignment).toBe('right');
  expect(
    container?.style.getPropertyValue('--work-document-image-wrap-distance'),
  ).toBe('5mm');
  expect(image?.dataset.officeImageLayout).toBe('square');

  expect(
    normalizeDocumentImageLayoutOptions({
      layout: 'unsupported',
      alignment: 'outside',
      wrapDistance: 99,
    }),
  ).toEqual({
    layout: 'inline',
    alignment: 'center',
    wrapDistance: 25,
  });
  editor.destroy();
});

test('reserves and observes floating image height during pagination', () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      `<img src="${pixelPng}" alt="Plan" width="120" height="80"`,
      ' data-office-image-layout="square"',
      ' data-office-image-alignment="left">',
      '<p>Text wraps beside the image.</p>',
      '</section>',
    ].join(''),
  });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  if (!container) throw new Error('Expected an image resize container.');
  Object.defineProperty(container, 'offsetHeight', {
    configurable: true,
    value: 80,
  });

  const imageBlock = measureDocumentLayoutBlocks(editor).blocks.find(
    ({ element }) => element === container,
  );

  expect(imageBlock?.block.height).toBe(80);
  expect(imageBlock?.block.keepTogether).toBe(true);
  expect(imageBlock?.observeResize).toBe(true);
  editor.destroy();
});

test('offers a contextual picture ribbon with typed layout and alt-text actions', async () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Original" width="120" height="80">`,
  );
  selectFirstImage(editor);
  const view = render(<DocumentPictureRibbon editor={editor} />);

  fireEvent.click(screen.getByRole('button', { name: '四周环绕' }));
  fireEvent.click(screen.getByRole('button', { name: '右对齐' }));

  expect(documentImageLayoutOptions(editor)).toMatchObject({
    layout: 'square',
    alignment: 'right',
  });

  fireEvent.click(screen.getByRole('button', { name: '替代文字' }));
  fireEvent.change(screen.getByRole('textbox', { name: '图片替代文字' }), {
    target: { value: '季度趋势图' },
  });
  fireEvent.click(screen.getByRole('button', { name: '保存' }));
  await waitFor(() => {
    expect(documentImageAlternativeText(editor)).toBe('季度趋势图');
  });
  expect(editor.getHTML()).toContain('alt="季度趋势图"');

  view.unmount();
  editor.destroy();
});

test('round-trips supported floating image anchors through DOCX', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document')
    throw new Error('Expected a document artifact.');
  artifact.content.html = [
    `<img src="${pixelPng}" alt="Right chart" width="120" height="80"`,
    ' data-office-image-layout="square"',
    ' data-office-image-alignment="right"',
    ' data-office-image-wrap-distance="5">',
    `<img src="${pixelPng}" alt="Centered diagram" width="140" height="90"`,
    ' data-office-image-layout="topBottom"',
    ' data-office-image-alignment="center"',
    ' data-office-image-wrap-distance="2">',
  ].join('');

  const blob = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml = await archive.file('word/document.xml')?.async('string');

  expect(documentXml).toBeDefined();
  expect(documentXml?.match(/<wp:anchor\b/g)).toHaveLength(2);
  expect(documentXml).toContain('<wp:wrapSquare');
  expect(documentXml).toContain('<wp:wrapTopAndBottom');
  expect(documentXml).toContain('<wp:align>right</wp:align>');
  expect(documentXml).toContain('<wp:align>center</wp:align>');
  expect(documentXml).toContain('distR="180000"');
  expect(documentXml).toContain('distT="72000"');

  const imported = await importOfficeFile(
    new File([blob], 'floating-images.docx', { type: blob.type }),
  );
  expect(imported.content.type).toBe('document');
  if (imported.content.type !== 'document')
    throw new Error('Expected an imported document artifact.');
  expect(imported.content.html).toContain('data-office-image-layout="square"');
  expect(imported.content.html).toContain(
    'data-office-image-alignment="right"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-wrap-distance="5"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-layout="topBottom"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-alignment="center"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-wrap-distance="2"',
  );
});

function createImageEditor(content: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content,
  });
}

function selectFirstImage(editor: Editor): void {
  let position: number | null = null;
  editor.state.doc.descendants((node, offset) => {
    if (position === null && node.type.name === 'image') position = offset;
  });
  if (position === null) throw new Error('Expected an image node.');
  editor.commands.setNodeSelection(position);
}
