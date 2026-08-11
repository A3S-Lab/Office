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
import {
  createDocumentPicturePropertiesDraft,
  documentPicturePropertiesErrors,
  documentPicturePropertyChanges,
  withDocumentPictureAspectRatioLock,
  withDocumentPictureDimension,
} from '../src/internal/features/work/editors/document-picture-properties-dialog-model';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentImageAlternativeText,
  documentImageLayoutOptions,
  documentImageProperties,
  normalizeDocumentImageLayoutOptions,
} from '../src/internal/features/work/work-document-image-layout';
import {
  documentImageWrapContourCss,
  normalizeDocumentImageWrapContour,
} from '../src/internal/features/work/work-document-image-wrap-contour';
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
    wrapSide: 'bothSides',
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
          wrapSide: 'bothSides',
          lockAspectRatio: true,
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
  expect(container?.dataset.officeImageLockAspectRatio).toBe('true');

  expect(
    normalizeDocumentImageLayoutOptions({
      layout: 'unsupported',
      alignment: 'outside',
      wrapDistance: 99,
      wrapSide: 'unsupported',
    }),
  ).toEqual({
    layout: 'inline',
    alignment: 'center',
    wrapDistance: 25,
    wrapSide: 'bothSides',
  });
  editor.destroy();
});

test('keeps edited image wrap contours in the model and live presentation', () => {
  const polygon = '0,0;0,21600;10800,16200;21600,21600;21600,0;0,0';
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="Contour" width="120" height="80"`,
      ' data-office-image-layout="tight"',
      ' data-office-image-alignment="left"',
      ' data-office-image-wrap-distance="4"',
      ' data-office-image-wrap-side="largest"',
      ` data-office-image-wrap-polygon="${polygon}"`,
      ' data-office-image-wrap-polygon-edited="true">',
    ].join(''),
  );
  selectFirstImage(editor);

  expect(documentImageProperties(editor)).toMatchObject({
    layout: 'tight',
    alignment: 'left',
    wrapDistance: 4,
    wrapSide: 'largest',
    contour: {
      edited: true,
      points: [
        { x: 0, y: 0 },
        { x: 0, y: 21_600 },
        { x: 10_800, y: 16_200 },
        { x: 21_600, y: 21_600 },
        { x: 21_600, y: 0 },
        { x: 0, y: 0 },
      ],
    },
  });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  expect(container?.dataset.officeImageWrapPolygon).toBe(polygon);
  expect(container?.dataset.officeImageWrapPolygonEdited).toBe('true');
  expect(
    container?.style.getPropertyValue('--work-document-image-wrap-contour'),
  ).toBe('polygon(0% 0%, 0% 100%, 50% 75%, 100% 100%, 100% 0%, 0% 0%)');
  expect(editor.getHTML()).toContain(
    '--work-document-image-wrap-contour: polygon(',
  );

  expect(
    normalizeDocumentImageWrapContour({
      points: [
        { x: 0, y: 0 },
        { x: 21_600, y: 0 },
      ],
    }),
  ).toBeNull();
  expect(
    documentImageWrapContourCss({
      edited: false,
      points: [
        { x: 0, y: 0 },
        { x: 10_800, y: 21_600 },
        { x: 21_600, y: 0 },
      ],
    }),
  ).toBe('polygon(0% 0%, 50% 100%, 100% 0%)');
  editor.destroy();
});

test('commits picture size, layout, and alternative text as one undoable update', () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Original" title="Original title" width="120" height="80">`,
  );
  selectFirstImage(editor);
  const originalHtml = editor.getHTML();
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });

  expect(
    editor.commands.setDocumentImageProperties({
      width: 240,
      height: 150,
      lockAspectRatio: false,
      layout: 'square',
      alignment: 'right',
      wrapDistance: 8,
      wrapSide: 'largest',
      alternativeText: '季度趋势图',
      position: {
        horizontalOffset: -6.5,
        verticalOffset: 4,
        horizontalReference: 'margin',
        verticalReference: 'page',
      },
      crop: { top: 12.5, right: 5, bottom: 2.25, left: 10 },
    }),
  ).toBe(true);

  expect(updateCount).toBe(1);
  expect(documentImageProperties(editor)).toEqual({
    width: 240,
    height: 150,
    lockAspectRatio: false,
    layout: 'square',
    alignment: 'right',
    wrapDistance: 8,
    wrapSide: 'largest',
    alternativeText: '季度趋势图',
    position: {
      horizontalOffset: -6.5,
      verticalOffset: 4,
      horizontalReference: 'margin',
      verticalReference: 'page',
    },
    crop: { top: 12.5, right: 5, bottom: 2.25, left: 10 },
  });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  expect(container?.dataset.officeImageLockAspectRatio).toBe('false');
  expect(container?.dataset.officeImageHorizontalOffset).toBe('-6.5');
  expect(container?.dataset.officeImageVerticalOffset).toBe('4');
  expect(container?.dataset.officeImageWrapSide).toBe('largest');
  expect(container?.dataset.officeImageCropTop).toBe('12.5');
  expect(container?.dataset.officeImageCropLeft).toBe('10');
  expect(
    container?.style.getPropertyValue('--work-document-image-crop-scale-x'),
  ).toBe('1.18');
  expect(
    container?.style.getPropertyValue('--work-document-image-crop-translate-x'),
  ).toBe('-11.76%');
  expect(
    container?.style.getPropertyValue(
      '--work-document-image-horizontal-offset',
    ),
  ).toBe('-6.5mm');
  expect(editor.getHTML()).toContain(
    '--work-document-image-horizontal-offset: -6.5mm',
  );
  expect(editor.getHTML()).toContain(
    '--work-document-image-vertical-offset: 4mm',
  );
  expect(editor.getHTML()).toContain('--work-document-image-crop-top: 12.5%');
  expect(editor.getHTML()).toContain(
    '--work-document-image-crop-scale-x: 1.18',
  );
  expect(container?.querySelector('img')?.style.width).toBe('240px');
  expect(container?.querySelector('img')?.style.height).toBe('150px');

  expect(editor.commands.undo()).toBe(true);
  expect(editor.getHTML()).toBe(originalHtml);
  editor.destroy();
});

test('couples picture dimensions only while aspect ratio is locked', () => {
  const initial = createDocumentPicturePropertiesDraft({
    properties: {
      width: 120,
      height: 80,
      lockAspectRatio: true,
      layout: 'inline',
      alignment: 'center',
      wrapDistance: 3,
      alternativeText: 'Original',
    },
  });
  const wider = withDocumentPictureDimension(initial, 'width', '6');
  expect(wider).toMatchObject({ width: '6', height: '4' });

  const unlocked = withDocumentPictureAspectRatioLock(wider, false);
  const taller = withDocumentPictureDimension(unlocked, 'height', '5');
  expect(taller).toMatchObject({ width: '6', height: '5' });

  const relocked = withDocumentPictureAspectRatioLock(taller, true);
  const resized = withDocumentPictureDimension(relocked, 'width', '2.4');
  expect(resized).toMatchObject({ width: '2.4', height: '2' });
});

test('preserves untouched imported dimensions in picture property changes', () => {
  const initial = createDocumentPicturePropertiesDraft({
    properties: {
      width: null,
      height: 79.125,
      lockAspectRatio: true,
      layout: 'square',
      alignment: 'left',
      wrapDistance: 4.25,
      alternativeText: 'Imported title',
    },
    renderedWidth: 241.375,
    renderedHeight: 79.125,
  });
  const changed = { ...initial, alternativeText: 'Accessible title' };

  expect(documentPicturePropertyChanges(initial, changed)).toEqual({
    alternativeText: 'Accessible title',
  });
  expect(
    documentPicturePropertiesErrors({
      ...initial,
      width: '0.001',
      layout: 'topBottom',
      wrapDistance: '30',
    }),
  ).toEqual({
    width: '请输入 0.01 到 55.87 之间的厘米数。',
    height: null,
    wrapDistance: '请输入 0 到 25 之间的毫米数。',
    horizontalOffset: null,
    verticalOffset: null,
    crop: null,
  });
});

test('validates and commits precise floating-image anchor positions', () => {
  const initial = createDocumentPicturePropertiesDraft({
    properties: {
      width: 120,
      height: 80,
      lockAspectRatio: true,
      layout: 'square',
      alignment: 'left',
      wrapDistance: 3,
      alternativeText: 'Positioned image',
    },
  });
  const positioned = {
    ...initial,
    precisePosition: true,
    horizontalOffset: '-12.5',
    verticalOffset: '7.25',
    horizontalReference: 'page' as const,
    verticalReference: 'margin' as const,
  };
  expect(documentPicturePropertiesErrors(positioned)).toMatchObject({
    horizontalOffset: null,
    verticalOffset: null,
  });
  expect(documentPicturePropertyChanges(initial, positioned)).toEqual({
    position: {
      horizontalOffset: -12.5,
      verticalOffset: 7.25,
      horizontalReference: 'page',
      verticalReference: 'margin',
    },
  });
  expect(
    documentPicturePropertiesErrors({
      ...positioned,
      horizontalOffset: '-600',
    }).horizontalOffset,
  ).not.toBeNull();
});

test('validates and commits four-edge image crop geometry', () => {
  const initial = createDocumentPicturePropertiesDraft({
    properties: {
      width: 120,
      height: 80,
      lockAspectRatio: true,
      layout: 'inline',
      alignment: 'center',
      wrapDistance: 3,
      alternativeText: 'Crop image',
    },
  });
  const cropped = {
    ...initial,
    cropTop: '12.5',
    cropRight: '5',
    cropBottom: '2.25',
    cropLeft: '10',
  };
  expect(documentPicturePropertiesErrors(cropped).crop).toBeNull();
  expect(documentPicturePropertyChanges(initial, cropped)).toEqual({
    crop: { top: 12.5, right: 5, bottom: 2.25, left: 10 },
  });
  expect(
    documentPicturePropertiesErrors({
      ...cropped,
      cropRight: '90',
    }).crop,
  ).not.toBeNull();
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

test('offers a contextual picture ribbon with one coherent properties workflow', async () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Original" width="120" height="80" data-office-image-wrap-distance="7">`,
  );
  selectFirstImage(editor);
  const view = render(<DocumentPictureRibbon editor={editor} />);
  const wrapDistance = screen.getByRole('combobox', {
    name: '图片与文字距离',
  });
  expect(wrapDistance).toHaveTextContent('7 毫米');
  expect(wrapDistance).toBeDisabled();

  fireEvent.click(screen.getByRole('button', { name: '四周环绕' }));
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(documentImageLayoutOptions(editor).layout).toBe('square');
  expect(wrapDistance).toBeEnabled();
  for (const [label, alignment] of [
    ['左对齐', 'left'],
    ['居中', 'center'],
    ['右对齐', 'right'],
  ] as const) {
    fireEvent.click(screen.getByRole('button', { name: label }));
    expect(documentImageLayoutOptions(editor).alignment).toBe(alignment);
  }
  fireEvent.click(screen.getByRole('combobox', { name: '图片与文字距离' }));
  fireEvent.click(screen.getByRole('option', { name: '10 毫米' }));
  expect(documentImageLayoutOptions(editor).wrapDistance).toBe(10);

  fireEvent.click(screen.getByRole('button', { name: '紧密环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('tight');
  fireEvent.click(screen.getByRole('button', { name: '穿越环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('through');
  fireEvent.click(screen.getByRole('button', { name: '上下环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('topBottom');
  fireEvent.click(screen.getByRole('button', { name: '嵌入文字' }));
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(documentImageLayoutOptions(editor).layout).toBe('inline');
  expect(wrapDistance).toBeDisabled();

  const propertiesButton = screen.getByRole('button', {
    name: '图片属性',
  });
  propertiesButton.focus();
  fireEvent.click(propertiesButton);
  expect(screen.getByRole('dialog', { name: '图片属性' })).toBeInTheDocument();
  expect(screen.getByRole('textbox', { name: '图片宽度（厘米）' })).toHaveValue(
    '3.18',
  );
  expect(screen.getByRole('textbox', { name: '图片高度（厘米）' })).toHaveValue(
    '2.12',
  );
  expect(screen.getByRole('checkbox', { name: '锁定纵横比' })).toBeChecked();

  fireEvent.change(screen.getByRole('textbox', { name: '图片宽度（厘米）' }), {
    target: { value: '6' },
  });
  expect(screen.getByRole('textbox', { name: '图片高度（厘米）' })).toHaveValue(
    '4',
  );
  fireEvent.click(screen.getByRole('checkbox', { name: '锁定纵横比' }));
  fireEvent.change(screen.getByRole('textbox', { name: '图片高度（厘米）' }), {
    target: { value: '5' },
  });
  fireEvent.click(screen.getByRole('radio', { name: '紧密环绕' }));
  fireEvent.click(screen.getByRole('combobox', { name: '图片文字环绕侧' }));
  fireEvent.click(screen.getByRole('option', { name: '较宽一侧' }));
  fireEvent.click(screen.getByRole('radio', { name: '右对齐' }));
  fireEvent.change(
    screen.getByRole('textbox', { name: '图片与文字距离（毫米）' }),
    { target: { value: '8' } },
  );
  fireEvent.click(screen.getByRole('checkbox', { name: '使用精确图片位置' }));
  fireEvent.change(
    screen.getByRole('textbox', { name: '图片水平偏移（毫米）' }),
    { target: { value: '-12.5' } },
  );
  fireEvent.change(
    screen.getByRole('textbox', { name: '图片垂直偏移（毫米）' }),
    { target: { value: '7.25' } },
  );
  fireEvent.click(screen.getByRole('combobox', { name: '水平相对于' }));
  fireEvent.click(screen.getByRole('option', { name: '页面' }));
  fireEvent.click(screen.getByRole('combobox', { name: '垂直相对于' }));
  fireEvent.click(screen.getByRole('option', { name: '页边距' }));
  for (const [name, value] of [
    ['图片上方裁剪（百分比）', '12.5'],
    ['图片右侧裁剪（百分比）', '5'],
    ['图片下方裁剪（百分比）', '2.25'],
    ['图片左侧裁剪（百分比）', '10'],
  ] as const) {
    fireEvent.change(screen.getByRole('textbox', { name }), {
      target: { value },
    });
  }
  fireEvent.change(screen.getByRole('textbox', { name: '图片替代文字' }), {
    target: { value: '季度趋势图' },
  });

  const beforeProperties = editor.getHTML();
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  await waitFor(() => {
    expect(screen.queryByRole('dialog', { name: '图片属性' })).toBeNull();
    expect(editor.isActive('image')).toBe(true);
    expect(propertiesButton).toHaveFocus();
  });
  expect(updateCount).toBe(1);
  expect(documentImageProperties(editor)).toMatchObject({
    width: expect.closeTo(226.77, 1),
    height: expect.closeTo(188.98, 1),
    lockAspectRatio: false,
    layout: 'tight',
    alignment: 'right',
    wrapDistance: 8,
    wrapSide: 'largest',
    alternativeText: '季度趋势图',
    position: {
      horizontalOffset: -12.5,
      verticalOffset: 7.25,
      horizontalReference: 'page',
      verticalReference: 'margin',
    },
    crop: { top: 12.5, right: 5, bottom: 2.25, left: 10 },
  });
  expect(editor.getHTML()).toContain('alt="季度趋势图"');

  expect(editor.commands.undo()).toBe(true);
  expect(editor.getHTML()).toBe(beforeProperties);

  selectFirstImage(editor);
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  fireEvent.click(propertiesButton);
  fireEvent.change(screen.getByRole('textbox', { name: '图片替代文字' }), {
    target: { value: '不应保存' },
  });
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  await waitFor(() => {
    expect(editor.isActive('image')).toBe(true);
    expect(propertiesButton).toHaveFocus();
  });
  expect(documentImageAlternativeText(editor)).toBe('Original');

  fireEvent.click(propertiesButton);
  fireEvent.change(screen.getByRole('textbox', { name: '图片替代文字' }), {
    target: { value: '仍不应保存' },
  });
  fireEvent.keyDown(screen.getByRole('textbox', { name: '图片替代文字' }), {
    key: 'Escape',
  });
  await waitFor(() => {
    expect(editor.isActive('image')).toBe(true);
    expect(propertiesButton).toHaveFocus();
  });
  expect(documentImageAlternativeText(editor)).toBe('Original');

  selectFirstImage(editor);
  fireEvent.click(screen.getByRole('button', { name: '删除图片' }));
  expect(editor.getJSON()).not.toMatchObject({
    content: expect.arrayContaining([
      expect.objectContaining({ type: 'image' }),
    ]),
  });

  view.unmount();
  editor.destroy();
});

test('disables stale picture commands when the image selection is gone', () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Plan" width="120" height="80"><p>After</p>`,
  );
  editor.commands.setTextSelection(textPosition(editor, 'After'));
  const view = render(<DocumentPictureRibbon editor={editor} />);

  expect(screen.getByRole('button', { name: '四周环绕' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '紧密环绕' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '穿越环绕' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '图片属性' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '删除图片' })).toBeDisabled();

  selectFirstImage(editor);
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(screen.getByRole('button', { name: '四周环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '紧密环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '穿越环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '图片属性' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '删除图片' })).toBeEnabled();

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
    ' data-office-image-wrap-distance="5"',
    ' data-office-image-wrap-side="right">',
    `<img src="${pixelPng}" alt="Centered diagram" width="140" height="90"`,
    ' data-office-image-layout="topBottom"',
    ' data-office-image-alignment="center"',
    ' data-office-image-wrap-distance="2">',
    `<img src="${pixelPng}" alt="Precisely positioned" width="100" height="70"`,
    ' data-office-image-layout="square"',
    ' data-office-image-horizontal-offset="-12.5"',
    ' data-office-image-vertical-offset="7.25"',
    ' data-office-image-horizontal-reference="page"',
    ' data-office-image-vertical-reference="margin"',
    ' data-office-image-crop-top="12.5"',
    ' data-office-image-crop-right="5"',
    ' data-office-image-crop-bottom="2.25"',
    ' data-office-image-crop-left="10">',
    `<img src="${pixelPng}" alt="Tight contour" width="110" height="75"`,
    ' data-office-image-layout="tight"',
    ' data-office-image-alignment="left"',
    ' data-office-image-wrap-distance="4"',
    ' data-office-image-wrap-side="largest"',
    ' data-office-image-wrap-polygon="0,0;0,21600;10800,16200;21600,21600;21600,0;0,0"',
    ' data-office-image-wrap-polygon-edited="true">',
    `<img src="${pixelPng}" alt="Through contour" width="105" height="65"`,
    ' data-office-image-layout="through"',
    ' data-office-image-alignment="right"',
    ' data-office-image-wrap-distance="6"',
    ' data-office-image-wrap-side="left"',
    ' data-office-image-wrap-polygon="0,0;0,21600;7200,10800;21600,21600;21600,0;0,0"',
    ' data-office-image-wrap-polygon-edited="false">',
    `<img src="${pixelPng}" alt="Inline crop" width="90" height="60"`,
    ' data-office-image-crop-bottom="10"',
    ' data-office-image-crop-left="20">',
  ].join('');

  const blob = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml = await archive.file('word/document.xml')?.async('string');

  expect(documentXml).toBeDefined();
  expect(documentXml?.match(/<wp:anchor\b/g)).toHaveLength(5);
  expect(documentXml).toMatch(/<wp:wrapSquare\b(?=[^>]*wrapText="right")/);
  expect(documentXml).toContain('<wp:wrapTopAndBottom');
  expect(documentXml).toMatch(
    /<wp:wrapTight\b(?=[^>]*wrapText="largest")(?=[^>]*distL="144000")(?=[^>]*distR="144000")/,
  );
  expect(documentXml).toMatch(
    /<wp:wrapThrough\b(?=[^>]*wrapText="left")(?=[^>]*distL="216000")(?=[^>]*distR="216000")/,
  );
  expect(documentXml).toContain(
    '<wp:wrapPolygon edited="1"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="10800" y="16200"/>',
  );
  expect(documentXml).toContain(
    '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="7200" y="10800"/>',
  );
  expect(documentXml).toContain('<wp:align>right</wp:align>');
  expect(documentXml).toContain('<wp:align>center</wp:align>');
  expect(documentXml).toContain('distR="180000"');
  expect(documentXml).toContain('distT="72000"');
  expect(documentXml).toMatch(
    /<wp:positionH relativeFrom="page"><wp:posOffset>-450000<\/wp:posOffset>/,
  );
  expect(documentXml).toMatch(
    /<wp:positionV relativeFrom="margin"><wp:posOffset>261000<\/wp:posOffset>/,
  );
  expect(documentXml).toMatch(
    /<a:srcRect\b(?=[^>]*t="12500")(?=[^>]*r="5000")(?=[^>]*b="2250")(?=[^>]*l="10000")/,
  );
  expect(documentXml).toMatch(
    /<a:srcRect\b(?=[^>]*b="10000")(?=[^>]*l="20000")/,
  );
  expect(documentXml).not.toContain('__A3S_IMAGE_CROP_');
  expect(documentXml).not.toContain('__A3S_IMAGE_WRAP_');

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
  expect(imported.content.html).toContain(
    'data-office-image-horizontal-offset="-12.5"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-vertical-offset="7.25"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-horizontal-reference="page"',
  );
  expect(imported.content.html).toContain(
    'data-office-image-vertical-reference="margin"',
  );
  expect(imported.content.html).toContain('data-office-image-crop-top="12.5"');
  expect(imported.content.html).toContain('data-office-image-crop-right="5"');
  expect(imported.content.html).toContain(
    'data-office-image-crop-bottom="2.25"',
  );
  expect(imported.content.html).toContain('data-office-image-crop-left="10"');
  expect(imported.content.html).toContain(
    '--work-document-image-crop-scale-x: 1.18',
  );
  const inlineCrop = new DOMParser()
    .parseFromString(imported.content.html, 'text/html')
    .querySelector<HTMLImageElement>('img[alt="Inline crop"]');
  expect(inlineCrop?.dataset.officeImageCropBottom).toBe('10');
  expect(inlineCrop?.dataset.officeImageCropLeft).toBe('20');
  const importedHtml = new DOMParser().parseFromString(
    imported.content.html,
    'text/html',
  );
  const tightContour = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Tight contour"]',
  );
  expect(tightContour?.dataset.officeImageLayout).toBe('tight');
  expect(tightContour?.dataset.officeImageWrapSide).toBe('largest');
  expect(tightContour?.dataset.officeImageWrapPolygon).toBe(
    '0,0;0,21600;10800,16200;21600,21600;21600,0;0,0',
  );
  expect(tightContour?.dataset.officeImageWrapPolygonEdited).toBe('true');
  expect(
    tightContour?.style.getPropertyValue('--work-document-image-wrap-contour'),
  ).toContain('50% 75%');
  const throughContour = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Through contour"]',
  );
  expect(throughContour?.dataset.officeImageLayout).toBe('through');
  expect(throughContour?.dataset.officeImageWrapSide).toBe('left');
  expect(throughContour?.dataset.officeImageWrapPolygon).toBe(
    '0,0;0,21600;7200,10800;21600,21600;21600,0;0,0',
  );
  const alignedImage = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Right chart"]',
  );
  expect(alignedImage?.dataset.officeImageHorizontalOffset).toBeUndefined();
  expect(alignedImage?.dataset.officeImageVerticalOffset).toBe('0');
  expect(alignedImage?.dataset.officeImageWrapSide).toBe('right');

  const regenerated = await createArtifactBlob(imported);
  const regeneratedArchive = await JSZip.loadAsync(
    await regenerated.arrayBuffer(),
  );
  const regeneratedXml = await regeneratedArchive
    .file('word/document.xml')
    ?.async('string');
  expect(regeneratedXml).toContain('<wp:align>right</wp:align>');
  expect(regeneratedXml).toMatch(/<wp:wrapSquare\b(?=[^>]*wrapText="right")/);
  expect(regeneratedXml).toContain(
    '<wp:positionH relativeFrom="page"><wp:posOffset>-450000</wp:posOffset>',
  );
  expect(regeneratedXml).toMatch(
    /<a:srcRect\b(?=[^>]*t="12500")(?=[^>]*r="5000")(?=[^>]*b="2250")(?=[^>]*l="10000")/,
  );
  expect(regeneratedXml).toMatch(/<wp:wrapTight\b(?=[^>]*wrapText="largest")/);
  expect(regeneratedXml).toMatch(/<wp:wrapThrough\b(?=[^>]*wrapText="left")/);
  expect(regeneratedXml).toContain('<wp:lineTo x="10800" y="16200"/>');
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

function textPosition(editor: Editor, text: string): number {
  let position: number | null = null;
  editor.state.doc.descendants((node, offset) => {
    if (position === null && node.isText && node.text?.includes(text)) {
      position = offset;
    }
    return position === null;
  });
  if (position === null) throw new Error(`Expected text "${text}".`);
  return position;
}
