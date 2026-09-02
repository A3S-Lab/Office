import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import {
  createDocumentPicturePropertiesDraft,
  documentPicturePropertiesErrors,
  documentPicturePropertyChanges,
  withDocumentPictureAspectRatioLock,
  withDocumentPictureDimension,
} from '../src/internal/features/work/editors/document-picture-properties-dialog-model';
import { DocumentPictureRibbon } from '../src/internal/features/work/editors/document-picture-ribbon';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentImageIdentityFromAttributes,
  normalizeDocumentImageAnchorId,
  normalizeDocumentImageEditId,
  type WorkDocumentImageIdentity,
} from '../src/internal/features/work/work-document-image-identity';
import {
  defaultDocumentImageTransform,
  documentImageAlternativeText,
  documentImageLayerCssZIndex,
  documentImageLayoutOptions,
  documentImageProperties,
  documentImageTransformCss,
  documentImageTransformFromElement,
  normalizeDocumentImageLayer,
  normalizeDocumentImageLayoutOptions,
  normalizeDocumentImageRotation,
} from '../src/internal/features/work/work-document-image-layout';
import {
  documentImageWrapContourCss,
  normalizeDocumentImageWrapContour,
} from '../src/internal/features/work/work-document-image-wrap-contour';
import { measureDocumentLayoutBlocks } from '../src/internal/features/work/work-document-pagination';
import { readDocxImageTransform } from '../src/internal/features/work/work-docx-image-transform';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';

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
  expect(normalizeDocumentImageLayoutOptions({ layout: 'none' }).layout).toBe(
    'none',
  );
  editor.destroy();
});

test('keeps bounded image rotation and reflection in the model and node view', () => {
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="Transform" width="120" height="80"`,
      ' data-office-image-rotation="90"',
      ' data-office-image-flip-horizontal="true">',
    ].join(''),
  );
  selectFirstImage(editor);

  expect(documentImageProperties(editor).transform).toEqual({
    rotation: 90,
    flipHorizontal: true,
    flipVertical: false,
  });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  const image = container?.querySelector<HTMLElement>('img');
  expect(container?.dataset.officeImageRotation).toBe('90');
  expect(container?.dataset.officeImageFlipHorizontal).toBe('true');
  expect(image?.dataset.officeImageRotation).toBe('90');
  expect(
    container?.style.getPropertyValue('--work-document-image-rotation'),
  ).toBe('90deg');
  expect(
    container?.style.getPropertyValue('--work-document-image-flip-x'),
  ).toBe('-1');
  expect(documentImageTransformFromElement(image as HTMLElement)).toEqual({
    rotation: 90,
    flipHorizontal: true,
    flipVertical: false,
  });
  expect(documentImageTransformCss(defaultDocumentImageTransform())).toBe('');
  expect(normalizeDocumentImageRotation(450)).toBe(90);
  expect(normalizeDocumentImageRotation(45)).toBe(0);

  expect(
    editor.commands.setDocumentImageProperties({
      transform: {
        rotation: 180,
        flipHorizontal: false,
        flipVertical: true,
      },
    }),
  ).toBe(true);
  expect(documentImageProperties(editor).transform).toEqual({
    rotation: 180,
    flipHorizontal: false,
    flipVertical: true,
  });
  expect(container?.dataset.officeImageRotation).toBe('180');
  expect(container?.dataset.officeImageFlipHorizontal).toBeUndefined();
  expect(container?.dataset.officeImageFlipVertical).toBe('true');
  expect(editor.commands.undo()).toBe(true);
  expect(documentImageProperties(editor).transform).toEqual({
    rotation: 90,
    flipHorizontal: true,
    flipVertical: false,
  });
  editor.destroy();
});

test('reads only exact quarter-turn picture transforms from DrawingML', () => {
  const document = new DOMParser().parseFromString(
    [
      '<wp:anchor xmlns:wp="urn:wp" xmlns:pic="urn:pic" xmlns:a="urn:a">',
      '<pic:pic><pic:spPr><a:xfrm rot="16200000" flipH="1" flipV="0"/></pic:spPr></pic:pic>',
      '</wp:anchor>',
    ].join(''),
    'application/xml',
  );
  const supported = readDocxImageTransform(document.documentElement);
  expect(supported).toEqual({
    transform: { rotation: 270, flipHorizontal: true, flipVertical: false },
    supported: true,
  });

  const malformed = new DOMParser().parseFromString(
    [
      '<wp:anchor xmlns:wp="urn:wp" xmlns:pic="urn:pic" xmlns:a="urn:a">',
      '<pic:pic><pic:spPr><a:xfrm rot="123" flipH="maybe"/></pic:spPr></pic:pic>',
      '</wp:anchor>',
    ].join(''),
    'application/xml',
  );
  expect(readDocxImageTransform(malformed.documentElement)).toEqual({
    transform: null,
    supported: false,
  });
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

test('keeps floating image drawing layers in the model and presentation', () => {
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="Layered" width="120" height="80"`,
      ' data-office-image-layout="square"',
      ' data-office-image-relative-height="50331648"',
      ' data-office-image-behind-document="true"',
      ' data-office-image-allow-overlap="true"',
      ' data-office-image-layout-in-cell="false"',
      ' data-office-image-lock-anchor="true">',
    ].join(''),
  );
  selectFirstImage(editor);

  const layer = {
    relativeHeight: 50_331_648,
    behindDocument: true,
    allowOverlap: true,
    layoutInCell: false,
    lockAnchor: true,
  };
  expect(documentImageProperties(editor)).toMatchObject({ layer });
  const container = editor.view.dom.querySelector<HTMLElement>(
    '[data-resize-container][data-node="image"]',
  );
  expect(container?.dataset.officeImageRelativeHeight).toBe('50331648');
  expect(container?.dataset.officeImageBehindDocument).toBe('true');
  expect(container?.dataset.officeImageAllowOverlap).toBe('true');
  expect(container?.dataset.officeImageLayoutInCell).toBe('false');
  expect(container?.dataset.officeImageLockAnchor).toBe('true');
  expect(
    container?.style.getPropertyValue('--work-document-image-z-index'),
  ).toBe(String(documentImageLayerCssZIndex(layer)));
  expect(documentImageLayerCssZIndex(layer)).toBeLessThan(0);
  expect(
    documentImageLayerCssZIndex({ ...layer, behindDocument: false }),
  ).toBeGreaterThan(0);
  expect(
    normalizeDocumentImageLayer({
      relativeHeight: 9_000_000_000,
      behindDocument: 'false',
      allowOverlap: '1',
      layoutInCell: '0',
      lockAnchor: true,
    }),
  ).toEqual({
    relativeHeight: 4_294_967_295,
    behindDocument: false,
    allowOverlap: true,
    layoutInCell: false,
    lockAnchor: true,
  });
  editor.destroy();
});

test('gives image copies unique identities and keeps the copied identity through redo', () => {
  expect(normalizeDocumentImageAnchorId('FFFFFFFF')).toBeNull();
  expect(normalizeDocumentImageEditId('80000000')).toBeNull();
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="Source" width="120" height="80"`,
      ' data-office-image-object-id="1A2B3C4D"',
      ' data-office-image-doc-properties-id="42"',
      ' data-office-image-anchor-id="1A2B3C4D"',
      ' data-office-image-edit-id="0A0B0C0D">',
      '<p>After</p>',
    ].join(''),
  );
  const source = imageNodes(editor)[0];
  expect(source?.identity).toEqual({
    objectId: '1A2B3C4D',
    docPropertiesId: 42,
    anchorId: '1A2B3C4D',
    editId: '0A0B0C0D',
  });
  if (!source) throw new Error('Expected a source image.');

  editor.view.dispatch(editor.state.tr.insert(0, source.node));
  const copied = imageNodes(editor);
  expect(copied).toHaveLength(2);
  expect(copied[1]?.identity).toEqual(source.identity);
  expect(copied[0]?.identity.objectId).not.toBe(source.identity.objectId);
  expect(copied[0]?.identity.docPropertiesId).not.toBe(
    source.identity.docPropertiesId,
  );
  expect(copied[0]?.identity.anchorId).not.toBe(source.identity.anchorId);
  expect(copied[0]?.identity.editId).not.toBe(source.identity.editId);
  const copiedIdentity = copied[0]?.identity;

  expect(editor.commands.undo()).toBe(true);
  expect(imageNodes(editor).map((image) => image.identity)).toEqual([
    source.identity,
  ]);
  expect(editor.commands.redo()).toBe(true);
  expect(imageNodes(editor)[0]?.identity).toEqual(copiedIdentity);
  editor.destroy();
});

test('repairs duplicate image identities when legacy content is reopened', () => {
  const duplicateIdentity = [
    ' data-office-image-object-id="12345678"',
    ' data-office-image-doc-properties-id="91"',
    ' data-office-image-anchor-id="12345678"',
    ' data-office-image-edit-id="23456701"',
  ].join('');
  const editor = createImageEditor(
    [
      `<img src="${pixelPng}" alt="First"${duplicateIdentity}>`,
      `<img src="${pixelPng}" alt="Second"${duplicateIdentity}>`,
    ].join(''),
  );
  const images = imageNodes(editor);

  expect(images).toHaveLength(2);
  expect(images[0]?.identity).toEqual({
    objectId: '12345678',
    docPropertiesId: 91,
    anchorId: '12345678',
    editId: '23456701',
  });
  expect(images[1]?.identity.objectId).not.toBe(images[0]?.identity.objectId);
  expect(images[1]?.identity.docPropertiesId).not.toBe(
    images[0]?.identity.docPropertiesId,
  );
  expect(images[1]?.identity.anchorId).not.toBe(images[0]?.identity.anchorId);
  editor.destroy();
});

test('preserves image identity across move, cut, delete, and undo', () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Movable" width="120" height="80"><p>After</p>`,
  );
  const original = imageNodes(editor)[0];
  if (!original) throw new Error('Expected an image node.');

  editor.view.dispatch(
    editor.state.tr.delete(
      original.position,
      original.position + original.node.nodeSize,
    ),
  );
  expect(imageNodes(editor)).toHaveLength(0);
  expect(editor.commands.undo()).toBe(true);
  expect(imageNodes(editor)[0]?.identity).toEqual(original.identity);
  expect(editor.commands.redo()).toBe(true);
  expect(imageNodes(editor)).toHaveLength(0);
  expect(editor.commands.undo()).toBe(true);

  const restored = imageNodes(editor)[0];
  if (!restored) throw new Error('Expected the restored image.');
  const move = editor.state.tr.delete(
    restored.position,
    restored.position + restored.node.nodeSize,
  );
  move.insert(move.doc.content.size, restored.node);
  editor.view.dispatch(move);
  const moved = imageNodes(editor)[0];
  expect(stableImageIdentity(moved?.identity)).toEqual(
    stableImageIdentity(original.identity),
  );
  expect(moved?.identity.editId).not.toBe(original.identity.editId);

  const cut = moved;
  if (!cut) throw new Error('Expected the moved image.');
  editor.view.dispatch(
    editor.state.tr.delete(cut.position, cut.position + cut.node.nodeSize),
  );
  editor.view.dispatch(
    editor.state.tr.insert(editor.state.doc.content.size, cut.node),
  );
  const pasted = imageNodes(editor)[0]?.identity;
  expect(stableImageIdentity(pasted)).toEqual(
    stableImageIdentity(cut.identity),
  );
  expect(pasted?.editId).not.toBe(cut.identity.editId);
  editor.destroy();
});

test('commits picture size, layout, and alternative text as one undoable update', () => {
  const editor = createImageEditor(
    `<img src="${pixelPng}" alt="Original" title="Original title" width="120" height="80">`,
  );
  selectFirstImage(editor);
  const originalHtml = editor.getHTML();
  const originalIdentity = imageNodes(editor)[0]?.identity;
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
      layer: {
        relativeHeight: 50_331_648,
        behindDocument: true,
        allowOverlap: true,
        layoutInCell: false,
        lockAnchor: true,
      },
    }),
  ).toBe(true);

  expect(updateCount).toBe(1);
  const editedIdentity = imageNodes(editor)[0]?.identity;
  expect(stableImageIdentity(editedIdentity)).toEqual(
    stableImageIdentity(originalIdentity),
  );
  expect(editedIdentity?.editId).not.toBe(originalIdentity?.editId);
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
    layer: {
      relativeHeight: 50_331_648,
      behindDocument: true,
      allowOverlap: true,
      layoutInCell: false,
      lockAnchor: true,
    },
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
  expect(container?.dataset.officeImageRelativeHeight).toBe('50331648');
  expect(container?.dataset.officeImageBehindDocument).toBe('true');
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
  expect(imageNodes(editor)[0]?.identity).toEqual(originalIdentity);
  expect(editor.commands.redo()).toBe(true);
  expect(imageNodes(editor)[0]?.identity).toEqual(editedIdentity);
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
    relativeHeight: null,
  });
  expect(
    documentPicturePropertiesErrors({
      ...initial,
      layout: 'none',
      wrapDistance: '30',
    }).wrapDistance,
  ).toBeNull();
});

test('validates and commits floating image layer options', () => {
  const initial = createDocumentPicturePropertiesDraft({
    properties: {
      width: 120,
      height: 80,
      lockAspectRatio: true,
      layout: 'square',
      alignment: 'left',
      wrapDistance: 3,
      alternativeText: 'Layered image',
    },
  });
  const layered = {
    ...initial,
    relativeHeight: '50331648',
    behindDocument: true,
    allowOverlap: true,
    layoutInCell: false,
    lockAnchor: true,
  };
  expect(documentPicturePropertiesErrors(layered).relativeHeight).toBeNull();
  expect(documentPicturePropertyChanges(initial, layered)).toEqual({
    layer: {
      relativeHeight: 50_331_648,
      behindDocument: true,
      allowOverlap: true,
      layoutInCell: false,
      lockAnchor: true,
    },
  });
  expect(
    documentPicturePropertiesErrors({
      ...layered,
      relativeHeight: '4.5',
    }).relativeHeight,
  ).not.toBeNull();
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

test('keeps a free-floating image out of the paginated body flow', () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true">',
      `<img src="${pixelPng}" alt="Free floating" width="120" height="80"`,
      ' data-office-image-layout="none"',
      ' data-office-image-alignment="right"',
      ' data-office-image-horizontal-offset="12.5"',
      ' data-office-image-vertical-offset="7.25">',
      '<p>Text continues underneath the image.</p>',
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

  expect(container.dataset.officeImageLayout).toBe('none');
  expect(imageBlock?.block.height).toBe(1);
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
  fireEvent.click(screen.getByRole('button', { name: '向右旋转' }));
  fireEvent.click(screen.getByRole('button', { name: '向右旋转' }));
  fireEvent.click(screen.getByRole('button', { name: '水平翻转' }));
  fireEvent.click(screen.getByRole('button', { name: '垂直翻转' }));
  expect(documentImageProperties(editor).transform).toEqual({
    rotation: 180,
    flipHorizontal: true,
    flipVertical: true,
  });
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(screen.getByRole('button', { name: '水平翻转' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  fireEvent.click(screen.getByRole('combobox', { name: '图片与文字距离' }));
  fireEvent.click(screen.getByRole('option', { name: '10 毫米' }));
  expect(documentImageLayoutOptions(editor).wrapDistance).toBe(10);

  fireEvent.click(screen.getByRole('button', { name: '紧密环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('tight');
  fireEvent.click(screen.getByRole('button', { name: '穿越环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('through');
  fireEvent.click(screen.getByRole('button', { name: '上下环绕' }));
  expect(documentImageLayoutOptions(editor).layout).toBe('topBottom');
  fireEvent.click(screen.getByRole('button', { name: '自由浮动' }));
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(documentImageLayoutOptions(editor).layout).toBe('none');
  expect(wrapDistance).toBeDisabled();
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
  const rotation = screen.getByRole('combobox', { name: '图片旋转角度' });
  expect(rotation).toHaveTextContent('180°');
  fireEvent.click(rotation);
  fireEvent.click(screen.getByRole('option', { name: '90°' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '垂直翻转图片' }));

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
  fireEvent.click(screen.getByRole('radio', { name: '自由浮动' }));
  expect(
    screen.getByRole('combobox', { name: '图片文字环绕侧' }),
  ).toBeDisabled();
  expect(
    screen.getByRole('textbox', { name: '图片与文字距离（毫米）' }),
  ).toBeDisabled();
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
  fireEvent.change(screen.getByRole('textbox', { name: '图片绘图层级顺序' }), {
    target: { value: '50331648' },
  });
  fireEvent.click(screen.getByRole('checkbox', { name: '图片衬于文字下方' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '允许图片重叠' }));
  fireEvent.click(
    screen.getByRole('checkbox', { name: '图片随表格单元格布局' }),
  );
  fireEvent.click(screen.getByRole('checkbox', { name: '锁定图片锚点' }));
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
    transform: {
      rotation: 90,
      flipHorizontal: true,
      flipVertical: false,
    },
    layer: {
      relativeHeight: 50_331_648,
      behindDocument: true,
      allowOverlap: true,
      layoutInCell: false,
      lockAnchor: true,
    },
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
  expect(screen.getByRole('button', { name: '自由浮动' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '图片属性' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '删除图片' })).toBeDisabled();

  selectFirstImage(editor);
  view.rerender(<DocumentPictureRibbon editor={editor} />);
  expect(screen.getByRole('button', { name: '四周环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '紧密环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '穿越环绕' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '自由浮动' })).toBeEnabled();
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
    ' data-office-image-object-id="1A2B3C4D"',
    ' data-office-image-doc-properties-id="42"',
    ' data-office-image-anchor-id="1A2B3C4D"',
    ' data-office-image-edit-id="0A0B0C0D"',
    ' data-office-image-alignment="right"',
    ' data-office-image-wrap-distance="5"',
    ' data-office-image-wrap-side="right"',
    ' data-office-image-rotation="90"',
    ' data-office-image-flip-horizontal="true"',
    ' data-office-image-flip-vertical="true"',
    ' data-office-image-relative-height="50331648"',
    ' data-office-image-behind-document="false"',
    ' data-office-image-allow-overlap="true"',
    ' data-office-image-layout-in-cell="false"',
    ' data-office-image-lock-anchor="true">',
    `<img src="${pixelPng}" alt="Centered diagram" width="140" height="90"`,
    ' data-office-image-layout="topBottom"',
    ' data-office-image-alignment="center"',
    ' data-office-image-wrap-distance="2"',
    ' data-office-image-relative-height="0"',
    ' data-office-image-behind-document="true">',
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
    `<img src="${pixelPng}" alt="Free floating" width="115" height="68"`,
    ' data-office-image-layout="none"',
    ' data-office-image-alignment="left"',
    ' data-office-image-horizontal-offset="15.5"',
    ' data-office-image-vertical-offset="-4"',
    ' data-office-image-horizontal-reference="margin"',
    ' data-office-image-vertical-reference="page"',
    ' data-office-image-relative-height="1024"',
    ' data-office-image-behind-document="true"',
    ' data-office-image-allow-overlap="true">',
    `<img src="${pixelPng}" alt="Inline crop" width="90" height="60"`,
    ' data-office-image-object-id="0BADF00D"',
    ' data-office-image-doc-properties-id="77"',
    ' data-office-image-anchor-id="0BADF00D"',
    ' data-office-image-edit-id="10203040"',
    ' data-office-image-crop-bottom="10"',
    ' data-office-image-crop-left="20">',
  ].join('');

  const blob = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml = await archive.file('word/document.xml')?.async('string');

  expect(documentXml).toBeDefined();
  const drawingPropertyIds = Array.from(
    documentXml?.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g) ?? [],
    (match) => match[1],
  );
  const anchorIds = Array.from(
    documentXml?.matchAll(/\bwp14:anchorId="([0-9A-F]{8})"/g) ?? [],
    (match) => match[1],
  );
  const editIds = Array.from(
    documentXml?.matchAll(/\bwp14:editId="([0-9A-F]{8})"/g) ?? [],
    (match) => match[1],
  );
  expect(new Set(drawingPropertyIds).size).toBe(drawingPropertyIds.length);
  expect(new Set(anchorIds).size).toBe(anchorIds.length);
  expect(editIds).toHaveLength(anchorIds.length);
  expect(documentXml).toMatch(/\bmc:Ignorable="[^"]*\bwp14\b[^"]*"/);
  expect(
    [...anchorIds, ...editIds].every((id) => {
      const value = Number.parseInt(id, 16);
      return value > 0 && value < 0x8000_0000;
    }),
  ).toBe(true);
  expect(documentXml).toMatch(
    /<wp:anchor\b(?=[^>]*wp14:anchorId="1A2B3C4D")(?=[^>]*wp14:editId="0A0B0C0D")[\s\S]*?<wp:docPr\b(?=[^>]*id="42")(?=[^>]*name="Right chart")/,
  );
  expect(documentXml).toMatch(
    /<wp:inline\b(?=[^>]*wp14:anchorId="0BADF00D")(?=[^>]*wp14:editId="10203040")[\s\S]*?<wp:docPr\b(?=[^>]*id="77")(?=[^>]*name="Inline crop")/,
  );
  expect(documentXml?.match(/<wp:anchor\b/g)).toHaveLength(6);
  expect(documentXml).toMatch(
    /<wp:anchor\b(?=[^>]*relativeHeight="50331648")(?=[^>]*behindDoc="0")(?=[^>]*allowOverlap="1")(?=[^>]*layoutInCell="0")(?=[^>]*locked="1")/,
  );
  expect(documentXml).toMatch(
    /<wp:anchor\b(?=[^>]*relativeHeight="0")(?=[^>]*behindDoc="1")/,
  );
  expect(documentXml).toMatch(/<wp:wrapSquare\b(?=[^>]*wrapText="right")/);
  expect(documentXml).toContain('<wp:wrapTopAndBottom');
  expect(documentXml).toContain('<wp:wrapNone');
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
    /<wp:positionH relativeFrom="margin"><wp:posOffset>558000<\/wp:posOffset>/,
  );
  expect(documentXml).toMatch(
    /<wp:positionV relativeFrom="page"><wp:posOffset>-144000<\/wp:posOffset>/,
  );
  expect(documentXml).toMatch(
    /<a:srcRect\b(?=[^>]*t="12500")(?=[^>]*r="5000")(?=[^>]*b="2250")(?=[^>]*l="10000")/,
  );
  expect(documentXml).toMatch(
    /<a:xfrm\b(?=[^>]*rot="5400000")(?=[^>]*flipH="1")(?=[^>]*flipV="1")/,
  );
  expect(documentXml).toMatch(
    /<a:srcRect\b(?=[^>]*b="10000")(?=[^>]*l="20000")/,
  );
  expect(documentXml).not.toContain('__A3S_IMAGE_CROP_');
  expect(documentXml).not.toContain('__A3S_IMAGE_WRAP_');
  expect(documentXml).not.toContain('__A3S_IMAGE_LAYER_');
  expect(documentXml).not.toContain('__A3S_IMAGE_IDENTITY_');

  const relationshipsXml = await archive
    .file('word/_rels/document.xml.rels')
    ?.async('string');
  const embeddedRelationshipIds = Array.from(
    documentXml?.matchAll(/\br:embed="([^"]+)"/g) ?? [],
    (match) => match[1],
  );
  const imageRelationshipIds = new Set(
    Array.from(
      relationshipsXml?.matchAll(
        /<Relationship\b(?=[^>]*Id="([^"]+)")(?=[^>]*Type="[^"]*\/image")/g,
      ) ?? [],
      (match) => match[1],
    ),
  );
  expect(embeddedRelationshipIds).not.toHaveLength(0);
  expect(
    embeddedRelationshipIds.every((id) => imageRelationshipIds.has(id)),
  ).toBe(true);

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
  expect(inlineCrop?.dataset.officeImageObjectId).toBe('0BADF00D');
  expect(inlineCrop?.dataset.officeImageDocPropertiesId).toBe('77');
  expect(inlineCrop?.dataset.officeImageAnchorId).toBe('0BADF00D');
  expect(inlineCrop?.dataset.officeImageEditId).toBe('10203040');
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
  const freeFloating = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Free floating"]',
  );
  expect(freeFloating?.dataset.officeImageLayout).toBe('none');
  expect(freeFloating?.dataset.officeImageHorizontalOffset).toBe('15.5');
  expect(freeFloating?.dataset.officeImageVerticalOffset).toBe('-4');
  expect(freeFloating?.dataset.officeImageHorizontalReference).toBe('margin');
  expect(freeFloating?.dataset.officeImageVerticalReference).toBe('page');
  expect(freeFloating?.dataset.officeImageRelativeHeight).toBe('1024');
  expect(freeFloating?.dataset.officeImageBehindDocument).toBe('true');
  expect(freeFloating?.dataset.officeImageAllowOverlap).toBe('true');
  const alignedImage = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Right chart"]',
  );
  expect(alignedImage?.dataset.officeImageHorizontalOffset).toBeUndefined();
  expect(alignedImage?.dataset.officeImageVerticalOffset).toBe('0');
  expect(alignedImage?.dataset.officeImageWrapSide).toBe('right');
  expect(alignedImage?.dataset.officeImageRelativeHeight).toBe('50331648');
  expect(alignedImage?.dataset.officeImageBehindDocument).toBe('false');
  expect(alignedImage?.dataset.officeImageAllowOverlap).toBe('true');
  expect(alignedImage?.dataset.officeImageLayoutInCell).toBe('false');
  expect(alignedImage?.dataset.officeImageLockAnchor).toBe('true');
  expect(alignedImage?.dataset.officeImageObjectId).toBe('1A2B3C4D');
  expect(alignedImage?.dataset.officeImageDocPropertiesId).toBe('42');
  expect(alignedImage?.dataset.officeImageAnchorId).toBe('1A2B3C4D');
  expect(alignedImage?.dataset.officeImageEditId).toBe('0A0B0C0D');
  expect(alignedImage?.dataset.officeImageRotation).toBe('90');
  expect(alignedImage?.dataset.officeImageFlipHorizontal).toBe('true');
  expect(alignedImage?.dataset.officeImageFlipVertical).toBe('true');
  expect(
    alignedImage?.style.getPropertyValue('--work-document-image-z-index'),
  ).toBe('11720');
  const behindImage = importedHtml.querySelector<HTMLImageElement>(
    'img[alt="Centered diagram"]',
  );
  expect(behindImage?.dataset.officeImageRelativeHeight).toBe('0');
  expect(behindImage?.dataset.officeImageBehindDocument).toBe('true');
  expect(
    behindImage?.style.getPropertyValue('--work-document-image-z-index'),
  ).toBe('-1000001');

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
  expect(regeneratedXml).toMatch(
    /<a:xfrm\b(?=[^>]*rot="5400000")(?=[^>]*flipH="1")(?=[^>]*flipV="1")/,
  );
  expect(regeneratedXml).toMatch(/<wp:wrapTight\b(?=[^>]*wrapText="largest")/);
  expect(regeneratedXml).toMatch(/<wp:wrapThrough\b(?=[^>]*wrapText="left")/);
  expect(regeneratedXml).toContain('<wp:wrapNone');
  expect(regeneratedXml).toContain('<wp:lineTo x="10800" y="16200"/>');
  expect(regeneratedXml).toMatch(
    /<wp:anchor\b(?=[^>]*relativeHeight="50331648")(?=[^>]*behindDoc="0")(?=[^>]*allowOverlap="1")(?=[^>]*layoutInCell="0")(?=[^>]*locked="1")/,
  );
  expect(regeneratedXml).toMatch(
    /<wp:anchor\b(?=[^>]*relativeHeight="0")(?=[^>]*behindDoc="1")/,
  );
  expect(regeneratedXml).toMatch(
    /<wp:anchor\b(?=[^>]*wp14:anchorId="1A2B3C4D")(?=[^>]*wp14:editId="0A0B0C0D")[\s\S]*?<wp:docPr\b(?=[^>]*id="42")(?=[^>]*name="Right chart")/,
  );
  expect(regeneratedXml).toMatch(
    /<wp:inline\b(?=[^>]*wp14:anchorId="0BADF00D")(?=[^>]*wp14:editId="10203040")[\s\S]*?<wp:docPr\b(?=[^>]*id="77")(?=[^>]*name="Inline crop")/,
  );
});

test('reports arbitrary native picture transforms as a visible compatibility boundary', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document')
    throw new Error('Expected a document artifact.');
  artifact.content.html = `<img src="${pixelPng}" alt="Rotated" width="120" height="80">`;
  const blob = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentEntry = archive.file('word/document.xml');
  if (!documentEntry) throw new Error('Expected a document XML part.');
  const documentXml = await documentEntry.async('string');
  archive.file(
    'word/document.xml',
    documentXml.replace('<a:xfrm>', '<a:xfrm rot="123">'),
  );
  const mutated = await archive.generateAsync({ type: 'arraybuffer' });
  const report = await analyzeDocxCompatibility(
    new File([mutated], 'arbitrary-transform.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    [],
  );
  expect(report.issues).toContainEqual(
    expect.objectContaining({ code: 'docx.images.transform' }),
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

function imageNodes(editor: Editor): Array<{
  node: ProseMirrorNode;
  position: number;
  identity: WorkDocumentImageIdentity;
}> {
  const images: Array<{
    node: ProseMirrorNode;
    position: number;
    identity: WorkDocumentImageIdentity;
  }> = [];
  editor.state.doc.descendants((node, position) => {
    if (node.type.name !== 'image') return;
    const identity = documentImageIdentityFromAttributes(node.attrs);
    if (!identity) throw new Error('Expected a normalized image identity.');
    images.push({ node, position, identity });
  });
  return images;
}

function stableImageIdentity(identity: WorkDocumentImageIdentity | undefined) {
  return identity
    ? {
        objectId: identity.objectId,
        docPropertiesId: identity.docPropertiesId,
        anchorId: identity.anchorId,
      }
    : undefined;
}
