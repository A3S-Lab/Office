import type { IFloating, ParagraphChild } from 'docx';
import {
  documentImageCropFromElement,
  documentImageLayerFromElement,
  documentImageLayoutFromElement,
  documentImagePositionFromElement,
  documentImageTransformFromElement,
  wrapsBesideImage,
} from './work-document-image-layout';
import { documentImageWrapContourFromElement } from './work-document-image-wrap-contour';
import type { DocxImageCropPatchCollector } from './work-docx-image-crop';
import type { DocxImageIdentityPatchCollector } from './work-docx-image-identity';
import type { DocxImageLayerPatchCollector } from './work-docx-image-layer';
import type { DocxImageTransformPatchCollector } from './work-docx-image-transform';
import type { DocxImageWrapPatchCollector } from './work-docx-image-wrap';

export async function imageToDocx(
  element: HTMLImageElement,
  docx: typeof import('docx'),
  cropPatches?: DocxImageCropPatchCollector,
  wrapPatches?: DocxImageWrapPatchCollector,
  layerPatches?: DocxImageLayerPatchCollector,
  identityPatches?: DocxImageIdentityPatchCollector,
  transformPatches?: DocxImageTransformPatchCollector,
): Promise<ParagraphChild> {
  const source = element.getAttribute('src');
  const alt =
    element.getAttribute('alt') || element.getAttribute('title') || 'Image';
  if (!source) return new docx.TextRun(`[${alt}]`);
  try {
    const blob = await fetch(source).then((response) => {
      if (!response.ok)
        throw new Error(`Image request failed with HTTP ${response.status}`);
      return response.blob();
    });
    const type = docxImageType(blob.type, source);
    if (!type) return new docx.TextRun(`[${alt}]`);
    const dimensions =
      element.width > 0 && element.height > 0
        ? { width: element.width, height: element.height }
        : await imageDimensions(blob);
    const maximumWidth = 520;
    const scale = Math.min(1, maximumWidth / Math.max(1, dimensions.width));
    const cropMarker = cropPatches?.marker(
      documentImageCropFromElement(element),
    );
    const imageLayout = documentImageLayoutFromElement(element);
    const identity = identityPatches?.register(element);
    const wrapMarker = wrapPatches?.marker(
      imageLayout,
      documentImageWrapContourFromElement(element),
    );
    const layer = documentImageLayerFromElement(element);
    const layerMarker = layerPatches?.marker(imageLayout.layout, layer);
    const transformMarker = transformPatches?.marker(
      documentImageTransformFromElement(element),
    );
    return new docx.ImageRun({
      type,
      data: await blob.arrayBuffer(),
      transformation: {
        width: Math.max(24, Math.round(dimensions.width * scale)),
        height: Math.max(24, Math.round(dimensions.height * scale)),
      },
      altText: {
        name: `${alt}${cropMarker ?? ''}${wrapMarker ?? ''}${layerMarker ?? ''}${identity?.marker ?? ''}${transformMarker ?? ''}`,
        description: alt,
        title: alt,
        id: identity ? String(identity.docPropertiesId) : undefined,
      },
      floating: documentImageFloatingOptions(element, docx),
    });
  } catch {
    return new docx.TextRun(`[${alt}]`);
  }
}

function documentImageFloatingOptions(
  element: HTMLImageElement,
  docx: typeof import('docx'),
): IFloating | undefined {
  const image = documentImageLayoutFromElement(element);
  const position = documentImagePositionFromElement(element);
  if (image.layout === 'inline') return undefined;
  const layer = documentImageLayerFromElement(element);
  const distance = Math.round(image.wrapDistance * 36_000);
  const align =
    image.alignment === 'left'
      ? docx.HorizontalPositionAlign.LEFT
      : image.alignment === 'right'
        ? docx.HorizontalPositionAlign.RIGHT
        : docx.HorizontalPositionAlign.CENTER;
  return {
    horizontalPosition: {
      relative: horizontalPositionReference(
        position?.horizontalReference ?? 'column',
        docx,
      ),
      ...(position?.horizontalOffset !== null &&
      position?.horizontalOffset !== undefined
        ? { offset: millimetersToEmus(position.horizontalOffset) }
        : { align }),
    },
    verticalPosition: {
      relative: verticalPositionReference(
        position?.verticalReference ?? 'paragraph',
        docx,
      ),
      offset: millimetersToEmus(position?.verticalOffset ?? 0),
    },
    allowOverlap: layer.allowOverlap,
    behindDocument: layer.behindDocument,
    layoutInCell: layer.layoutInCell,
    lockAnchor: layer.lockAnchor,
    zIndex: layer.relativeHeight,
    margins: {
      top: distance,
      right: distance,
      bottom: distance,
      left: distance,
    },
    wrap: {
      type:
        image.layout === 'none'
          ? docx.TextWrappingType.NONE
          : wrapsBesideImage(image.layout)
            ? image.layout === 'square'
              ? docx.TextWrappingType.SQUARE
              : docx.TextWrappingType.TIGHT
            : docx.TextWrappingType.TOP_AND_BOTTOM,
      ...(wrapsBesideImage(image.layout)
        ? { side: textWrappingSide(image.wrapSide, docx) }
        : {}),
    },
  };
}

function textWrappingSide(
  value: 'bothSides' | 'largest' | 'left' | 'right',
  docx: typeof import('docx'),
) {
  if (value === 'left') return docx.TextWrappingSide.LEFT;
  if (value === 'right') return docx.TextWrappingSide.RIGHT;
  if (value === 'largest') return docx.TextWrappingSide.LARGEST;
  return docx.TextWrappingSide.BOTH_SIDES;
}

function horizontalPositionReference(
  value: 'column' | 'margin' | 'page',
  docx: typeof import('docx'),
) {
  if (value === 'margin') return docx.HorizontalPositionRelativeFrom.MARGIN;
  if (value === 'page') return docx.HorizontalPositionRelativeFrom.PAGE;
  return docx.HorizontalPositionRelativeFrom.COLUMN;
}

function verticalPositionReference(
  value: 'margin' | 'page' | 'paragraph',
  docx: typeof import('docx'),
) {
  if (value === 'margin') return docx.VerticalPositionRelativeFrom.MARGIN;
  if (value === 'page') return docx.VerticalPositionRelativeFrom.PAGE;
  return docx.VerticalPositionRelativeFrom.PARAGRAPH;
}

function millimetersToEmus(value: number): number {
  return Math.round(value * 36_000);
}

function docxImageType(
  contentType: string,
  source: string,
): 'jpg' | 'png' | 'gif' | 'bmp' | null {
  const value = `${contentType} ${source}`.toLowerCase();
  if (value.includes('png')) return 'png';
  if (value.includes('jpeg') || value.includes('jpg')) return 'jpg';
  if (value.includes('gif')) return 'gif';
  if (value.includes('bmp')) return 'bmp';
  return null;
}

function imageDimensions(
  blob: Blob,
): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const image = new window.Image();
    image.addEventListener(
      'load',
      () => {
        URL.revokeObjectURL(url);
        resolve({
          width: image.naturalWidth || 640,
          height: image.naturalHeight || 360,
        });
      },
      { once: true },
    );
    image.addEventListener(
      'error',
      () => {
        URL.revokeObjectURL(url);
        reject(new Error('Image dimensions could not be read'));
      },
      { once: true },
    );
    image.src = url;
  });
}
