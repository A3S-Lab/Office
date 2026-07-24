import type { IFloating, ParagraphChild } from 'docx';
import { documentImageLayoutFromElement } from './work-document-image-layout';

export async function imageToDocx(
  element: HTMLImageElement,
  docx: typeof import('docx'),
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
    return new docx.ImageRun({
      type,
      data: await blob.arrayBuffer(),
      transformation: {
        width: Math.max(24, Math.round(dimensions.width * scale)),
        height: Math.max(24, Math.round(dimensions.height * scale)),
      },
      altText: { name: alt, description: alt, title: alt },
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
  if (image.layout === 'inline') return undefined;
  const distance = Math.round(image.wrapDistance * 36_000);
  const align =
    image.alignment === 'left'
      ? docx.HorizontalPositionAlign.LEFT
      : image.alignment === 'right'
        ? docx.HorizontalPositionAlign.RIGHT
        : docx.HorizontalPositionAlign.CENTER;
  return {
    horizontalPosition: {
      relative: docx.HorizontalPositionRelativeFrom.COLUMN,
      align,
    },
    verticalPosition: {
      relative: docx.VerticalPositionRelativeFrom.PARAGRAPH,
      offset: 0,
    },
    allowOverlap: false,
    behindDocument: false,
    layoutInCell: true,
    lockAnchor: false,
    margins: {
      top: distance,
      right: distance,
      bottom: distance,
      left: distance,
    },
    wrap: {
      type:
        image.layout === 'square'
          ? docx.TextWrappingType.SQUARE
          : docx.TextWrappingType.TOP_AND_BOTTOM,
      ...(image.layout === 'square'
        ? { side: docx.TextWrappingSide.BOTH_SIDES }
        : {}),
    },
  };
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
