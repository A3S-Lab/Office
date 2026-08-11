import { type Editor, mergeAttributes, ResizableNodeView } from '@tiptap/core';
import Image from '@tiptap/extension-image';
import { closeHistory } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { NodeSelection } from '@tiptap/pm/state';
import {
  applyDocumentImageWrapContourToElement,
  defaultDocumentImageWrapContour,
  documentImageWrapContourCss,
  documentImageWrapContourFromElement,
  normalizeDocumentImageWrapContour,
  normalizeDocumentImageWrapSide,
  serializeDocumentImageWrapPoints,
  type WorkDocumentImageWrapContour,
  type WorkDocumentImageWrapSide,
} from './work-document-image-wrap-contour';

export type WorkDocumentImageLayout =
  | 'inline'
  | 'square'
  | 'tight'
  | 'through'
  | 'topBottom';
export type WorkDocumentImageAlignment = 'left' | 'center' | 'right';
export type WorkDocumentImageHorizontalReference = 'column' | 'margin' | 'page';
export type WorkDocumentImageVerticalReference =
  | 'paragraph'
  | 'margin'
  | 'page';

export interface WorkDocumentImagePosition {
  horizontalOffset: number | null;
  verticalOffset: number | null;
  horizontalReference: WorkDocumentImageHorizontalReference;
  verticalReference: WorkDocumentImageVerticalReference;
}

export interface WorkDocumentImageCrop {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface WorkDocumentImageLayer {
  relativeHeight: number;
  behindDocument: boolean;
  allowOverlap: boolean;
  layoutInCell: boolean;
  lockAnchor: boolean;
}

export interface WorkDocumentImageLayoutOptions {
  layout: WorkDocumentImageLayout;
  alignment: WorkDocumentImageAlignment;
  wrapDistance: number;
  wrapSide: WorkDocumentImageWrapSide;
}

export interface WorkDocumentImageProperties
  extends WorkDocumentImageLayoutOptions {
  width: number | null;
  height: number | null;
  lockAspectRatio: boolean;
  alternativeText: string;
  position?: WorkDocumentImagePosition | null;
  crop?: WorkDocumentImageCrop | null;
  contour?: WorkDocumentImageWrapContour | null;
  layer?: WorkDocumentImageLayer | null;
}

export type {
  WorkDocumentImageWrapContour,
  WorkDocumentImageWrapPoint,
  WorkDocumentImageWrapSide,
} from './work-document-image-wrap-contour';

export interface DocumentImageCommandOptions {
  restoreFocus?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentImage: {
      setDocumentImageProperties: (
        value: Partial<WorkDocumentImageProperties>,
        options?: DocumentImageCommandOptions,
      ) => ReturnType;
      setDocumentImageAlternativeText: (
        alternativeText: string,
        options?: DocumentImageCommandOptions,
      ) => ReturnType;
      setDocumentImageLayoutOptions: (
        value: Partial<WorkDocumentImageLayoutOptions>,
        options?: DocumentImageCommandOptions,
      ) => ReturnType;
    };
  }
}

const DEFAULT_IMAGE_LAYOUT: WorkDocumentImageLayout = 'inline';
const DEFAULT_IMAGE_ALIGNMENT: WorkDocumentImageAlignment = 'center';
const DEFAULT_WRAP_DISTANCE_MILLIMETERS = 3;
const MAX_WRAP_DISTANCE_MILLIMETERS = 25;
const DEFAULT_LOCK_ASPECT_RATIO = true;
const DEFAULT_HORIZONTAL_REFERENCE: WorkDocumentImageHorizontalReference =
  'column';
const DEFAULT_VERTICAL_REFERENCE: WorkDocumentImageVerticalReference =
  'paragraph';
const MAX_IMAGE_OFFSET_MILLIMETERS = 558.7;
export const MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT = 0xffff_ffff;
const DEFAULT_IMAGE_LAYER: WorkDocumentImageLayer = {
  relativeHeight: 0,
  behindDocument: false,
  allowOverlap: false,
  layoutInCell: true,
  lockAnchor: false,
};
const DOCUMENT_IMAGE_CSS_LAYER_RANGE = 1_000_000;

export const DocumentImage = Image.extend({
  addCommands() {
    return {
      ...(this.parent?.() ?? {}),
      setDocumentImageProperties:
        (value, options = {}) =>
        ({ chain, state, tr }) => {
          if (!isDocumentImageSelection(state.selection)) return false;
          const attributes = documentImageAttributesForChanges(value);
          if (!Object.keys(attributes).length) return false;
          closeHistory(tr);
          let commandChain = chain();
          if (options.restoreFocus !== false) {
            commandChain = commandChain.focus();
          }
          return commandChain.updateAttributes('image', attributes).run();
        },
      setDocumentImageAlternativeText:
        (alternativeText, options = {}) =>
        ({ commands }) =>
          commands.setDocumentImageProperties({ alternativeText }, options),
      setDocumentImageLayoutOptions:
        (value, options = {}) =>
        ({ commands }) =>
          commands.setDocumentImageProperties(value, options),
    };
  },

  addAttributes() {
    return {
      ...this.parent?.(),
      layout: {
        default: DEFAULT_IMAGE_LAYOUT,
        parseHTML: (element) =>
          normalizeDocumentImageLayout(
            element.getAttribute('data-office-image-layout'),
          ),
        renderHTML: (attributes) => ({
          'data-office-image-layout': normalizeDocumentImageLayout(
            attributes.layout,
          ),
        }),
      },
      alignment: {
        default: DEFAULT_IMAGE_ALIGNMENT,
        parseHTML: (element) =>
          normalizeDocumentImageAlignment(
            element.getAttribute('data-office-image-alignment'),
          ),
        renderHTML: (attributes) => ({
          'data-office-image-alignment': normalizeDocumentImageAlignment(
            attributes.alignment,
          ),
        }),
      },
      wrapDistance: {
        default: DEFAULT_WRAP_DISTANCE_MILLIMETERS,
        parseHTML: (element) =>
          normalizeDocumentImageWrapDistance(
            element.getAttribute('data-office-image-wrap-distance'),
          ),
        renderHTML: (attributes) => {
          const distance = normalizeDocumentImageWrapDistance(
            attributes.wrapDistance,
          );
          return {
            'data-office-image-wrap-distance':
              formatImageLayoutNumber(distance),
            style: `--work-document-image-wrap-distance:${formatImageLayoutNumber(distance)}mm`,
          };
        },
      },
      wrapSide: {
        default: 'bothSides',
        parseHTML: (element) =>
          normalizeDocumentImageWrapSide(
            element.getAttribute('data-office-image-wrap-side'),
          ),
        renderHTML: (attributes) => {
          const side = normalizeDocumentImageWrapSide(attributes.wrapSide);
          const layout = normalizeDocumentImageLayout(attributes.layout);
          return side !== 'bothSides' || wrapsBesideImage(layout)
            ? { 'data-office-image-wrap-side': side }
            : {};
        },
      },
      lockAspectRatio: {
        default: DEFAULT_LOCK_ASPECT_RATIO,
        parseHTML: (element) =>
          normalizeDocumentImageLockAspectRatio(
            element.getAttribute('data-office-image-lock-aspect-ratio'),
          ),
        renderHTML: (attributes) => ({
          'data-office-image-lock-aspect-ratio': String(
            normalizeDocumentImageLockAspectRatio(attributes.lockAspectRatio),
          ),
        }),
      },
      horizontalOffset: imagePositionNumberAttribute(
        'data-office-image-horizontal-offset',
      ),
      verticalOffset: imagePositionNumberAttribute(
        'data-office-image-vertical-offset',
      ),
      horizontalReference: {
        default: DEFAULT_HORIZONTAL_REFERENCE,
        parseHTML: (element) =>
          normalizeDocumentImageHorizontalReference(
            element.getAttribute('data-office-image-horizontal-reference'),
          ),
        renderHTML: (attributes) => ({
          'data-office-image-horizontal-reference':
            normalizeDocumentImageHorizontalReference(
              attributes.horizontalReference,
            ),
        }),
      },
      verticalReference: {
        default: DEFAULT_VERTICAL_REFERENCE,
        parseHTML: (element) =>
          normalizeDocumentImageVerticalReference(
            element.getAttribute('data-office-image-vertical-reference'),
          ),
        renderHTML: (attributes) => ({
          'data-office-image-vertical-reference':
            normalizeDocumentImageVerticalReference(
              attributes.verticalReference,
            ),
        }),
      },
      cropTop: imageCropNumberAttribute('top'),
      cropRight: imageCropNumberAttribute('right'),
      cropBottom: imageCropNumberAttribute('bottom'),
      cropLeft: imageCropNumberAttribute('left'),
      wrapPolygon: imageWrapPolygonAttribute(),
      wrapPolygonEdited: {
        default: false,
        parseHTML: (element) =>
          documentImageWrapContourFromElement(element)?.edited ?? false,
        renderHTML: () => ({}),
      },
      relativeHeight: imageRelativeHeightAttribute(),
      behindDocument: imageLayerBooleanAttribute(
        'behindDocument',
        'data-office-image-behind-document',
        DEFAULT_IMAGE_LAYER.behindDocument,
      ),
      allowOverlap: imageLayerBooleanAttribute(
        'allowOverlap',
        'data-office-image-allow-overlap',
        DEFAULT_IMAGE_LAYER.allowOverlap,
      ),
      layoutInCell: imageLayerBooleanAttribute(
        'layoutInCell',
        'data-office-image-layout-in-cell',
        DEFAULT_IMAGE_LAYER.layoutInCell,
      ),
      lockAnchor: imageLayerBooleanAttribute(
        'lockAnchor',
        'data-office-image-lock-anchor',
        DEFAULT_IMAGE_LAYER.lockAnchor,
      ),
    };
  },

  addNodeView() {
    const resize = this.options.resize;
    if (
      resize === false ||
      !resize.enabled ||
      typeof document === 'undefined'
    ) {
      return null;
    }
    const { directions, minWidth, minHeight, alwaysPreserveAspectRatio } =
      resize;
    return ({ node, getPos, HTMLAttributes, editor }) => {
      const element = document.createElement('img');
      element.draggable = false;
      applyInitialImageHtmlAttributes(element, HTMLAttributes);
      element.style.visibility = 'hidden';
      element.style.pointerEvents = 'none';
      const reveal = () => {
        element.style.visibility = '';
        element.style.pointerEvents = '';
      };
      element.addEventListener('load', reveal);
      element.addEventListener('error', reveal);

      let container: HTMLElement | null = null;
      const sync = (updatedNode: ProseMirrorNode) => {
        syncDocumentImageNodeView(element, container, updatedNode.attrs);
      };
      const nodeView = new ResizableNodeView({
        element,
        editor,
        node,
        getPos,
        onResize: (width, height) => {
          element.style.width = `${width}px`;
          element.style.height = `${height}px`;
        },
        onCommit: (width, height) => {
          const position = getPos();
          if (position === undefined) return;
          editor
            .chain()
            .setNodeSelection(position)
            .updateAttributes(this.name, { width, height })
            .run();
        },
        onUpdate: (updatedNode) => {
          if (updatedNode.type !== node.type) return false;
          if (!sameDocumentImageResizeConfiguration(node, updatedNode)) {
            return false;
          }
          sync(updatedNode);
          return true;
        },
        options: {
          directions,
          min: { width: minWidth, height: minHeight },
          preserveAspectRatio:
            node.attrs.lockAspectRatio === null ||
            node.attrs.lockAspectRatio === undefined
              ? alwaysPreserveAspectRatio === true
              : normalizeDocumentImageLockAspectRatio(
                  node.attrs.lockAspectRatio,
                ),
        },
      });
      container = nodeView.dom as HTMLElement;
      sync(node);
      if (element.complete) reveal();
      return nodeView;
    };
  },
});

export function documentImageLayoutOptions(
  editor: Editor,
): WorkDocumentImageLayoutOptions {
  const attributes = editor.getAttributes('image') as Record<string, unknown>;
  return normalizeDocumentImageLayoutOptions(attributes);
}

export function documentImageProperties(
  editor: Editor,
): WorkDocumentImageProperties {
  const attributes = editor.getAttributes('image') as Record<string, unknown>;
  const layout = normalizeDocumentImageLayoutOptions(attributes);
  const position = normalizeDocumentImagePosition(attributes);
  const crop = normalizeDocumentImageCrop(attributes);
  const contour = effectiveDocumentImageWrapContour(attributes);
  return {
    ...layout,
    width: normalizeDocumentImageDimension(attributes.width),
    height: normalizeDocumentImageDimension(attributes.height),
    lockAspectRatio: normalizeDocumentImageLockAspectRatio(
      attributes.lockAspectRatio,
    ),
    alternativeText: documentImageAlternativeText(editor),
    ...(position ? { position } : {}),
    ...(crop ? { crop } : {}),
    ...(contour ? { contour } : {}),
    ...(layout.layout === 'inline'
      ? {}
      : { layer: normalizeDocumentImageLayer(attributes) }),
  };
}

export function setDocumentImageProperties(
  editor: Editor,
  value: Partial<WorkDocumentImageProperties>,
  options: DocumentImageCommandOptions = {},
): boolean {
  return editor.commands.setDocumentImageProperties(value, options);
}

export function setDocumentImageLayoutOptions(
  editor: Editor,
  value: Partial<WorkDocumentImageLayoutOptions>,
  options: DocumentImageCommandOptions = {},
): boolean {
  return editor.commands.setDocumentImageLayoutOptions(value, options);
}

export function setDocumentImageAlternativeText(
  editor: Editor,
  alternativeText: string,
  options: DocumentImageCommandOptions = {},
): boolean {
  return editor.commands.setDocumentImageAlternativeText(
    alternativeText,
    options,
  );
}

export function documentImageAlternativeText(editor: Editor): string {
  if (!editor.isActive('image')) return '';
  const attributes = editor.getAttributes('image') as Record<string, unknown>;
  const value =
    typeof attributes.alt === 'string'
      ? attributes.alt
      : typeof attributes.title === 'string'
        ? attributes.title
        : '';
  return value.trim();
}

export function documentImageLayoutFromElement(
  element: Element,
): WorkDocumentImageLayoutOptions {
  return normalizeDocumentImageLayoutOptions({
    layout: element.getAttribute('data-office-image-layout'),
    alignment: element.getAttribute('data-office-image-alignment'),
    wrapDistance: element.getAttribute('data-office-image-wrap-distance'),
    wrapSide: element.getAttribute('data-office-image-wrap-side'),
  });
}

export function documentImagePositionFromElement(
  element: Element,
): WorkDocumentImagePosition | null {
  return normalizeDocumentImagePosition({
    horizontalOffset: element.getAttribute(
      'data-office-image-horizontal-offset',
    ),
    verticalOffset: element.getAttribute('data-office-image-vertical-offset'),
    horizontalReference: element.getAttribute(
      'data-office-image-horizontal-reference',
    ),
    verticalReference: element.getAttribute(
      'data-office-image-vertical-reference',
    ),
  });
}

export function documentImageCropFromElement(
  element: Element,
): WorkDocumentImageCrop | null {
  return normalizeDocumentImageCrop({
    cropTop: element.getAttribute('data-office-image-crop-top'),
    cropRight: element.getAttribute('data-office-image-crop-right'),
    cropBottom: element.getAttribute('data-office-image-crop-bottom'),
    cropLeft: element.getAttribute('data-office-image-crop-left'),
  });
}

export function documentImageLayerFromElement(
  element: Element,
): WorkDocumentImageLayer {
  return normalizeDocumentImageLayer({
    relativeHeight: element.getAttribute('data-office-image-relative-height'),
    behindDocument: element.getAttribute('data-office-image-behind-document'),
    allowOverlap: element.getAttribute('data-office-image-allow-overlap'),
    layoutInCell: element.getAttribute('data-office-image-layout-in-cell'),
    lockAnchor: element.getAttribute('data-office-image-lock-anchor'),
  });
}

export function normalizeDocumentImageLayoutOptions(
  value: Partial<Record<keyof WorkDocumentImageLayoutOptions, unknown>>,
): WorkDocumentImageLayoutOptions {
  return {
    layout: normalizeDocumentImageLayout(value.layout),
    alignment: normalizeDocumentImageAlignment(value.alignment),
    wrapDistance: normalizeDocumentImageWrapDistance(value.wrapDistance),
    wrapSide: normalizeDocumentImageWrapSide(value.wrapSide),
  };
}

export function normalizeDocumentImageLayout(
  value: unknown,
): WorkDocumentImageLayout {
  return value === 'square' ||
    value === 'tight' ||
    value === 'through' ||
    value === 'topBottom'
    ? value
    : DEFAULT_IMAGE_LAYOUT;
}

export function normalizeDocumentImageAlignment(
  value: unknown,
): WorkDocumentImageAlignment {
  return value === 'left' || value === 'right'
    ? value
    : DEFAULT_IMAGE_ALIGNMENT;
}

export function normalizeDocumentImageWrapDistance(value: unknown): number {
  const number =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value)
        : DEFAULT_WRAP_DISTANCE_MILLIMETERS;
  if (!Number.isFinite(number)) return DEFAULT_WRAP_DISTANCE_MILLIMETERS;
  return (
    Math.round(
      Math.min(MAX_WRAP_DISTANCE_MILLIMETERS, Math.max(0, number)) * 2,
    ) / 2
  );
}

export function normalizeDocumentImageDimension(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) && number > 0
    ? Math.round(number * 100) / 100
    : null;
}

export function normalizeDocumentImageLockAspectRatio(value: unknown): boolean {
  return value === false || value === 'false' || value === 0 || value === '0'
    ? false
    : DEFAULT_LOCK_ASPECT_RATIO;
}

export function normalizeDocumentImagePosition(
  value: Partial<Record<keyof WorkDocumentImagePosition, unknown>>,
): WorkDocumentImagePosition | null {
  const horizontalOffset = normalizeDocumentImageOffset(value.horizontalOffset);
  const verticalOffset = normalizeDocumentImageOffset(value.verticalOffset);
  if (horizontalOffset === null && verticalOffset === null) return null;
  return {
    horizontalOffset,
    verticalOffset,
    horizontalReference: normalizeDocumentImageHorizontalReference(
      value.horizontalReference,
    ),
    verticalReference: normalizeDocumentImageVerticalReference(
      value.verticalReference,
    ),
  };
}

export function normalizeDocumentImageOffset(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(number)) return null;
  return (
    Math.round(
      Math.min(
        MAX_IMAGE_OFFSET_MILLIMETERS,
        Math.max(-MAX_IMAGE_OFFSET_MILLIMETERS, number),
      ) * 100,
    ) / 100
  );
}

export function normalizeDocumentImageHorizontalReference(
  value: unknown,
): WorkDocumentImageHorizontalReference {
  return value === 'margin' || value === 'page'
    ? value
    : DEFAULT_HORIZONTAL_REFERENCE;
}

export function normalizeDocumentImageVerticalReference(
  value: unknown,
): WorkDocumentImageVerticalReference {
  return value === 'margin' || value === 'page'
    ? value
    : DEFAULT_VERTICAL_REFERENCE;
}

export function normalizeDocumentImageCrop(
  value: Partial<
    Record<'cropBottom' | 'cropLeft' | 'cropRight' | 'cropTop', unknown>
  >,
): WorkDocumentImageCrop | null {
  const crop = {
    top: normalizeDocumentImageCropEdge(value.cropTop),
    right: normalizeDocumentImageCropEdge(value.cropRight),
    bottom: normalizeDocumentImageCropEdge(value.cropBottom),
    left: normalizeDocumentImageCropEdge(value.cropLeft),
  };
  if (Object.values(crop).every((edge) => edge === 0)) return null;
  if (crop.left + crop.right >= 100 || crop.top + crop.bottom >= 100) {
    return null;
  }
  return crop;
}

export function normalizeDocumentImageCropEdge(value: unknown): number {
  const number = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(number)) return 0;
  return Math.round(Math.min(99.99, Math.max(0, number)) * 100) / 100;
}

export function normalizeDocumentImageLayer(
  value: Partial<Record<keyof WorkDocumentImageLayer, unknown>>,
): WorkDocumentImageLayer {
  return {
    relativeHeight: normalizeDocumentImageRelativeHeight(value.relativeHeight),
    behindDocument: normalizeDocumentImageLayerBoolean(
      value.behindDocument,
      DEFAULT_IMAGE_LAYER.behindDocument,
    ),
    allowOverlap: normalizeDocumentImageLayerBoolean(
      value.allowOverlap,
      DEFAULT_IMAGE_LAYER.allowOverlap,
    ),
    layoutInCell: normalizeDocumentImageLayerBoolean(
      value.layoutInCell,
      DEFAULT_IMAGE_LAYER.layoutInCell,
    ),
    lockAnchor: normalizeDocumentImageLayerBoolean(
      value.lockAnchor,
      DEFAULT_IMAGE_LAYER.lockAnchor,
    ),
  };
}

export function normalizeDocumentImageRelativeHeight(value: unknown): number {
  const number =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value)
        : DEFAULT_IMAGE_LAYER.relativeHeight;
  if (!Number.isFinite(number)) return DEFAULT_IMAGE_LAYER.relativeHeight;
  return Math.round(
    Math.min(MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT, Math.max(0, number)),
  );
}

export function documentImageLayerCssZIndex(
  layer: WorkDocumentImageLayer,
): number {
  const rank = Math.round(
    (layer.relativeHeight / MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT) *
      DOCUMENT_IMAGE_CSS_LAYER_RANGE,
  );
  return layer.behindDocument
    ? rank - DOCUMENT_IMAGE_CSS_LAYER_RANGE - 1
    : rank + 1;
}

function formatImageLayoutNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function documentImageAttributesForChanges(
  value: Partial<WorkDocumentImageProperties>,
): Record<string, unknown> {
  const attributes: Record<string, unknown> = {};
  if (Object.hasOwn(value, 'width')) {
    attributes.width = normalizeDocumentImageDimension(value.width);
  }
  if (Object.hasOwn(value, 'height')) {
    attributes.height = normalizeDocumentImageDimension(value.height);
  }
  if (Object.hasOwn(value, 'lockAspectRatio')) {
    attributes.lockAspectRatio = normalizeDocumentImageLockAspectRatio(
      value.lockAspectRatio,
    );
  }
  if (Object.hasOwn(value, 'layout')) {
    attributes.layout = normalizeDocumentImageLayout(value.layout);
  }
  if (Object.hasOwn(value, 'alignment')) {
    attributes.alignment = normalizeDocumentImageAlignment(value.alignment);
  }
  if (Object.hasOwn(value, 'wrapDistance')) {
    attributes.wrapDistance = normalizeDocumentImageWrapDistance(
      value.wrapDistance,
    );
  }
  if (Object.hasOwn(value, 'wrapSide')) {
    attributes.wrapSide = normalizeDocumentImageWrapSide(value.wrapSide);
  }
  if (Object.hasOwn(value, 'alternativeText')) {
    const alternativeText = value.alternativeText?.trim() ?? '';
    attributes.alt = alternativeText || null;
    attributes.title = alternativeText || null;
  }
  if (Object.hasOwn(value, 'position')) {
    const position = value.position
      ? normalizeDocumentImagePosition(value.position)
      : null;
    attributes.horizontalOffset = position?.horizontalOffset ?? null;
    attributes.verticalOffset = position?.verticalOffset ?? null;
    attributes.horizontalReference =
      position?.horizontalReference ?? DEFAULT_HORIZONTAL_REFERENCE;
    attributes.verticalReference =
      position?.verticalReference ?? DEFAULT_VERTICAL_REFERENCE;
  }
  if (Object.hasOwn(value, 'crop')) {
    const crop = value.crop
      ? normalizeDocumentImageCrop({
          cropTop: value.crop.top,
          cropRight: value.crop.right,
          cropBottom: value.crop.bottom,
          cropLeft: value.crop.left,
        })
      : null;
    attributes.cropTop = crop?.top ?? 0;
    attributes.cropRight = crop?.right ?? 0;
    attributes.cropBottom = crop?.bottom ?? 0;
    attributes.cropLeft = crop?.left ?? 0;
  }
  if (Object.hasOwn(value, 'contour')) {
    const contour = value.contour
      ? normalizeDocumentImageWrapContour(value.contour)
      : null;
    attributes.wrapPolygon = contour
      ? serializeDocumentImageWrapPoints(contour.points)
      : null;
    attributes.wrapPolygonEdited = contour?.edited ?? false;
  }
  if (Object.hasOwn(value, 'layer')) {
    const layer = value.layer
      ? normalizeDocumentImageLayer(value.layer)
      : DEFAULT_IMAGE_LAYER;
    attributes.relativeHeight = layer.relativeHeight;
    attributes.behindDocument = layer.behindDocument;
    attributes.allowOverlap = layer.allowOverlap;
    attributes.layoutInCell = layer.layoutInCell;
    attributes.lockAnchor = layer.lockAnchor;
  }
  return attributes;
}

function isDocumentImageSelection(selection: unknown): boolean {
  return (
    selection instanceof NodeSelection && selection.node.type.name === 'image'
  );
}

function sameDocumentImageResizeConfiguration(
  initial: ProseMirrorNode,
  current: ProseMirrorNode,
): boolean {
  return (
    normalizeDocumentImageDimension(initial.attrs.width) ===
      normalizeDocumentImageDimension(current.attrs.width) &&
    normalizeDocumentImageDimension(initial.attrs.height) ===
      normalizeDocumentImageDimension(current.attrs.height) &&
    normalizeDocumentImageLockAspectRatio(initial.attrs.lockAspectRatio) ===
      normalizeDocumentImageLockAspectRatio(current.attrs.lockAspectRatio)
  );
}

function applyInitialImageHtmlAttributes(
  element: HTMLImageElement,
  attributes: Record<string, unknown>,
): void {
  const merged = mergeAttributes(attributes);
  for (const [name, value] of Object.entries(merged)) {
    if (
      value === null ||
      value === undefined ||
      name === 'width' ||
      name === 'height'
    )
      continue;
    element.setAttribute(name, String(value));
  }
}

function syncDocumentImageNodeView(
  element: HTMLImageElement,
  container: HTMLElement | null,
  attributes: Record<string, unknown>,
): void {
  const layout = normalizeDocumentImageLayout(attributes.layout);
  const alignment = normalizeDocumentImageAlignment(attributes.alignment);
  const wrapDistance = normalizeDocumentImageWrapDistance(
    attributes.wrapDistance,
  );
  const lockAspectRatio = normalizeDocumentImageLockAspectRatio(
    attributes.lockAspectRatio,
  );
  const position = normalizeDocumentImagePosition(attributes);
  const crop = normalizeDocumentImageCrop(attributes);
  const wrapSide = normalizeDocumentImageWrapSide(attributes.wrapSide);
  const contour = effectiveDocumentImageWrapContour(attributes);
  const layer = normalizeDocumentImageLayer(attributes);
  setOptionalImageAttribute(element, 'src', attributes.src);
  setOptionalImageAttribute(element, 'alt', attributes.alt);
  setOptionalImageAttribute(element, 'title', attributes.title);
  element.dataset.officeImageLayout = layout;
  element.dataset.officeImageAlignment = alignment;
  element.dataset.officeImageWrapDistance =
    formatImageLayoutNumber(wrapDistance);
  element.dataset.officeImageLockAspectRatio = String(lockAspectRatio);
  element.dataset.officeImageWrapSide = wrapSide;
  syncDocumentImagePosition(element, position);
  applyDocumentImageCropToElement(element, crop);
  applyDocumentImageWrapContourToElement(element, contour);
  applyDocumentImageLayerToElement(element, layer);
  element.style.setProperty(
    '--work-document-image-wrap-distance',
    `${formatImageLayoutNumber(wrapDistance)}mm`,
  );
  syncImageDimension(element, 'width', attributes.width);
  syncImageDimension(element, 'height', attributes.height);
  if (!container) return;
  container.dataset.officeImageLayout = layout;
  container.dataset.officeImageAlignment = alignment;
  container.dataset.officeImageWrapDistance =
    formatImageLayoutNumber(wrapDistance);
  container.dataset.officeImageLockAspectRatio = String(lockAspectRatio);
  container.dataset.officeImageWrapSide = wrapSide;
  syncDocumentImagePosition(container, position);
  applyDocumentImageCropToElement(container, crop);
  applyDocumentImageWrapContourToElement(container, contour);
  applyDocumentImageLayerToElement(container, layer);
  container.style.setProperty(
    '--work-document-image-wrap-distance',
    `${formatImageLayoutNumber(wrapDistance)}mm`,
  );
}

function imagePositionNumberAttribute(name: string) {
  return {
    default: null,
    parseHTML: (element: Element) =>
      normalizeDocumentImageOffset(element.getAttribute(name)),
    renderHTML: (attributes: Record<string, unknown>) => {
      const key = name.includes('horizontal')
        ? 'horizontalOffset'
        : 'verticalOffset';
      const offset = normalizeDocumentImageOffset(attributes[key]);
      if (offset === null) return {};
      const formatted = formatImageLayoutNumber(offset);
      const variable = name.includes('horizontal')
        ? '--work-document-image-horizontal-offset'
        : '--work-document-image-vertical-offset';
      return { [name]: formatted, style: `${variable}:${formatted}mm` };
    },
  };
}

function imageCropNumberAttribute(edge: keyof WorkDocumentImageCrop) {
  const name = `data-office-image-crop-${edge}`;
  const key = `crop${edge[0]?.toUpperCase()}${edge.slice(1)}`;
  return {
    default: 0,
    parseHTML: (element: Element) =>
      normalizeDocumentImageCropEdge(element.getAttribute(name)),
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = normalizeDocumentImageCropEdge(attributes[key]);
      const result: Record<string, string> = {};
      if (value > 0) result[name] = formatImageLayoutNumber(value);
      if (edge !== 'top') return result;
      const crop = normalizeDocumentImageCrop({
        cropTop: attributes.cropTop,
        cropRight: attributes.cropRight,
        cropBottom: attributes.cropBottom,
        cropLeft: attributes.cropLeft,
      });
      if (crop) result.style = documentImageCropStyle(crop);
      return result;
    },
  };
}

function imageWrapPolygonAttribute() {
  return {
    default: null,
    parseHTML: (element: Element) => {
      const contour = documentImageWrapContourFromElement(element);
      return contour ? serializeDocumentImageWrapPoints(contour.points) : null;
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      const contour = effectiveDocumentImageWrapContour(attributes);
      if (!contour) return {};
      return {
        'data-office-image-wrap-polygon': serializeDocumentImageWrapPoints(
          contour.points,
        ),
        'data-office-image-wrap-polygon-edited': String(contour.edited),
        style: `--work-document-image-wrap-contour:${documentImageWrapContourCss(contour)}`,
      };
    },
  };
}

function imageRelativeHeightAttribute() {
  return {
    default: DEFAULT_IMAGE_LAYER.relativeHeight,
    parseHTML: (element: Element) =>
      normalizeDocumentImageRelativeHeight(
        element.getAttribute('data-office-image-relative-height'),
      ),
    renderHTML: (attributes: Record<string, unknown>) => {
      const layer = normalizeDocumentImageLayer(attributes);
      return {
        'data-office-image-relative-height': String(layer.relativeHeight),
        style: `--work-document-image-z-index:${documentImageLayerCssZIndex(layer)}`,
      };
    },
  };
}

function imageLayerBooleanAttribute(
  key: keyof Omit<WorkDocumentImageLayer, 'relativeHeight'>,
  name: string,
  fallback: boolean,
) {
  return {
    default: fallback,
    parseHTML: (element: Element) =>
      normalizeDocumentImageLayerBoolean(element.getAttribute(name), fallback),
    renderHTML: (attributes: Record<string, unknown>) => ({
      [name]: String(
        normalizeDocumentImageLayerBoolean(attributes[key], fallback),
      ),
    }),
  };
}

function effectiveDocumentImageWrapContour(
  attributes: Record<string, unknown>,
): WorkDocumentImageWrapContour | null {
  const contour = normalizeDocumentImageWrapContour(attributes);
  if (contour) return contour;
  return isContourImageLayout(normalizeDocumentImageLayout(attributes.layout))
    ? defaultDocumentImageWrapContour()
    : null;
}

export function isContourImageLayout(
  layout: WorkDocumentImageLayout,
): layout is 'through' | 'tight' {
  return layout === 'tight' || layout === 'through';
}

export function wrapsBesideImage(layout: WorkDocumentImageLayout): boolean {
  return layout === 'square' || isContourImageLayout(layout);
}

function syncDocumentImagePosition(
  element: HTMLElement,
  position: WorkDocumentImagePosition | null,
): void {
  if (!position) {
    delete element.dataset.officeImageHorizontalOffset;
    delete element.dataset.officeImageVerticalOffset;
    delete element.dataset.officeImageHorizontalReference;
    delete element.dataset.officeImageVerticalReference;
    element.style.removeProperty('--work-document-image-horizontal-offset');
    element.style.removeProperty('--work-document-image-vertical-offset');
    return;
  }
  if (position.horizontalOffset === null) {
    delete element.dataset.officeImageHorizontalOffset;
    element.style.removeProperty('--work-document-image-horizontal-offset');
  } else {
    element.dataset.officeImageHorizontalOffset = formatImageLayoutNumber(
      position.horizontalOffset,
    );
    element.style.setProperty(
      '--work-document-image-horizontal-offset',
      `${formatImageLayoutNumber(position.horizontalOffset)}mm`,
    );
  }
  if (position.verticalOffset === null) {
    delete element.dataset.officeImageVerticalOffset;
    element.style.removeProperty('--work-document-image-vertical-offset');
  } else {
    element.dataset.officeImageVerticalOffset = formatImageLayoutNumber(
      position.verticalOffset,
    );
    element.style.setProperty(
      '--work-document-image-vertical-offset',
      `${formatImageLayoutNumber(position.verticalOffset)}mm`,
    );
  }
  element.dataset.officeImageHorizontalReference = position.horizontalReference;
  element.dataset.officeImageVerticalReference = position.verticalReference;
}

export function applyDocumentImageCropToElement(
  element: HTMLElement,
  crop: WorkDocumentImageCrop | null,
): void {
  for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
    const datasetKey = `officeImageCrop${edge[0].toUpperCase()}${edge.slice(1)}`;
    const variable = `--work-document-image-crop-${edge}`;
    if (!crop || crop[edge] === 0) {
      delete element.dataset[datasetKey];
      element.style.removeProperty(variable);
      continue;
    }
    const formatted = formatImageLayoutNumber(crop[edge]);
    element.dataset[datasetKey] = formatted;
    element.style.setProperty(variable, `${formatted}%`);
  }
  if (!crop) {
    element.style.removeProperty('--work-document-image-crop-scale-x');
    element.style.removeProperty('--work-document-image-crop-scale-y');
    element.style.removeProperty('--work-document-image-crop-translate-x');
    element.style.removeProperty('--work-document-image-crop-translate-y');
    return;
  }
  for (const [name, value] of documentImageCropStyleEntries(crop)) {
    element.style.setProperty(name, value);
  }
}

export function applyDocumentImageLayerToElement(
  element: HTMLElement,
  layer: WorkDocumentImageLayer,
): void {
  const normalized = normalizeDocumentImageLayer(layer);
  element.dataset.officeImageRelativeHeight = String(normalized.relativeHeight);
  element.dataset.officeImageBehindDocument = String(normalized.behindDocument);
  element.dataset.officeImageAllowOverlap = String(normalized.allowOverlap);
  element.dataset.officeImageLayoutInCell = String(normalized.layoutInCell);
  element.dataset.officeImageLockAnchor = String(normalized.lockAnchor);
  element.style.setProperty(
    '--work-document-image-z-index',
    String(documentImageLayerCssZIndex(normalized)),
  );
}

function normalizeDocumentImageLayerBoolean(
  value: unknown,
  fallback: boolean,
): boolean {
  if (value === true || value === 1 || value === '1' || value === 'true') {
    return true;
  }
  if (value === false || value === 0 || value === '0' || value === 'false') {
    return false;
  }
  return fallback;
}

function documentImageCropStyle(crop: WorkDocumentImageCrop): string {
  return documentImageCropStyleEntries(crop)
    .map(([name, value]) => `${name}:${value}`)
    .join(';');
}

function documentImageCropStyleEntries(
  crop: WorkDocumentImageCrop,
): Array<readonly [string, string]> {
  const visibleWidth = 100 - crop.left - crop.right;
  const visibleHeight = 100 - crop.top - crop.bottom;
  return [
    ...(['top', 'right', 'bottom', 'left'] as const).map(
      (edge) =>
        [
          `--work-document-image-crop-${edge}`,
          `${formatImageLayoutNumber(crop[edge])}%`,
        ] as const,
    ),
    [
      '--work-document-image-crop-scale-x',
      formatImageLayoutNumber(100 / visibleWidth),
    ],
    [
      '--work-document-image-crop-scale-y',
      formatImageLayoutNumber(100 / visibleHeight),
    ],
    [
      '--work-document-image-crop-translate-x',
      `${formatImageLayoutNumber((-crop.left * 100) / visibleWidth)}%`,
    ],
    [
      '--work-document-image-crop-translate-y',
      `${formatImageLayoutNumber((-crop.top * 100) / visibleHeight)}%`,
    ],
  ];
}

function setOptionalImageAttribute(
  element: HTMLImageElement,
  name: 'alt' | 'src' | 'title',
  value: unknown,
): void {
  if (typeof value === 'string' && value) element.setAttribute(name, value);
  else element.removeAttribute(name);
}

function syncImageDimension(
  element: HTMLImageElement,
  dimension: 'height' | 'width',
  value: unknown,
): void {
  const pixels = typeof value === 'number' ? value : Number(value);
  if (Number.isFinite(pixels) && pixels > 0) {
    element.style[dimension] = `${pixels}px`;
    element.setAttribute(dimension, String(Math.round(pixels)));
    return;
  }
  element.style.removeProperty(dimension);
  element.removeAttribute(dimension);
}
