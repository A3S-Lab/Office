import { type Editor, mergeAttributes, ResizableNodeView } from '@tiptap/core';
import Image from '@tiptap/extension-image';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export type WorkDocumentImageLayout = 'inline' | 'square' | 'topBottom';
export type WorkDocumentImageAlignment = 'left' | 'center' | 'right';

export interface WorkDocumentImageLayoutOptions {
  layout: WorkDocumentImageLayout;
  alignment: WorkDocumentImageAlignment;
  wrapDistance: number;
}

const DEFAULT_IMAGE_LAYOUT: WorkDocumentImageLayout = 'inline';
const DEFAULT_IMAGE_ALIGNMENT: WorkDocumentImageAlignment = 'center';
const DEFAULT_WRAP_DISTANCE_MILLIMETERS = 3;
const MAX_WRAP_DISTANCE_MILLIMETERS = 25;

export const DocumentImage = Image.extend({
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
          sync(updatedNode);
          return true;
        },
        options: {
          directions,
          min: { width: minWidth, height: minHeight },
          preserveAspectRatio: alwaysPreserveAspectRatio === true,
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

export function setDocumentImageLayoutOptions(
  editor: Editor,
  value: Partial<WorkDocumentImageLayoutOptions>,
  options: { restoreFocus?: boolean } = {},
): boolean {
  if (!editor.isActive('image')) return false;
  const current = documentImageLayoutOptions(editor);
  const next = normalizeDocumentImageLayoutOptions({
    ...current,
    ...value,
  });
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  return chain.updateAttributes('image', next).run();
}

export function setDocumentImageAlternativeText(
  editor: Editor,
  alternativeText: string,
): boolean {
  if (!editor.isActive('image')) return false;
  const normalized = alternativeText.trim();
  return editor
    .chain()
    .focus()
    .updateAttributes('image', {
      alt: normalized || null,
      title: normalized || null,
    })
    .run();
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
  });
}

export function normalizeDocumentImageLayoutOptions(
  value: Partial<Record<keyof WorkDocumentImageLayoutOptions, unknown>>,
): WorkDocumentImageLayoutOptions {
  return {
    layout: normalizeDocumentImageLayout(value.layout),
    alignment: normalizeDocumentImageAlignment(value.alignment),
    wrapDistance: normalizeDocumentImageWrapDistance(value.wrapDistance),
  };
}

export function normalizeDocumentImageLayout(
  value: unknown,
): WorkDocumentImageLayout {
  return value === 'square' || value === 'topBottom'
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

function formatImageLayoutNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
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
  setOptionalImageAttribute(element, 'src', attributes.src);
  setOptionalImageAttribute(element, 'alt', attributes.alt);
  setOptionalImageAttribute(element, 'title', attributes.title);
  element.dataset.officeImageLayout = layout;
  element.dataset.officeImageAlignment = alignment;
  element.dataset.officeImageWrapDistance =
    formatImageLayoutNumber(wrapDistance);
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
  container.style.setProperty(
    '--work-document-image-wrap-distance',
    `${formatImageLayoutNumber(wrapDistance)}mm`,
  );
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
