import { withPresentationDesign } from './work-presentation-layouts';
import { presentationGroupPath } from './work-presentation-groups';
import type { PptxGroupExportRegistry } from './work-pptx-groups';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlideElement,
} from './work-types';

type PptxConstructor = typeof import('pptxgenjs').default;
type PptxPresentation = InstanceType<PptxConstructor>;
type PptxSlideMaster = Parameters<PptxPresentation['defineSlideMaster']>[0];
type PptxSlideMasterObject = NonNullable<PptxSlideMaster['objects']>[number];

interface PptxLayoutElement {
  element: WorkSlideElement;
  groupScope: string;
}

export interface PptxLayoutBinding {
  masterName: string;
  placeholderNames: Map<string, string>;
}

export function definePptxSlideLayouts(
  presentation: PptxPresentation,
  source: WorkPresentationContent,
  slideWidth: number,
  slideHeight: number,
  groups: PptxGroupExportRegistry,
): {
  content: WorkPresentationContent;
  bindings: Map<string, PptxLayoutBinding>;
} {
  const content = withPresentationDesign(source);
  const bindings = new Map<string, PptxLayoutBinding>();
  const usedNames = new Set<string>();
  for (const layout of content.layouts ?? []) {
    const master = content.masters?.find(
      (candidate) => candidate.id === layout.masterId,
    );
    const masterName = uniqueMasterName(layout.name, usedNames);
    const placeholders = effectivePlaceholders(master, layout);
    const inherited = scopedLayoutElements(master, layout).filter(
      ({ element }) => !element.placeholder,
    );
    const groupedPlaceholders = placeholders.filter(
      ({ element }) => presentationGroupPath(element).length,
    );
    const nativePlaceholders = placeholders.filter(
      ({ element }) => !presentationGroupPath(element).length,
    );
    const placeholderNames = new Map(
      nativePlaceholders.map(({ element }, index) => [
        element.placeholder?.key ?? `placeholder:${index + 1}`,
        placeholderName(element, index),
      ]),
    );
    const objects = [
      ...inherited.flatMap(({ element, groupScope }) =>
        slideMasterObjects(
          element,
          slideWidth,
          slideHeight,
          groups,
          groupScope,
        ),
      ),
      ...groupedPlaceholders.flatMap(({ element, groupScope }) =>
        slideMasterObjects(
          element,
          slideWidth,
          slideHeight,
          groups,
          groupScope,
        ),
      ),
      ...nativePlaceholders.map(({ element }, index) =>
        slideMasterPlaceholder(
          element,
          placeholderNames.get(
            element.placeholder?.key ?? `placeholder:${index + 1}`,
          ) ?? placeholderName(element, index),
          slideWidth,
          slideHeight,
        ),
      ),
    ];
    presentation.defineSlideMaster({
      title: masterName,
      background: {
        color: colorValue(layout.background ?? master?.background ?? '#ffffff'),
      },
      objects,
    });
    bindings.set(layout.id, { masterName, placeholderNames });
  }
  return { content, bindings };
}

function effectivePlaceholders(
  master: WorkPresentationMaster | undefined,
  layout: WorkPresentationLayout,
): PptxLayoutElement[] {
  const placeholders = new Map<string, PptxLayoutElement>();
  for (const item of scopedLayoutElements(master, layout)) {
    if (item.element.placeholder) {
      placeholders.set(item.element.placeholder.key, item);
    }
  }
  return Array.from(placeholders.values());
}

function scopedLayoutElements(
  master: WorkPresentationMaster | undefined,
  layout: WorkPresentationLayout,
): PptxLayoutElement[] {
  return [
    ...(layout.showMasterElements === false
      ? []
      : (master?.elements ?? []).map((element) => ({
          element,
          groupScope: `master:${master?.id ?? layout.masterId}`,
        }))),
    ...layout.elements.map((element) => ({
      element,
      groupScope: `layout:${layout.id}`,
    })),
  ];
}

function slideMasterObjects(
  element: WorkSlideElement,
  slideWidth: number,
  slideHeight: number,
  groups: PptxGroupExportRegistry,
  groupScope: string,
): PptxSlideMasterObject[] {
  const box = elementBox(element, slideWidth, slideHeight);
  if (element.type === 'image' && element.image) {
    const objectName = groups.objectName(groupScope, element, 'image');
    return [
      {
        image: {
          data: element.image.dataUrl,
          ...box,
          rotate: element.rotation,
          altText: element.altText,
          ...(objectName ? { objectName } : {}),
        },
      },
    ];
  }
  if (element.type === 'line') {
    const objectName = groups.objectName(groupScope, element, 'line');
    return [
      {
        line: {
          ...box,
          rotate: element.rotation,
          line: {
            color: colorValue(element.borderColor ?? element.color),
            width: element.borderWidth ?? 1,
          },
          ...(objectName ? { objectName } : {}),
        },
      },
    ];
  }
  if (element.type === 'text') {
    const objectName = groups.objectName(groupScope, element, 'text');
    return [
      {
        text: {
          text: element.text,
          options: textOptions(element, slideWidth, slideHeight, objectName),
        },
      },
    ];
  }
  if (element.type !== 'shape') return [];
  const shapeObjectName = groups.objectName(groupScope, element, 'shape');
  const objects: PptxSlideMasterObject[] = [
    {
      rect: {
        ...box,
        rotate: element.rotation,
        fill:
          element.fill === 'transparent'
            ? { color: 'FFFFFF', transparency: 100 }
            : {
                color: colorValue(element.fill),
                transparency: Math.round((1 - (element.opacity ?? 1)) * 100),
              },
        line: {
          color: colorValue(element.borderColor ?? element.fill),
          width: element.borderWidth ?? 0,
          transparency: element.borderWidth ? 0 : 100,
        },
        ...(shapeObjectName ? { objectName: shapeObjectName } : {}),
      },
    },
  ];
  if (element.text) {
    const textObjectName = groups.objectName(groupScope, element, 'text');
    objects.push({
      text: {
        text: element.text,
        options: textOptions(element, slideWidth, slideHeight, textObjectName),
      },
    });
  }
  return objects;
}

function slideMasterPlaceholder(
  element: WorkSlideElement,
  name: string,
  slideWidth: number,
  slideHeight: number,
): PptxSlideMasterObject {
  return {
    placeholder: {
      text: element.placeholder?.prompt ?? element.text,
      options: {
        name,
        type: placeholderType(element.placeholder?.type),
        ...textOptions(element, slideWidth, slideHeight),
      },
    },
  };
}

function textOptions(
  element: WorkSlideElement,
  slideWidth: number,
  slideHeight: number,
  objectName?: string,
) {
  return {
    ...elementBox(element, slideWidth, slideHeight),
    rotate: element.rotation,
    fontFace: element.fontFamily ?? 'Aptos',
    fontSize: Math.max(8, element.fontSize * 0.75),
    color: colorValue(element.color),
    bold: element.bold,
    italic: element.italic,
    underline: element.underline ? ({ style: 'sng' } as const) : undefined,
    align: element.align,
    valign: element.verticalAlign ?? ('middle' as const),
    margin: 0,
    ...(objectName ? { objectName } : {}),
  };
}

function elementBox(
  element: WorkSlideElement,
  slideWidth: number,
  slideHeight: number,
) {
  return {
    x: (element.x / 100) * slideWidth,
    y: (element.y / 100) * slideHeight,
    w: (element.width / 100) * slideWidth,
    h: (element.height / 100) * slideHeight,
  };
}

function placeholderName(element: WorkSlideElement, index: number): string {
  const type = element.placeholder?.type || 'body';
  const key =
    element.placeholder?.key.replace(/[^a-z0-9_-]/gi, '-') || `${index + 1}`;
  return `A3S-${type}-${key}`;
}

function placeholderType(
  type: string | undefined,
): 'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media' {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === 'pic') return 'pic';
  if (type === 'chart') return 'chart';
  if (type === 'tbl') return 'tbl';
  if (type === 'media') return 'media';
  return 'body';
}

function uniqueMasterName(name: string, used: Set<string>): string {
  const base = name.trim() || 'Slide layout';
  let candidate = base;
  let index = 2;
  while (used.has(candidate)) {
    candidate = `${base} ${index}`;
    index += 1;
  }
  used.add(candidate);
  return candidate;
}

function colorValue(color: string): string {
  return color.replace(/^#/, '').toUpperCase();
}
