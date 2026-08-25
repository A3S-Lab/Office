import JSZip from 'jszip';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from './work-ooxml-package';
import type { PptxGroupExportRegistry } from './work-pptx-groups';
import {
  normalizeWorkSlideAnimation,
  WORK_SLIDE_ANIMATION_LIMIT,
} from './work-presentation-animation';
import { createWorkId } from './work-templates';
import type {
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideAnimationDirection,
  WorkSlideAnimationEffect,
  WorkSlideAnimationTrigger,
} from './work-types';

const PRESENTATIONML_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

interface PptxAnimationDiagnostic {
  code: string;
  message: string;
}

interface PptxAnimationReadResult {
  animations: WorkSlideAnimation[];
  diagnostics: PptxAnimationDiagnostic[];
}

interface PptxAnimationEffectResult {
  direction?: WorkSlideAnimationDirection;
  effect?: WorkSlideAnimationEffect;
}

interface PptxAnimationExportTarget {
  animation: WorkSlideAnimation;
  spid: number;
}

interface PptxAnimationGroup {
  automatic: boolean;
  wrappers: WorkSlideAnimation[][];
}

export function readPptxAnimations(
  document: Document,
  elementIdByPptxShapeId: ReadonlyMap<string, string>,
): PptxAnimationReadResult {
  const timing = directChild(document.documentElement, 'timing');
  if (!timing) return { animations: [], diagnostics: [] };
  const diagnostics: PptxAnimationDiagnostic[] = [];
  if (elementNamespace(timing) !== PRESENTATIONML_NAMESPACE) {
    return {
      animations: [],
      diagnostics: [
        {
          code: 'pptx.animation.namespace',
          message:
            'A namespace-spoofed animation timing tree was ignored instead of being treated as trusted PresentationML.',
        },
      ],
    };
  }

  const animations: WorkSlideAnimation[] = [];
  const animatedElements = new Set<string>();
  const effectNodes = pptxDescendants(timing, 'cTn').filter((node) =>
    ['clickEffect', 'withEffect', 'afterEffect'].includes(
      attribute(node, 'nodeType') ?? '',
    ),
  );
  if (effectNodes.length > WORK_SLIDE_ANIMATION_LIMIT) {
    diagnostics.push({
      code: 'pptx.animation.limit',
      message: `Only the first ${WORK_SLIDE_ANIMATION_LIMIT} object animation timing records on a slide are inspected and imported.`,
    });
  }
  for (const node of effectNodes.slice(0, WORK_SLIDE_ANIMATION_LIMIT)) {
    if (attribute(node, 'presetClass') !== 'entr') {
      diagnostics.push({
        code: 'pptx.animation.class',
        message:
          'An emphasis, exit, motion-path, or malformed animation remains in the source PPTX only.',
      });
      continue;
    }
    const targets = Array.from(
      new Set(
        pptxDescendants(node, 'spTgt')
          .map((target) => attribute(target, 'spid')?.trim())
          .filter((target): target is string => Boolean(target)),
      ),
    );
    if (targets.length !== 1) {
      diagnostics.push({
        code: 'pptx.animation.target',
        message:
          'An entrance animation with a missing or ambiguous object target was ignored.',
      });
      continue;
    }
    const elementId = elementIdByPptxShapeId.get(targets[0]);
    if (!elementId) {
      diagnostics.push({
        code: 'pptx.animation.target',
        message:
          'An entrance animation points to an unavailable slide object and was ignored.',
      });
      continue;
    }
    if (animatedElements.has(elementId)) {
      diagnostics.push({
        code: 'pptx.animation.duplicate-target',
        message:
          'Only the first supported entrance animation for an object is editable; additional effects remain in the source PPTX only.',
      });
      continue;
    }
    const effectResult = readPptxAnimationEffect(node);
    if (!effectResult.effect) {
      diagnostics.push({
        code: 'pptx.animation.effect',
        message:
          'An entrance animation uses an unsupported or malformed effect and remains in the source PPTX only.',
      });
      continue;
    }
    const timingResult = readPptxAnimationTiming(node);
    if (timingResult.normalized) {
      diagnostics.push({
        code: 'pptx.animation.timing',
        message:
          'An invalid or out-of-range animation duration or delay was normalized to safe editable bounds.',
      });
    }
    const animation = normalizeWorkSlideAnimation({
      id: createWorkId('slide-animation'),
      elementId,
      effect: effectResult.effect,
      trigger: pptxAnimationTrigger(attribute(node, 'nodeType')),
      durationMs: timingResult.durationMs,
      delayMs: timingResult.delayMs,
      direction: effectResult.direction,
    });
    animations.push(animation);
    animatedElements.add(elementId);
  }

  if (!effectNodes.length) {
    diagnostics.push({
      code: 'pptx.animation.structure',
      message:
        'The slide timing tree does not contain a supported object entrance sequence and remains in the source PPTX only.',
    });
  }
  return { animations, diagnostics };
}

export async function patchPptxAnimations(
  buffer: ArrayBuffer,
  slides: readonly WorkSlide[],
  registry: PptxGroupExportRegistry,
): Promise<ArrayBuffer> {
  if (!slides.some((slide) => slide.animations?.length)) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  for (const [index, slide] of slides.entries()) {
    if (!slide.animations?.length) continue;
    const path = `ppt/slides/slide${index + 1}.xml`;
    const entry = archive.file(path);
    if (!entry) {
      throw new Error(
        `PPTX animation export is missing generated slide ${index + 1}.`,
      );
    }
    const document = parseXml(await entry.async('text'), path);
    normalizePptxShapeIds(document);
    const targetByMarker = pptxTargetIdsByMarker(document);
    const scope = `slide:${slide.id}`;
    const targets: PptxAnimationExportTarget[] = slide.animations.map(
      (animation) => {
        const marker = registry.markerForElement(scope, animation.elementId);
        const ids = marker ? targetByMarker.get(marker) : undefined;
        if (!marker || ids?.length !== 1) {
          throw new Error(
            `PPTX animation export expected one generated object for animation ${animation.id}, but found ${ids?.length ?? 0}.`,
          );
        }
        return { animation, spid: ids[0] };
      },
    );
    writePptxAnimations(document, targets);
    archive.file(path, new XMLSerializer().serializeToString(document));
  }
  return archive.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function writePptxAnimations(
  document: Document,
  targets: readonly PptxAnimationExportTarget[],
): void {
  const root = document.documentElement;
  directChild(root, 'timing')?.remove();
  if (!targets.length) return;
  const source = parseXml(pptxAnimationTimingXml(targets), 'PPTX animations');
  const timing = clonePptxAnimationElement(document, source.documentElement);
  const extensionList = directChild(root, 'extLst') ?? null;
  root.insertBefore(timing, extensionList);
}

function clonePptxAnimationElement(
  document: Document,
  source: Element,
): Element {
  const element = document.createElementNS(
    PRESENTATIONML_NAMESPACE,
    `p:${source.localName}`,
  );
  for (const item of Array.from(source.attributes)) {
    if (item.name === 'xmlns' || item.name.startsWith('xmlns:')) continue;
    element.setAttribute(item.name, item.value);
  }
  for (const child of Array.from(source.childNodes)) {
    element.append(
      child instanceof Element
        ? clonePptxAnimationElement(document, child)
        : document.createTextNode(child.textContent ?? ''),
    );
  }
  return element;
}

function readPptxAnimationEffect(node: Element): PptxAnimationEffectResult {
  const presetId = attribute(node, 'presetID');
  const animationEffects = pptxDescendants(node, 'animEffect');
  const filters = animationEffects
    .map((effect) => attribute(effect, 'filter')?.trim())
    .filter((filter): filter is string => Boolean(filter));
  const attributes = new Set(
    pptxDescendants(node, 'attrName').map((name) => name.textContent?.trim()),
  );
  if (filters.some((filter) => filter === 'fade')) return { effect: 'fade' };
  for (const filter of filters) {
    const direction = directionFromPptxSlideFilter(filter);
    if (!direction) continue;
    return {
      effect: 'fly-in',
      direction,
    };
  }
  if (
    presetId === '23' ||
    pptxDescendants(node, 'animScale').length > 0 ||
    (attributes.has('ppt_w') && attributes.has('ppt_h'))
  ) {
    return { effect: 'zoom' };
  }
  if (
    presetId === '2' &&
    (attributes.has('ppt_x') || attributes.has('ppt_y'))
  ) {
    return {
      effect: 'fly-in',
      direction: directionFromPptxMotion(node),
    };
  }
  if (
    presetId === '1' &&
    pptxDescendants(node, 'set').length > 0 &&
    animationEffects.length === 0 &&
    pptxDescendants(node, 'anim').length === 0 &&
    pptxDescendants(node, 'animScale').length === 0
  ) {
    return { effect: 'appear' };
  }
  return {};
}

function readPptxAnimationTiming(node: Element): {
  delayMs: number;
  durationMs: number;
  normalized: boolean;
} {
  const directDuration = nonnegativeInteger(attribute(node, 'dur'));
  const behaviorDurations = pptxDirectChildren(
    pptxDirectChild(node, 'childTnLst') ?? node,
  )
    .flatMap((behavior) => pptxDescendants(behavior, 'cTn'))
    .map((time) => nonnegativeInteger(attribute(time, 'dur')))
    .filter((duration): duration is number => duration !== undefined);
  const behaviorDuration = Math.max(0, ...behaviorDurations);
  const rawDuration = directDuration ?? (behaviorDuration || 500);
  const delayValue = attribute(
    pptxChildPath(node, 'stCondLst', 'cond') ?? node,
    'delay',
  );
  const rawDelay = nonnegativeInteger(delayValue) ?? 0;
  const normalized = normalizeWorkSlideAnimation({
    id: 'timing',
    elementId: 'timing',
    effect: 'fade',
    trigger: 'on-click',
    durationMs: rawDuration,
    delayMs: rawDelay,
  });
  return {
    durationMs: normalized.durationMs,
    delayMs: normalized.delayMs,
    normalized:
      normalized.durationMs !== rawDuration ||
      normalized.delayMs !== rawDelay ||
      (attribute(node, 'dur') !== null && directDuration === undefined) ||
      (delayValue !== null && nonnegativeInteger(delayValue) === undefined),
  };
}

function pptxAnimationTrigger(value: string | null): WorkSlideAnimationTrigger {
  if (value === 'withEffect') return 'with-previous';
  if (value === 'afterEffect') return 'after-previous';
  return 'on-click';
}

function directionFromPptxSlideFilter(
  filter: string,
): WorkSlideAnimationDirection | undefined {
  if (filter === 'slide(fromLeft)') return 'left';
  if (filter === 'slide(fromRight)') return 'right';
  if (filter === 'slide(fromTop)') return 'up';
  if (filter === 'slide(fromBottom)') return 'down';
  return undefined;
}

function directionFromPptxMotion(node: Element): WorkSlideAnimationDirection {
  const values = pptxDescendants(node, 'strVal')
    .map((value) => attribute(value, 'val') ?? '')
    .join(' ');
  if (values.includes('-#ppt_w/2')) return 'right';
  if (values.includes('1+#ppt_h/2')) return 'down';
  if (values.includes('-#ppt_h/2')) return 'up';
  return 'left';
}

function pptxAnimationTimingXml(
  targets: readonly PptxAnimationExportTarget[],
): string {
  let timeNodeId = 1;
  const nextId = () => timeNodeId++;
  const rootId = nextId();
  const mainSequenceId = nextId();
  const targetByAnimationId = new Map(
    targets.map((target) => [target.animation.id, target]),
  );
  const groups = pptxAnimationGroups(targets.map((target) => target.animation));
  const groupXml = groups
    .map((group) => {
      const outerId = nextId();
      const wrappers = group.wrappers
        .map((animations) => {
          const wrapperId = nextId();
          const effects = animations
            .map((animation) => {
              const target = targetByAnimationId.get(animation.id);
              if (!target) return '';
              return pptxAnimationEffectXml(target, nextId);
            })
            .join('');
          return `<p:par><p:cTn id="${wrapperId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`;
        })
        .join('');
      const delay = group.automatic ? '0' : 'indefinite';
      return `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${wrappers}</p:childTnLst></p:cTn></p:par>`;
    })
    .join('');
  const builds = targets
    .map(({ spid }) => `<p:bldP spid="${spid}" grpId="0" animBg="1"/>`)
    .join('');
  return (
    `<p:timing xmlns:p="${PRESENTATIONML_NAMESPACE}"><p:tnLst><p:par>` +
    `<p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${mainSequenceId}" dur="indefinite" nodeType="mainSeq">` +
    `<p:childTnLst>${groupXml}</p:childTnLst></p:cTn>` +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
    `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${builds}</p:bldLst></p:timing>`
  );
}

function pptxAnimationGroups(
  animations: readonly WorkSlideAnimation[],
): PptxAnimationGroup[] {
  const groups: PptxAnimationGroup[] = [];
  let current: PptxAnimationGroup | undefined;
  for (const animation of animations) {
    if (!current || animation.trigger === 'on-click') {
      current = {
        automatic: animation.trigger !== 'on-click',
        wrappers: [[animation]],
      };
      groups.push(current);
    } else if (animation.trigger === 'with-previous') {
      current.wrappers.at(-1)?.push(animation);
    } else {
      current.wrappers.push([animation]);
    }
  }
  return groups;
}

function pptxAnimationEffectXml(
  target: PptxAnimationExportTarget,
  nextId: () => number,
): string {
  const animation = normalizeWorkSlideAnimation(target.animation);
  const preset = pptxAnimationPreset(animation);
  const leafId = nextId();
  const visibilityId = nextId();
  const visibility =
    `<p:set><p:cBhvr><p:cTn id="${visibilityId}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
    `<p:tgtEl><p:spTgt spid="${target.spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>` +
    '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>';
  const behavior = pptxAnimationBehaviorXml(animation, target.spid, nextId);
  return (
    `<p:par><p:cTn id="${leafId}" presetID="${preset.id}" presetClass="entr" presetSubtype="${preset.subtype}" ` +
    `fill="hold" nodeType="${pptxAnimationNodeType(animation.trigger)}" dur="${animation.durationMs}">` +
    `<p:stCondLst><p:cond delay="${animation.delayMs}"/></p:stCondLst><p:childTnLst>${visibility}${behavior}</p:childTnLst>` +
    '</p:cTn></p:par>'
  );
}

function pptxAnimationBehaviorXml(
  animation: WorkSlideAnimation,
  spid: number,
  nextId: () => number,
): string {
  if (animation.effect === 'appear') return '';
  if (animation.effect === 'zoom') {
    return ['ppt_w', 'ppt_h']
      .map((attributeName) => {
        const id = nextId();
        return (
          `<p:anim calcMode="lin" valueType="num"><p:cBhvr><p:cTn id="${id}" dur="${animation.durationMs}" fill="hold"/>` +
          `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>${attributeName}</p:attrName></p:attrNameLst></p:cBhvr>` +
          `<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#${attributeName}"/></p:val></p:tav></p:tavLst></p:anim>`
        );
      })
      .join('');
  }
  const id = nextId();
  const filter =
    animation.effect === 'fade'
      ? 'fade'
      : pptxSlideFilter(animation.direction ?? 'left');
  return (
    `<p:animEffect transition="in" filter="${filter}"><p:cBhvr><p:cTn id="${id}" dur="${animation.durationMs}"/>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`
  );
}

function pptxAnimationPreset(animation: WorkSlideAnimation): {
  id: number;
  subtype: number;
} {
  if (animation.effect === 'appear') return { id: 1, subtype: 0 };
  if (animation.effect === 'fade') return { id: 10, subtype: 0 };
  if (animation.effect === 'zoom') return { id: 23, subtype: 16 };
  const subtypes: Record<WorkSlideAnimationDirection, number> = {
    up: 1,
    right: 2,
    down: 4,
    left: 8,
  };
  return { id: 2, subtype: subtypes[animation.direction ?? 'left'] };
}

function pptxAnimationNodeType(
  trigger: WorkSlideAnimationTrigger,
): 'afterEffect' | 'clickEffect' | 'withEffect' {
  if (trigger === 'with-previous') return 'withEffect';
  if (trigger === 'after-previous') return 'afterEffect';
  return 'clickEffect';
}

function pptxSlideFilter(direction: WorkSlideAnimationDirection): string {
  if (direction === 'right') return 'slide(fromRight)';
  if (direction === 'up') return 'slide(fromTop)';
  if (direction === 'down') return 'slide(fromBottom)';
  return 'slide(fromLeft)';
}

function normalizePptxShapeIds(document: Document): void {
  for (const shapeTree of descendants(document, 'spTree')) {
    const used = new Set<number>();
    let candidate = 1;
    for (const properties of descendants(shapeTree, 'cNvPr')) {
      const current = Number(attribute(properties, 'id'));
      if (Number.isInteger(current) && current > 0 && !used.has(current)) {
        used.add(current);
        continue;
      }
      while (used.has(candidate)) candidate += 1;
      properties.setAttribute('id', String(candidate));
      used.add(candidate);
      candidate += 1;
    }
  }
}

function pptxTargetIdsByMarker(document: Document): Map<string, number[]> {
  const targets = new Map<string, number[]>();
  for (const properties of descendants(document, 'cNvPr')) {
    const marker = attribute(properties, 'name');
    const id = Number(attribute(properties, 'id'));
    if (!marker || !Number.isInteger(id) || id <= 0) continue;
    const ids = targets.get(marker) ?? [];
    ids.push(id);
    targets.set(marker, ids);
  }
  return targets;
}

function nonnegativeInteger(value: string | null): number | undefined {
  if (value === null || value.trim() === '') return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : undefined;
}

function elementNamespace(element: Element): string | null {
  if (element.namespaceURI) return element.namespaceURI;
  const separator = element.tagName.indexOf(':');
  const prefix =
    element.prefix ??
    (separator >= 0 ? element.tagName.slice(0, separator) : undefined);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  let current: Element | null = element;
  while (current) {
    const namespace = current.getAttribute(declaration);
    if (namespace) return namespace;
    current = current.parentElement;
  }
  return null;
}

function pptxDescendants(parent: ParentNode, localName: string): Element[] {
  return descendants(parent, localName).filter(
    (element) => elementNamespace(element) === PRESENTATIONML_NAMESPACE,
  );
}

function pptxDirectChildren(parent: ParentNode, localName?: string): Element[] {
  return directChildren(parent, localName).filter(
    (element) => elementNamespace(element) === PRESENTATIONML_NAMESPACE,
  );
}

function pptxDirectChild(
  parent: ParentNode,
  localName: string,
): Element | undefined {
  return pptxDirectChildren(parent, localName)[0];
}

function pptxChildPath(
  parent: ParentNode,
  ...localNames: string[]
): Element | undefined {
  let current: ParentNode | undefined = parent;
  for (const localName of localNames) {
    if (!current) return undefined;
    current = pptxDirectChild(current, localName);
  }
  return current instanceof Element ? current : undefined;
}
