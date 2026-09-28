import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  initializeOfficePresentationCollaboration,
  type PresentationContent,
  readOfficePresentationCollaboration,
  replaceOfficePresentationCollaboration,
} from '../src/core';
import { presentationCollaborationFixture } from './fixtures/presentation-collaboration';

test('a stale slide move leaves the order array and shape text alone', () => {
  const session = startedPresentation('presentation-stale-slide-move');
  const withThird = appendSlide(readOfficePresentationCollaboration(session));
  replaceOfficePresentationCollaboration(
    session,
    readOfficePresentationCollaboration(session),
    withThird,
  );
  const observed = readOfficePresentationCollaboration(session);
  const remote = moveSlide(observed, 'slide-3', 0);
  replaceOfficePresentationCollaboration(session, observed, remote);
  const titles = shapeText(readOfficePresentationCollaboration(session));

  replaceOfficePresentationCollaboration(
    session,
    observed,
    moveSlide(observed, 'slide-2', 0),
  );

  const current = readOfficePresentationCollaboration(session);
  expect(current.slides.map((slide) => slide.id)).toEqual([
    'slide-3',
    'slide-1',
    'slide-2',
  ]);
  expect(shapeText(current)).toEqual(titles);
});

test('a stale group path writes no member and leaves an unrelated shape', () => {
  const session = startedPresentation('presentation-stale-group');
  const withPeer = addShape(
    readOfficePresentationCollaboration(session),
    'slide-1',
    'element-a',
    'Alpha',
  );
  replaceOfficePresentationCollaboration(
    session,
    readOfficePresentationCollaboration(session),
    withPeer,
  );
  const observed = readOfficePresentationCollaboration(session);
  const remote = setGroupIds(observed, 'slide-1', ['element-a'], ['remote']);
  replaceOfficePresentationCollaboration(session, observed, remote);

  replaceOfficePresentationCollaboration(
    session,
    observed,
    setGroupIds(observed, 'slide-1', ['element-title', 'element-a'], ['local']),
  );

  const current = readOfficePresentationCollaboration(session);
  const slide = current.slides.find((item) => item.id === 'slide-1');
  expect(groupIds(slide, 'element-title')).toEqual([]);
  expect(groupIds(slide, 'element-a')).toEqual(['remote']);
  expect(shapeText(current).get('element-body')).toBe('Body');
  expect(shapeText(current).get('element-title')).toBe('Shared presentation');
});

test('a stale background writes nothing on a slide, master, or layout', () => {
  const session = startedPresentation('presentation-stale-background');
  const observed = readOfficePresentationCollaboration(session);
  const remote = setBackgrounds(observed, {
    slide: '#111111',
    master: '#222222',
    layout: '#333333',
  });
  replaceOfficePresentationCollaboration(session, observed, remote);
  const titles = shapeText(readOfficePresentationCollaboration(session));

  replaceOfficePresentationCollaboration(
    session,
    observed,
    setBackgrounds(observed, {
      slide: '#010203',
      master: '#020304',
      layout: '#030405',
      useLayoutBackground: false,
    }),
  );

  const current = readOfficePresentationCollaboration(session);
  expect(current.slides[0]?.background).toBe('#111111');
  expect(current.slides[0]?.useLayoutBackground).toBeUndefined();
  expect(current.slides[1]?.background).toBe('#FFFFFF');
  expect(current.masters?.[0]?.background).toBe('#222222');
  expect(current.layouts?.[0]?.background).toBe('#333333');
  expect(shapeText(current)).toEqual(titles);
});

test('a stale element order writes nothing and a fresh order writes', () => {
  const session = startedPresentation('presentation-stale-element-order');
  const seeded = addShape(
    addShape(
      readOfficePresentationCollaboration(session),
      'slide-1',
      'element-a',
      'Alpha',
    ),
    'slide-1',
    'element-b',
    'Beta',
  );
  replaceOfficePresentationCollaboration(
    session,
    readOfficePresentationCollaboration(session),
    seeded,
  );
  const observed = readOfficePresentationCollaboration(session);
  expect(elementIds(observed, 'slide-1')).toEqual([
    'element-title',
    'element-a',
    'element-b',
  ]);
  const remote = reorderElements(observed, 'slide-1', [
    'element-b',
    'element-title',
    'element-a',
  ]);
  replaceOfficePresentationCollaboration(session, observed, remote);
  const titles = shapeText(readOfficePresentationCollaboration(session));

  replaceOfficePresentationCollaboration(
    session,
    observed,
    reorderElements(observed, 'slide-1', [
      'element-a',
      'element-title',
      'element-b',
    ]),
  );

  const current = readOfficePresentationCollaboration(session);
  expect(elementIds(current, 'slide-1')).toEqual([
    'element-b',
    'element-title',
    'element-a',
  ]);
  expect(shapeText(current)).toEqual(titles);

  const freshBase = readOfficePresentationCollaboration(session);
  replaceOfficePresentationCollaboration(
    session,
    freshBase,
    reorderElements(freshBase, 'slide-1', [
      'element-title',
      'element-a',
      'element-b',
    ]),
  );
  expect(
    elementIds(readOfficePresentationCollaboration(session), 'slide-1'),
  ).toEqual(['element-title', 'element-a', 'element-b']);
  expect(
    shapeText(readOfficePresentationCollaboration(session)).get('element-a'),
  ).toBe('Alpha');
  expect(
    shapeText(readOfficePresentationCollaboration(session)).get('element-b'),
  ).toBe('Beta');
});

test('derived thumbnails and layout measurements stay out of the replica', () => {
  const session = startedPresentation('presentation-local-derived');
  const observed = readOfficePresentationCollaboration(session);
  const before = Y.encodeStateVector(session.document);
  const derivedOnly = withDerivedMeasurements(observed, 'Shared presentation');
  expect(
    replaceOfficePresentationCollaboration(session, observed, derivedOnly),
  ).toBe(false);
  expect(Y.encodeStateVector(session.document)).toEqual(before);

  const edited = withDerivedMeasurements(observed, 'Shared presentation!');
  expect(
    replaceOfficePresentationCollaboration(session, observed, edited),
  ).toBe(true);
  const current = readOfficePresentationCollaboration(session);
  const stored = JSON.stringify(current);
  expect(stored).not.toContain('THUMB');
  expect(stored).not.toContain('measuredBox');
  expect(stored).not.toContain('lineBoxes');
  expect(stored).not.toContain('measuredLayout');
  expect(current.slides[0]?.elements[0]?.text).toBe('Shared presentation!');
  expect(current.slides[0]?.elements[0]?.x).toBe(10);
  expect(current.slides[0]?.elements[0]?.width).toBe(80);
  const bytes = Y.encodeStateAsUpdate(session.document);
  expect(new TextDecoder().decode(bytes)).not.toContain('THUMB');
});

function withDerivedMeasurements(
  content: PresentationContent,
  text: string,
): PresentationContent {
  return {
    ...content,
    slides: content.slides.map((slide, index) =>
      index === 0
        ? ({
            ...slide,
            thumbnail: 'data:image/png;base64,THUMB',
            measuredLayout: { width: 12, height: 8 },
            elements: slide.elements.map((element, elementIndex) =>
              elementIndex === 0
                ? {
                    ...element,
                    text,
                    measuredBox: { left: 1, top: 2, width: 3, height: 4 },
                    lineBoxes: [{ from: 0, to: 1 }],
                  }
                : element,
            ),
          } as PresentationContent['slides'][number])
        : slide,
    ),
  };
}

function startedPresentation(artifactId: string) {
  const session = createOfficeCollaborationSession({
    artifactId,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(
    session,
    presentationCollaborationFixture(),
  );
  return session;
}

function appendSlide(content: PresentationContent): PresentationContent {
  return {
    ...content,
    slides: [
      ...content.slides,
      {
        id: 'slide-3',
        name: 'Notes',
        background: '#FFFFFF',
        layoutId: 'layout-1',
        elements: [
          {
            id: 'element-note',
            type: 'text',
            x: 12,
            y: 12,
            width: 40,
            height: 16,
            text: 'Note',
            fontSize: 16,
            color: '#172033',
            fill: 'transparent',
            bold: false,
            align: 'left',
          },
        ],
      },
    ],
  };
}

function elementIds(content: PresentationContent, slideId: string): string[] {
  return (
    content.slides
      .find((slide) => slide.id === slideId)
      ?.elements.map((element) => element.id) ?? []
  );
}

function reorderElements(
  content: PresentationContent,
  slideId: string,
  ids: readonly string[],
): PresentationContent {
  return {
    ...content,
    slides: content.slides.map((slide) => {
      if (slide.id !== slideId) return slide;
      const byId = new Map(
        slide.elements.map((element) => [element.id, element]),
      );
      return {
        ...slide,
        elements: ids.map((id) => {
          const element = byId.get(id);
          if (!element) throw new Error(`missing ${id}`);
          return element;
        }),
      };
    }),
  };
}

function moveSlide(
  content: PresentationContent,
  slideId: string,
  index: number,
): PresentationContent {
  const slides = [...content.slides];
  const current = slides.findIndex((slide) => slide.id === slideId);
  const [slide] = slides.splice(current, 1);
  if (!slide) throw new Error(`missing ${slideId}`);
  slides.splice(index, 0, slide);
  return { ...content, slides };
}

function addShape(
  content: PresentationContent,
  slideId: string,
  elementId: string,
  text: string,
): PresentationContent {
  return {
    ...content,
    slides: content.slides.map((slide) =>
      slide.id === slideId
        ? {
            ...slide,
            elements: [
              ...slide.elements,
              {
                id: elementId,
                type: 'shape' as const,
                x: 8,
                y: 8,
                width: 20,
                height: 12,
                text,
                fontSize: 14,
                color: '#172033',
                fill: '#F8FAFC',
                bold: false,
                align: 'left' as const,
              },
            ],
          }
        : slide,
    ),
  };
}

function setGroupIds(
  content: PresentationContent,
  slideId: string,
  elementIds: readonly string[],
  groupIds: readonly string[],
): PresentationContent {
  const selected = new Set(elementIds);
  return {
    ...content,
    slides: content.slides.map((slide) =>
      slide.id === slideId
        ? {
            ...slide,
            elements: slide.elements.map((element) =>
              selected.has(element.id)
                ? { ...element, groupIds: [...groupIds] }
                : element,
            ),
          }
        : slide,
    ),
  };
}

function setBackgrounds(
  content: PresentationContent,
  next: {
    slide: string;
    master: string;
    layout: string;
    useLayoutBackground?: false;
  },
): PresentationContent {
  return {
    ...content,
    slides: content.slides.map((slide, index) =>
      index === 0
        ? {
            ...slide,
            background: next.slide,
            ...(next.useLayoutBackground === false
              ? { useLayoutBackground: false as const }
              : {}),
          }
        : slide,
    ),
    masters: content.masters?.map((master, index) =>
      index === 0 ? { ...master, background: next.master } : master,
    ),
    layouts: content.layouts?.map((layout, index) =>
      index === 0 ? { ...layout, background: next.layout } : layout,
    ),
  };
}

function shapeText(content: PresentationContent): Map<string, string> {
  return new Map(
    content.slides.flatMap((slide) =>
      slide.elements.map((element) => [element.id, element.text] as const),
    ),
  );
}

function groupIds(
  slide: PresentationContent['slides'][number] | undefined,
  elementId: string,
): string[] {
  return (
    slide?.elements.find((element) => element.id === elementId)?.groupIds ?? []
  );
}
