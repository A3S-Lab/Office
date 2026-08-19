import { expect, test } from '@rstest/core';
import {
  presentationSlideViewFromDesign,
  withPresentationDesign,
  withPresentationDesignMetadata,
} from '../src/internal/features/work/work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkSlide,
} from '../src/internal/features/work/work-types';

test('resolves a slide view from normalized design data without traversing the deck', () => {
  const content = presentationContent();
  const designContent = withPresentationDesign(content);
  const slide = designContent.slides[0];
  if (!slide) throw new Error('The presentation fixture requires a slide.');
  const guardedDesignContent = Object.defineProperty(
    { ...designContent },
    'slides',
    {
      configurable: true,
      get(): WorkSlide[] {
        throw new Error('The slide view traversed the complete deck.');
      },
    },
  );

  expect(
    presentationSlideViewFromDesign(guardedDesignContent, slide),
  ).toMatchObject({
    background: '#f5f7fa',
    inheritedElements: [
      expect.objectContaining({ id: 'master-decoration' }),
      expect.objectContaining({ id: 'layout-decoration' }),
    ],
    layout: expect.objectContaining({ id: 'layout-title' }),
    master: expect.objectContaining({ id: 'master-default' }),
  });
});

test('normalizes presentation design metadata without copying slides', () => {
  const content = presentationContent();
  const slides = content.slides;

  const designContent = withPresentationDesignMetadata(content);

  expect(designContent.slides).toBe(slides);
  expect(designContent.layouts).toHaveLength(1);
  expect(designContent.masters).toHaveLength(1);
});

function presentationContent(): WorkPresentationContent {
  return {
    type: 'presentation',
    masters: [
      {
        id: 'master-default',
        name: 'Default',
        background: '#ffffff',
        elements: [element('master-decoration')],
      },
    ],
    layouts: [
      {
        id: 'layout-title',
        name: 'Title',
        masterId: 'master-default',
        background: '#f5f7fa',
        elements: [element('layout-decoration')],
      },
    ],
    slides: [
      {
        id: 'slide-1',
        name: 'Slide 1',
        background: '#ffffff',
        layoutId: 'layout-title',
        useLayoutBackground: true,
        elements: [],
      },
    ],
  };
}

function element(id: string) {
  return {
    id,
    type: 'shape' as const,
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    text: '',
    fontSize: 12,
    color: '#000000',
    fill: '#ffffff',
    bold: false,
    align: 'left' as const,
  };
}
