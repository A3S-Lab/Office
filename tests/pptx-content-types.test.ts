import { expect, test } from '@rstest/core';
import PptxGenJS from 'pptxgenjs';
import { createPptxBlob } from '../src/internal/features/work/work-pptx-export';
import {
  attribute,
  directChildren,
  OoxmlPackage,
} from '../src/internal/features/work/work-ooxml-package';
import type {
  WorkArtifact,
  WorkSlide,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('a saved multi-slide PPTX does not declare missing slide masters', async () => {
  const blob = await createPptxBlob(threeSlideArtifact(), PptxGenJS);
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const contentTypes = await archive.xml('[Content_Types].xml');
  const missing = directChildren(contentTypes.documentElement, 'Override')
    .map((node) => attribute(node, 'PartName') ?? '')
    .filter((partName) => partName.startsWith('/'))
    .filter((partName) => !archive.has(partName.slice(1)));

  expect(missing).toEqual([]);
  expect(archive.has('ppt/slideMasters/slideMaster1.xml')).toBe(true);
  expect(archive.has('ppt/slideMasters/slideMaster2.xml')).toBe(false);
});

function threeSlideArtifact(): WorkArtifact {
  return {
    id: 'three-slide-deck',
    kind: 'presentation',
    title: 'Three slides',
    favorite: false,
    createdAt: 0,
    updatedAt: 0,
    lastOpenedAt: 0,
    revision: 0,
    content: {
      type: 'presentation',
      slides: [
        slide('slide-1', '季度数据'),
        slide('slide-2', ''),
        slide('slide-3', '幻灯'),
      ],
    },
  };
}

function slide(id: string, text: string): WorkSlide {
  const elements: WorkSlideElement[] = text
    ? [
        {
          id: `${id}-title`,
          type: 'text',
          x: 10,
          y: 12,
          width: 80,
          height: 16,
          text,
          fontSize: 28,
          color: '#172033',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
      ]
    : [];
  return {
    id,
    name: text || id,
    background: '#ffffff',
    elements,
  };
}
