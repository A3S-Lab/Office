import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { createPptxBlob } from '../src/internal/features/work/work-pptx-export';
import { importPptxPresentation } from '../src/internal/features/work/work-pptx-import';
import {
  attribute,
  childPath,
  descendants,
  directChild,
  directChildren,
  OoxmlPackage,
} from '../src/internal/features/work/work-ooxml-package';
import type {
  WorkArtifact,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('exports nested browser groups as native PPTX group nodes', async () => {
  const blob = await createPptxBlob(groupedArtifact(), PptxGenJS);
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const slide = await archive.xml('ppt/slides/slide1.xml');
  const slideXml = new XMLSerializer().serializeToString(slide);
  const shapeTree = descendants(slide, 'spTree')[0];
  const sceneNodes = directChildren(shapeTree).filter((node) =>
    ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'].includes(node.localName),
  );
  const groups = descendants(shapeTree, 'grpSp');

  expect(sceneNodes.map((node) => node.localName)).toEqual(['sp', 'grpSp']);
  expect(groups).toHaveLength(2);
  expect(slideXml).not.toContain('__A3S_OFFICE_EXPORT__');
  expectUniqueNonVisualIds(shapeTree);
  for (const group of groups) {
    const transform = childPath(group, 'grpSpPr', 'xfrm');
    const offset = directChild(transform ?? group, 'off');
    const extent = directChild(transform ?? group, 'ext');
    const childOffset = directChild(transform ?? group, 'chOff');
    const childExtent = directChild(transform ?? group, 'chExt');
    expect(attribute(offset ?? group, 'x')).toBe(
      attribute(childOffset ?? group, 'x'),
    );
    expect(attribute(offset ?? group, 'y')).toBe(
      attribute(childOffset ?? group, 'y'),
    );
    expect(attribute(extent ?? group, 'cx')).toBe(
      attribute(childExtent ?? group, 'cx'),
    );
    expect(attribute(extent ?? group, 'cy')).toBe(
      attribute(childExtent ?? group, 'cy'),
    );
  }

  const imported = await importPptxPresentation(
    await happyDomImportableFile(blob),
  );
  const elements = imported.content.slides[0]?.elements ?? [];
  const grouped = elements.filter((element) => element.groupIds?.length);
  const ungrouped = elements.filter((element) => !element.groupIds?.length);

  expect(grouped.map((element) => element.groupIds?.length).sort()).toEqual([
    1, 1, 1, 1, 2, 2,
  ]);
  expect(new Set(grouped.map((element) => element.groupIds?.[0])).size).toBe(1);
  expect(
    new Set(
      grouped
        .filter((element) => element.groupIds?.length === 2)
        .map((element) => element.groupIds?.[1]),
    ).size,
  ).toBe(1);
  expect(grouped.map((element) => element.type).sort()).toEqual([
    'chart',
    'image',
    'line',
    'shape',
    'table',
    'text',
  ]);
  expect(ungrouped).toHaveLength(1);
  expect(ungrouped[0]?.type).toBe('text');
  expect(imported.compatibility.issues).not.toContainEqual(
    expect.objectContaining({ code: 'pptx.group' }),
  );
});

test('keeps master and layout group scopes separate and materializes grouped placeholders', async () => {
  const blob = await createPptxBlob(groupedDesignArtifact(), PptxGenJS);
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const layoutPaths = archive
    .paths('ppt/slideLayouts/')
    .filter((path) => /slideLayout\d+\.xml$/.test(path));
  const layouts = await Promise.all(
    layoutPaths.map(async (path) => ({
      document: await archive.xml(path),
      path,
    })),
  );
  const design = layouts.find(
    ({ document }) => descendants(document, 'grpSp').length > 0,
  );

  if (!design) throw new Error('Generated PPTX is missing its grouped layout.');
  const shapeTree = descendants(design.document, 'spTree')[0];
  const rootGroups = directChildren(shapeTree, 'grpSp');
  expect(rootGroups).toHaveLength(3);
  expect(descendants(design.document, 'grpSp')).toHaveLength(3);
  expect(descendants(design.document, 'ph')).toHaveLength(1);
  expectUniqueNonVisualIds(shapeTree);
  expect(
    rootGroups.some((group) => group.textContent?.includes('Grouped prompt')),
  ).toBe(true);

  for (const path of archive
    .paths('ppt/')
    .filter((candidate) =>
      /^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(candidate),
    )) {
    expect(await archive.text(path)).not.toContain('__A3S_OFFICE_EXPORT__');
  }
});

function expectUniqueNonVisualIds(shapeTree: Element): void {
  const ids = descendants(shapeTree, 'cNvPr').map((node) =>
    attribute(node, 'id'),
  );
  expect(ids.every(Boolean)).toBe(true);
  expect(new Set(ids).size).toBe(ids.length);
}

async function happyDomImportableFile(blob: Blob): Promise<File> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const path = 'ppt/presentation.xml';
  const entry = archive.file(path);
  if (!entry) throw new Error('Generated PPTX is missing its presentation.');
  const xml = await entry.async('text');
  archive.file(
    path,
    xml.replace(
      /(<p:sldId\b[^>]*?)\sid="[^"]*"([^>]*?)\sr:id="([^"]+)"/g,
      '$1 id="$3"$2',
    ),
  );
  return new File(
    [await archive.generateAsync({ type: 'uint8array' })],
    'grouped-export.pptx',
    { type: blob.type },
  );
}

function groupedDesignArtifact(): WorkArtifact {
  const masterGroup = [
    element('master-shape', 'shape', 4, 6, 16, 10, ['shared']),
    element('master-text', 'text', 22, 6, 18, 10, ['shared']),
  ];
  const layoutGroup = [
    element('layout-shape', 'shape', 4, 22, 16, 10, ['shared']),
    element('layout-text', 'text', 22, 22, 18, 10, ['shared']),
  ];
  const groupedPlaceholder = {
    ...element('grouped-placeholder', 'text', 4, 40, 24, 12, [
      'placeholder-group',
    ]),
    text: 'Grouped prompt',
    placeholder: {
      key: 'body:grouped',
      type: 'body',
      prompt: 'Grouped prompt',
    },
  };
  const groupedPeer = element(
    'grouped-placeholder-peer',
    'text',
    30,
    40,
    24,
    12,
    ['placeholder-group'],
  );
  const nativePlaceholder = {
    ...element('native-placeholder', 'text', 4, 64, 50, 12),
    text: '',
    placeholder: {
      key: 'body:native',
      type: 'body',
      prompt: 'Native prompt',
    },
  };
  return {
    id: 'design-artifact',
    kind: 'presentation',
    title: 'Native design groups',
    favorite: false,
    createdAt: 0,
    updatedAt: 0,
    lastOpenedAt: 0,
    revision: 0,
    content: {
      type: 'presentation',
      width: 10,
      height: 5.625,
      masters: [
        {
          id: 'master',
          name: 'Master',
          background: '#ffffff',
          elements: masterGroup,
        },
      ],
      layouts: [
        {
          id: 'layout',
          name: 'Layout',
          masterId: 'master',
          elements: [
            ...layoutGroup,
            groupedPlaceholder,
            groupedPeer,
            nativePlaceholder,
          ],
        },
      ],
      slides: [
        {
          id: 'slide',
          name: 'Design slide',
          background: '#ffffff',
          layoutId: 'layout',
          elements: [],
        },
      ],
    },
  };
}

function groupedArtifact(): WorkArtifact {
  return {
    id: 'artifact',
    kind: 'presentation',
    title: 'Native groups',
    favorite: false,
    createdAt: 0,
    updatedAt: 0,
    lastOpenedAt: 0,
    revision: 0,
    content: {
      type: 'presentation',
      width: 10,
      height: 5.625,
      slides: [
        {
          id: 'slide',
          name: 'Grouped slide',
          background: '#ffffff',
          elements: [
            element('outer-shape', 'shape', 8, 12, 16, 14, ['outer']),
            element('outside', 'text', 28, 18, 18, 12),
            element('inner-text', 'text', 50, 22, 20, 12, ['outer', 'inner']),
            element('inner-line', 'line', 52, 44, 22, 2, ['outer', 'inner']),
            {
              ...element('outer-table', 'table', 12, 52, 32, 20, ['outer']),
              table: {
                headerRows: 1,
                rows: [
                  ['Item', 'Value'],
                  ['A', '12'],
                ],
              },
            },
            {
              ...element('outer-image', 'image', 78, 10, 10, 12, ['outer']),
              image: {
                dataUrl:
                  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=',
                contentType: 'image/png',
                name: 'pixel.png',
              },
            },
            {
              ...element('outer-chart', 'chart', 62, 54, 28, 22, ['outer']),
              chart: {
                type: 'column',
                categories: ['A'],
                series: [{ name: 'Value', values: [12] }],
              },
            },
          ],
        },
      ],
    },
  };
}

function element(
  id: string,
  type: WorkSlideElement['type'],
  x: number,
  y: number,
  width: number,
  height: number,
  groupIds?: string[],
): WorkSlideElement {
  return {
    id,
    type,
    x,
    y,
    width,
    height,
    text: type === 'text' ? id : '',
    fontSize: 18,
    color: '#172033',
    fill:
      type === 'shape'
        ? '#dce6fb'
        : type === 'table'
          ? '#ffffff'
          : 'transparent',
    bold: false,
    align: 'left',
    borderColor: '#657087',
    borderWidth: type === 'line' ? 2 : 1,
    groupIds,
  };
}
