import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { createPptxBlob } from '../src/internal/features/work/work-pptx-export';
import { importPptxPresentation } from '../src/internal/features/work/work-pptx-import';
import {
  attribute,
  descendants,
  directChild,
  OoxmlPackage,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { readPptxAnimations } from '../src/internal/features/work/work-pptx-animation';
import type {
  WorkArtifact,
  WorkPresentationContent,
  WorkSlideAnimation,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('round-trips supported entrance and exit animations through native PPTX timing trees', async () => {
  const blob = await createPptxBlob(animatedArtifact(), PptxGenJS);
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const slide = await archive.xml('ppt/slides/slide1.xml');
  const timing = directChild(slide.documentElement, 'timing');
  if (!timing) throw new Error('Generated PPTX is missing animation timing.');

  const timeNodeIds = descendants(timing, 'cTn').map((node) =>
    attribute(node, 'id'),
  );
  const shapeIds = new Set(
    descendants(slide, 'cNvPr').map((node) => attribute(node, 'id')),
  );
  const targetIds = Array.from(
    new Set(
      descendants(timing, 'spTgt').map((node) => attribute(node, 'spid')),
    ),
  );
  expect(new Set(timeNodeIds).size).toBe(timeNodeIds.length);
  const effectNodes = descendants(timing, 'cTn').filter((node) =>
    ['entr', 'exit'].includes(attribute(node, 'presetClass') ?? ''),
  );
  expect(
    effectNodes.map((node) => [
      attribute(node, 'presetClass'),
      attribute(node, 'nodeType'),
    ]),
  ).toEqual([
    ['entr', 'clickEffect'],
    ['entr', 'withEffect'],
    ['entr', 'afterEffect'],
    ['entr', 'clickEffect'],
    ['exit', 'clickEffect'],
    ['exit', 'withEffect'],
    ['exit', 'afterEffect'],
    ['exit', 'afterEffect'],
  ]);
  expect(
    descendants(timing, 'animEffect')
      .filter((node) => attribute(node, 'transition') === 'out')
      .map((node) => attribute(node, 'filter')),
  ).toEqual(['fade', 'slide(fromBottom)']);
  expect(targetIds).toHaveLength(4);
  expect(targetIds.every((id) => shapeIds.has(id))).toBe(true);
  expect(
    descendants(timing, 'bldP').map((node) => attribute(node, 'spid')),
  ).toEqual(targetIds);
  expect(await archive.text('ppt/slides/slide1.xml')).not.toContain(
    '__A3S_OFFICE_EXPORT__',
  );

  const imported = await importPptxPresentation(
    await happyDomImportableFile(blob, 'animated-export.pptx'),
  );
  const importedElements = imported.content.slides[0]?.elements ?? [];
  expect(importedElements).toHaveLength(5);
  const importedMessageShapes = importedElements.filter(
    (element) => element.text === 'One browser object',
  );
  expect(importedMessageShapes).toHaveLength(1);
  expect(importedMessageShapes[0]).toMatchObject({
    type: 'shape',
    text: 'One browser object',
  });
  expect(animationProjection(imported.content)).toEqual([
    ['appear', 'on-click', 300, 0, undefined],
    ['fade', 'with-previous', 600, 100, undefined],
    ['fly-in', 'after-previous', 700, 150, 'right'],
    ['zoom', 'on-click', 800, 200, undefined],
    ['disappear', 'on-click', 200, 0, undefined],
    ['fade-out', 'with-previous', 500, 100, undefined],
    ['fly-out', 'after-previous', 650, 50, 'down'],
    ['zoom-out', 'after-previous', 400, 0, undefined],
  ]);
  expect(imported.compatibility.issues).toContainEqual(
    expect.objectContaining({ code: 'pptx.animation', severity: 'info' }),
  );
  expect(imported.compatibility.issues).not.toContainEqual(
    expect.objectContaining({ code: 'pptx.animation.target' }),
  );

  const reopenedBlob = await createPptxBlob(
    artifactWithContent(imported.content),
    PptxGenJS,
  );
  const reopened = await importPptxPresentation(
    await happyDomImportableFile(reopenedBlob, 'animated-reopened.pptx'),
  );
  expect(animationProjection(reopened.content)).toEqual(
    animationProjection(imported.content),
  );
});

test('imports the supported entrance and exit subset and diagnoses duplicate, unsupported, and missing targets', async () => {
  const imported = await importPptxPresentation(
    await malformedAnimationPresentation(),
  );
  const animations = imported.content.slides[0]?.animations ?? [];

  expect(animations).toEqual([
    expect.objectContaining({
      effect: 'fade',
      trigger: 'on-click',
      durationMs: 700,
      delayMs: 50,
    }),
    expect.objectContaining({
      effect: 'fade-out',
      trigger: 'after-previous',
      durationMs: 400,
      delayMs: 0,
    }),
  ]);
  expect(new Set(animations.map(({ elementId }) => elementId))).toEqual(
    new Set([imported.content.slides[0]?.elements[0]?.id]),
  );
  for (const code of [
    'pptx.animation.duplicate-target',
    'pptx.animation.class',
    'pptx.animation.target',
  ]) {
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code, severity: 'warning' }),
    );
  }
});

test('ignores namespace-spoofed animation timing instead of trusting local names', () => {
  const document = parseXml(
    '<p:sld xmlns:p="urn:not-presentation"><p:timing><p:cTn nodeType="clickEffect" presetClass="entr"/></p:timing></p:sld>',
  );
  const result = readPptxAnimations(document, new Map([['2', 'element']]));

  expect(result.animations).toEqual([]);
  expect(result.diagnostics).toContainEqual(
    expect.objectContaining({ code: 'pptx.animation.namespace' }),
  );
});

test('rejects spoofed behavior nodes, unsupported filters, and mismatched transitions', () => {
  const document = parseXml(
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:evil="urn:not-presentation">
      <p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>
        <p:par><p:cTn id="2" nodeType="clickEffect" presetClass="entr" presetID="10">
          <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:set>
          <evil:animEffect filter="fade"/></p:childTnLst>
        </p:cTn></p:par>
        <p:par><p:cTn id="3" nodeType="clickEffect" presetClass="entr" presetID="999">
          <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:set>
          <p:animEffect filter="slide(fromTopLeft)"/></p:childTnLst>
        </p:cTn></p:par>
        <p:par><p:cTn id="4" nodeType="clickEffect" presetClass="exit" presetID="10">
          <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:set>
          <p:animEffect transition="in" filter="fade"/></p:childTnLst>
        </p:cTn></p:par>
      </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>
    </p:sld>`,
  );
  const result = readPptxAnimations(
    document,
    new Map([
      ['2', 'first'],
      ['3', 'second'],
      ['4', 'third'],
    ]),
  );

  expect(result.animations).toEqual([]);
  expect(
    result.diagnostics.filter(({ code }) => code === 'pptx.animation.effect'),
  ).toHaveLength(3);
});

function animationProjection(
  content: WorkPresentationContent,
): Array<
  [
    WorkSlideAnimation['effect'],
    WorkSlideAnimation['trigger'],
    number,
    number,
    WorkSlideAnimation['direction'],
  ]
> {
  return (content.slides[0]?.animations ?? []).map((animation) => [
    animation.effect,
    animation.trigger,
    animation.durationMs,
    animation.delayMs,
    animation.direction,
  ]);
}

function animatedArtifact(): WorkArtifact {
  const elements: WorkSlideElement[] = [
    element('title', 'text', 'Launch plan', 8, 8, 38, 12),
    {
      ...element('message', 'shape', 'One browser object', 8, 28, 38, 18),
      fill: '#dce6fb',
      groupIds: ['message-group'],
    },
    {
      ...element('picture', 'image', '', 54, 8, 18, 20),
      image: {
        dataUrl:
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=',
        contentType: 'image/png',
        name: 'pixel.png',
      },
    },
    {
      ...element('table', 'table', '', 54, 38, 34, 24),
      fill: '#ffffff',
      table: {
        headerRows: 1,
        rows: [
          ['Milestone', 'Owner'],
          ['Launch', 'Team'],
        ],
      },
    },
    {
      ...element('message-peer', 'shape', '', 8, 50, 20, 12),
      fill: '#eef1f6',
      groupIds: ['message-group'],
    },
  ];
  return {
    id: 'animated-artifact',
    kind: 'presentation',
    title: 'Entrance animation round trip',
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
          id: 'animated-slide',
          name: 'Animated slide',
          background: '#ffffff',
          elements,
          animations: [
            animation('appear-title', 'title', 'appear', 'on-click', 300, 0),
            animation(
              'fade-message',
              'message',
              'fade',
              'with-previous',
              600,
              100,
            ),
            {
              ...animation(
                'fly-picture',
                'picture',
                'fly-in',
                'after-previous',
                700,
                150,
              ),
              direction: 'right',
            },
            animation('zoom-table', 'table', 'zoom', 'on-click', 800, 200),
            animation(
              'disappear-title',
              'title',
              'disappear',
              'on-click',
              200,
              0,
            ),
            animation(
              'fade-out-message',
              'message',
              'fade-out',
              'with-previous',
              500,
              100,
            ),
            {
              ...animation(
                'fly-out-picture',
                'picture',
                'fly-out',
                'after-previous',
                650,
                50,
              ),
              direction: 'down',
            },
            animation(
              'zoom-out-table',
              'table',
              'zoom-out',
              'after-previous',
              400,
              0,
            ),
          ],
        },
      ],
    },
  };
}

function animation(
  id: string,
  elementId: string,
  effect: WorkSlideAnimation['effect'],
  trigger: WorkSlideAnimation['trigger'],
  durationMs: number,
  delayMs: number,
): WorkSlideAnimation {
  return { id, elementId, effect, trigger, durationMs, delayMs };
}

function element(
  id: string,
  type: WorkSlideElement['type'],
  text: string,
  x: number,
  y: number,
  width: number,
  height: number,
): WorkSlideElement {
  return {
    id,
    type,
    x,
    y,
    width,
    height,
    text,
    fontSize: 18,
    color: '#172033',
    fill: 'transparent',
    bold: false,
    align: 'left',
    borderColor: '#657087',
    borderWidth: type === 'shape' ? 1 : undefined,
  };
}

function artifactWithContent(content: WorkPresentationContent): WorkArtifact {
  return {
    id: 'reopened-artifact',
    kind: 'presentation',
    title: 'Reopened animations',
    favorite: false,
    createdAt: 0,
    updatedAt: 0,
    lastOpenedAt: 0,
    revision: 0,
    content,
  };
}

async function happyDomImportableFile(blob: Blob, name: string): Promise<File> {
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
  return new File([await archive.generateAsync({ type: 'uint8array' })], name, {
    type: blob.type,
  });
}

async function malformedAnimationPresentation(): Promise<File> {
  const archive = new JSZip();
  archive.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldIdLst><p:sldId id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/>
      </p:presentation>`,
  );
  archive.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
      </Relationships>`,
  );
  archive.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree>
          ${shapeXml(2, 'Alpha', 500000)}
          ${shapeXml(3, 'Beta', 3000000)}
        </p:spTree></p:cSld>
        <p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>
          ${fadeEffectXml(10, 2, 'entr', 'clickEffect', 700, 50)}
          ${fadeEffectXml(20, 2, 'exit', 'afterEffect', 400, 0)}
          ${fadeEffectXml(30, 2, 'exit', 'clickEffect', 400, 0)}
          ${fadeEffectXml(40, 3, 'emph', 'afterEffect', 400, 0)}
          ${fadeEffectXml(50, 999, 'entr', 'afterEffect', 400, 0)}
        </p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>
      </p:sld>`,
  );
  return new File(
    [await archive.generateAsync({ type: 'uint8array' })],
    'malformed-animations.pptx',
    {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    },
  );
}

function shapeXml(id: number, name: string, x: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="500000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${name}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function fadeEffectXml(
  id: number,
  spid: number,
  presetClass: 'emph' | 'entr' | 'exit',
  nodeType: 'afterEffect' | 'clickEffect',
  durationMs: number,
  delayMs: number,
): string {
  const transition = presetClass === 'exit' ? 'out' : 'in';
  return `<p:par><p:cTn id="${id}" presetID="10" presetClass="${presetClass}" presetSubtype="0" nodeType="${nodeType}" dur="${durationMs}"><p:stCondLst><p:cond delay="${delayMs}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set><p:animEffect transition="${transition}" filter="fade"><p:cBhvr><p:cTn id="${id + 2}" dur="${durationMs}"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`;
}
