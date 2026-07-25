import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importPptxPresentation } from '../src/internal/features/work/work-pptx-import';
import {
  attribute,
  descendants,
  OoxmlPackage,
} from '../src/internal/features/work/work-ooxml-package';

test('retains nested PPTX group selection paths without a blanket compatibility warning', async () => {
  const file = await groupedPresentationFile();
  const archive = await OoxmlPackage.load(await file.arrayBuffer());
  const presentation = await archive.xml('ppt/presentation.xml');
  const relationships = await archive.relationships('ppt/presentation.xml');
  expect(descendants(presentation, 'sldId')).toHaveLength(1);
  expect(attribute(descendants(presentation, 'sldId')[0], 'id')).toBe('rId1');
  expect(relationships.get('rId1')).toMatchObject({
    target: 'ppt/slides/slide1.xml',
  });
  const result = await importPptxPresentation(file);
  const elements = result.content.slides[0]?.elements ?? [];

  expect(elements).toHaveLength(2);
  expect(elements[0].groupIds).toHaveLength(1);
  expect(elements[1].groupIds).toHaveLength(2);
  expect(elements[1].groupIds?.[0]).toBe(elements[0].groupIds?.[0]);
  expect(elements[1].groupIds?.[1]).not.toBe(elements[0].groupIds?.[0]);
  expect(result.compatibility.issues).not.toContainEqual(
    expect.objectContaining({ code: 'pptx.group' }),
  );
});

test('scales grouped typography and borders by the smaller cumulative axis', async () => {
  const file = await presentationFile(
    groupShape(
      transform(0, 0, 4_000_000, 2_000_000, {
        childWidth: 2_000_000,
        childHeight: 1_000_000,
      }),
      [
        groupShape(
          transform(0, 0, 1_500_000, 750_000, {
            childWidth: 2_000_000,
            childHeight: 1_000_000,
          }),
          [styledShape('Scaled member', 0, 0, 1_000_000, 500_000)],
        ),
      ],
    ),
    'scaled-group.pptx',
  );

  const result = await importPptxPresentation(file);
  const element = result.content.slides[0]?.elements[0];

  expect(element?.fontSize).toBe(30);
  expect(element?.borderWidth).toBe(3);
  expect(element?.textRuns?.[0]?.fontSize).toBe(30);
  expect(element?.textRuns?.[1]?.fontSize).toBeUndefined();
  expect(element?.groupIds).toHaveLength(2);
});

test('reports only unsupported rotated or reflected group transforms', async () => {
  const file = await presentationFile(
    groupShape(
      transform(0, 0, 2_000_000, 1_000_000, {
        flipH: true,
        rotation: 60_000,
      }),
      [shape('Transformed member', 0, 0, 1_000_000, 500_000)],
    ),
    'transformed-group.pptx',
  );

  const result = await importPptxPresentation(file);

  expect(result.compatibility.issues).toContainEqual(
    expect.objectContaining({ code: 'pptx.group.transform' }),
  );
  expect(result.compatibility.issues).not.toContainEqual(
    expect.objectContaining({ code: 'pptx.group' }),
  );
});

async function groupedPresentationFile(): Promise<File> {
  return presentationFile(
    groupShape(transform(0, 0, 12_192_000, 6_858_000), [
      shape('Outer member', 500_000, 500_000, 1_800_000, 900_000),
      groupShape(transform(3_000_000, 1_000_000, 5_000_000, 3_000_000), [
        shape('Inner member', 0, 0, 2_000_000, 1_000_000),
      ]),
    ]),
    'grouped.pptx',
  );
}

async function presentationFile(scene: string, name: string): Promise<File> {
  const zip = new JSZip();
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldIdLst><p:sldId id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/>
      </p:presentation>`,
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship
          Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
          Target="slides/slide1.xml"/>
      </Relationships>`,
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
          <p:spTree>
            ${scene}
          </p:spTree>
        </p:cSld>
      </p:sld>`,
  );
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return new File([bytes], name, {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}

function groupShape(groupTransform: string, children: string[]): string {
  return `<p:grpSp>
    <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr>${groupTransform}</p:grpSpPr>
    ${children.join('')}
  </p:grpSp>`;
}

function shape(
  name: string,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="DCE6FB"/></a:solidFill>
    </p:spPr>
  </p:sp>`;
}

function styledShape(
  name: string,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="DCE6FB"/></a:solidFill>
      <a:ln w="25400"><a:solidFill><a:srgbClr val="657087"/></a:solidFill></a:ln>
    </p:spPr>
    <p:txBody>
      <a:bodyPr/><a:lstStyle/>
      <a:p>
        <a:r><a:rPr sz="2000"/><a:t>Scaled text</a:t></a:r>
        <a:r><a:rPr b="1"/><a:t> with inherited size</a:t></a:r>
      </a:p>
    </p:txBody>
  </p:sp>`;
}

function transform(
  x: number,
  y: number,
  width: number,
  height: number,
  options: {
    childHeight?: number;
    childWidth?: number;
    childX?: number;
    childY?: number;
    flipH?: boolean;
    flipV?: boolean;
    rotation?: number;
  } = {},
): string {
  const attributes = [
    options.rotation ? `rot="${options.rotation}"` : '',
    options.flipH ? 'flipH="1"' : '',
    options.flipV ? 'flipV="1"' : '',
  ]
    .filter(Boolean)
    .join(' ');
  return `<a:xfrm${attributes ? ` ${attributes}` : ''}>
    <a:off x="${x}" y="${y}"/>
    <a:ext cx="${width}" cy="${height}"/>
    <a:chOff x="${options.childX ?? 0}" y="${options.childY ?? 0}"/>
    <a:chExt cx="${options.childWidth ?? width}" cy="${options.childHeight ?? height}"/>
  </a:xfrm>`;
}
