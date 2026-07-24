import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importPptxPresentation } from '../src/internal/features/work/work-pptx-import';
import {
  attribute,
  descendants,
  OoxmlPackage,
} from '../src/internal/features/work/work-ooxml-package';

test('retains nested PPTX group selection paths after flattening geometry', async () => {
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
  expect(result.compatibility.issues).toContainEqual(
    expect.objectContaining({ code: 'pptx.group' }),
  );
});

async function groupedPresentationFile(): Promise<File> {
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
            ${groupShape(transform(0, 0, 12192000, 6858000), [
              shape('Outer member', 500000, 500000, 1800000, 900000),
              groupShape(transform(3000000, 1000000, 5000000, 3000000), [
                shape('Inner member', 0, 0, 2000000, 1000000),
              ]),
            ])}
          </p:spTree>
        </p:cSld>
      </p:sld>`,
  );
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return new File([bytes], 'grouped.pptx', {
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

function transform(
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  return `<a:xfrm>
    <a:off x="${x}" y="${y}"/>
    <a:ext cx="${width}" cy="${height}"/>
    <a:chOff x="0" y="0"/>
    <a:chExt cx="${width}" cy="${height}"/>
  </a:xfrm>`;
}
