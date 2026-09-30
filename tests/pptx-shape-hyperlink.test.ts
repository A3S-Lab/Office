import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { createPptxBlob } from '../src/internal/features/work/work-pptx-export';
import { importPptxPresentation } from '../src/internal/features/work/work-pptx-import';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';
import type {
  WorkArtifact,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('keeps an https hyperlink on a shape through PPTX save and reopen', async () => {
  const blob = await createPptxBlob(linkedShapeArtifact(), PptxGenJS);
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const slideXml = await archive.text('ppt/slides/slide1.xml');
  const rels = await archive.text('ppt/slides/_rels/slide1.xml.rels');

  expect(slideXml).toContain('hlinkClick');
  expect(rels).toContain('https://example.com');
  expect(rels).toContain('https://example.com/labeled');
  expect(slideXml).not.toContain('rIdundefined');

  const imported = await importPptxPresentation(
    await happyDomImportableFile(blob, 'linked-shape.pptx'),
  );
  const shapes = (imported.content.slides[0]?.elements ?? []).filter(
    (element) => element.type === 'shape',
  );
  expect(shapes.map((element) => element.href).sort()).toEqual([
    'https://example.com',
    'https://example.com/labeled',
  ]);
});

function linkedShapeArtifact(): WorkArtifact {
  const shape: WorkSlideElement = {
    id: 'rectangle',
    type: 'shape',
    x: 20,
    y: 30,
    width: 40,
    height: 20,
    text: '',
    fontSize: 18,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
    shapeType: 'rect',
    borderColor: '#657087',
    borderWidth: 1,
    href: 'https://example.com',
  };
  const labeled: WorkSlideElement = {
    ...shape,
    id: 'labeled',
    y: 55,
    text: 'Label',
    href: 'https://example.com/labeled',
  };
  return {
    id: 'linked-shape',
    kind: 'presentation',
    title: 'Linked shape',
    favorite: false,
    createdAt: 0,
    updatedAt: 0,
    lastOpenedAt: 0,
    revision: 0,
    content: {
      type: 'presentation',
      slides: [
        {
          id: 'slide-1',
          name: 'Slide',
          background: '#ffffff',
          elements: [shape, labeled],
        },
      ],
    },
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
