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

const TINY_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=';

test('keeps author alt text on a picture through PPTX save and reopen', async () => {
  const blob = await createPptxBlob(
    pictureArtifact({ altText: '图', name: 'a3s-pic-large.png' }),
    PptxGenJS,
  );
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const slideXml = await archive.text('ppt/slides/slide1.xml');

  expect(slideXml).toContain('descr="图"');
  expect(slideXml).not.toContain('preencoded.png');

  const imported = await importPptxPresentation(
    await happyDomImportableFile(blob, 'alt-text.pptx'),
  );
  expect(imported.content.slides[0]?.elements[0]?.altText).toBe('图');
});

test('keeps a picture file name when the author has not set alt text', async () => {
  const blob = await createPptxBlob(
    pictureArtifact({ name: 'a3s-pic-large.png' }),
    PptxGenJS,
  );
  const archive = await OoxmlPackage.load(await blob.arrayBuffer());
  const slideXml = await archive.text('ppt/slides/slide1.xml');

  expect(slideXml).toContain('descr="a3s-pic-large.png"');
  expect(slideXml).not.toContain('preencoded.png');

  const imported = await importPptxPresentation(
    await happyDomImportableFile(blob, 'picture-name.pptx'),
  );
  expect(imported.content.slides[0]?.elements[0]?.altText).toBe(
    'a3s-pic-large.png',
  );
});

test('does not adopt a synthetic preencoded description over the picture name', async () => {
  const blob = await createPptxBlob(
    pictureArtifact({ altText: '图', name: 'a3s-pic-large.png' }),
    PptxGenJS,
  );
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const path = 'ppt/slides/slide1.xml';
  const xml = await archive.file(path)?.async('text');
  if (!xml) throw new Error('Generated PPTX is missing slide1.xml.');
  const rewritten = xml.replace(/<p:cNvPr\b[^>]*>/g, (tag) =>
    tag.includes('descr=')
      ? tag
          .replace(/name="[^"]*"/, 'name="a3s-pic-large.png"')
          .replace(/descr="[^"]*"/, 'descr="preencoded.png"')
      : tag,
  );
  if (!rewritten.includes('descr="preencoded.png"')) {
    throw new Error(`Picture description was not rewritten: ${xml}`);
  }
  archive.file(path, rewritten);
  const patched = new Blob(
    [await archive.generateAsync({ type: 'uint8array' })],
    { type: blob.type },
  );
  const imported = await importPptxPresentation(
    await happyDomImportableFile(patched, 'synthetic-descr.pptx'),
  );
  expect(imported.content.slides[0]?.elements[0]?.altText).toBe(
    'a3s-pic-large.png',
  );
  expect(imported.content.slides[0]?.elements[0]?.altText).not.toBe(
    'preencoded.png',
  );
});

function pictureArtifact(options: {
  altText?: string;
  name: string;
}): WorkArtifact {
  const image: WorkSlideElement = {
    id: 'picture',
    type: 'image',
    x: 10,
    y: 10,
    width: 40,
    height: 30,
    text: '',
    fontSize: 12,
    color: '#172033',
    fill: 'transparent',
    bold: false,
    align: 'center',
    altText: options.altText,
    image: {
      dataUrl: TINY_PNG,
      contentType: 'image/png',
      name: options.name,
    },
  };
  return {
    id: 'picture-deck',
    kind: 'presentation',
    title: 'Picture alt text',
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
          elements: [image],
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
