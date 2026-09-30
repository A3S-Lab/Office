import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';

const PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=';

test('exports a data-url PNG as a word/media drawing when data URLs cannot be fetched', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    if (url.startsWith('data:')) {
      throw new TypeError('data URLs are not fetchable');
    }
    return originalFetch(input, init);
  }) as typeof fetch;

  try {
    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: `<p><img src="${PNG_DATA_URL}" alt="a3s-pic.png" width="48" height="32"></p>`,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await archive.file('word/document.xml')?.async('text');
    const media = Object.keys(archive.files).filter((path) =>
      path.startsWith('word/media/'),
    );

    expect(documentXml).toBeTruthy();
    expect(documentXml).not.toContain('[a3s-pic.png]');
    expect(documentXml).toContain('a:blip');
    expect(media.length).toBeGreaterThan(0);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
