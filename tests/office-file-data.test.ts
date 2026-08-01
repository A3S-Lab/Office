import { expect, test } from '@rstest/core';
import { readOfficeFileAsDataUrl } from '../src/internal/features/work/editors/office-file-data';

test('materializes transient files before reading a durable data URL', async () => {
  const source = new File(['image'], 'diagram.png', { type: 'image/png' });
  let arrayBufferReads = 0;
  Object.defineProperty(source, 'arrayBuffer', {
    configurable: true,
    value: async () => {
      arrayBufferReads += 1;
      return new Uint8Array([1, 2, 3]).buffer;
    },
  });
  let readerSource: File | null = null;
  const originalFileReader = globalThis.FileReader;

  class InspectingFileReader extends EventTarget {
    error: DOMException | null = null;
    result: string | ArrayBuffer | null = null;

    readAsDataURL(blob: File) {
      readerSource = blob;
      this.result = 'data:image/png;base64,AQID';
      this.dispatchEvent(new Event('load'));
    }
  }

  Object.defineProperty(globalThis, 'FileReader', {
    configurable: true,
    value: InspectingFileReader,
  });
  try {
    await expect(readOfficeFileAsDataUrl(source)).resolves.toBe(
      'data:image/png;base64,AQID',
    );
  } finally {
    Object.defineProperty(globalThis, 'FileReader', {
      configurable: true,
      value: originalFileReader,
    });
  }

  expect(arrayBufferReads).toBe(1);
  expect(readerSource).toBeInstanceOf(File);
  expect(readerSource).not.toBe(source);
  expect(readerSource?.type).toBe('image/png');
  expect(readerSource?.name).toBe('diagram.png');
});
