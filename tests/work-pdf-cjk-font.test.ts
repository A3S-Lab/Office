import { afterEach, describe, expect, test } from '@rstest/core';
import { readFileSync } from 'node:fs';
import { jsPDF } from 'jspdf';
import {
  clearWorkPdfCjkFont,
  registerWorkPdfCjkFont,
  workPdfCjkFontRegistered,
} from '../src/internal/features/work/work-pdf-cjk-font';
import { workPdfTextRunsFromClientRects } from '../src/internal/features/work/work-pdf-text-layer';
import { appendWorkPdfVectorTextLayer } from '../src/internal/features/work/work-pdf-vector-text';

afterEach(() => {
  clearWorkPdfCjkFont();
});

describe('Work PDF CJK font registration', () => {
  test('rejects OpenType CFF payloads and admits TrueType bytes', () => {
    expect(() =>
      registerWorkPdfCjkFont({
        data: new Uint8Array([0x4f, 0x54, 0x54, 0x4f, 0x00, 0x00]),
      }),
    ).toThrow(/TrueType/);
    expect(workPdfCjkFontRegistered()).toBe(false);

    registerWorkPdfCjkFont({
      data: new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x00, 0x0c]),
      family: 'TestCjk',
      vfsName: 'TestCjk.ttf',
    });
    expect(workPdfCjkFontRegistered()).toBe(true);
  });

  test('admits CJK searchable runs only after a host TTF is registered', () => {
    expect(
      workPdfTextRunsFromClientRects(
        '中文标题',
        [{ left: 10, top: 10, width: 40, height: 16 }],
        { left: 0, top: 0, width: 200, height: 200 },
        14,
      ),
    ).toEqual([]);

    registerWorkPdfCjkFont({
      data: new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x00, 0x0c]),
      family: 'TestCjk',
      vfsName: 'TestCjk.ttf',
    });
    expect(
      workPdfTextRunsFromClientRects(
        '中文标题',
        [{ left: 10, top: 10, width: 40, height: 16 }],
        { left: 0, top: 0, width: 200, height: 200 },
        14,
      ),
    ).toEqual([
      {
        fontSize: 14,
        height: 16,
        text: '中文标题',
        width: 40,
        x: 10,
        y: 10,
      },
    ]);
  });

  test('paints CJK vector text when a real host TrueType face is available', () => {
    const fontPath = 'C:\\Windows\\Fonts\\simhei.ttf';
    let bytes: Buffer;
    try {
      bytes = readFileSync(fontPath);
    } catch {
      // Host without SimHei; registration unit coverage above still applies.
      return;
    }
    registerWorkPdfCjkFont({
      data: bytes,
      family: 'SimHei',
      vfsName: 'simhei.ttf',
    });
    const pdf = new jsPDF({ unit: 'pt', format: 'a4' });
    expect(() =>
      appendWorkPdfVectorTextLayer(
        pdf,
        [
          {
            color: '#000000',
            fontSize: 16,
            fontStyle: 'normal',
            height: 18,
            text: '季度计划',
            width: 80,
            x: 40,
            y: 60,
          },
        ],
        { width: 595.28, height: 841.89 },
        { pageWidthPoints: 595.28, pageHeightPoints: 841.89 },
      ),
    ).not.toThrow();
    const output = pdf.output('arraybuffer');
    expect(output.byteLength).toBeGreaterThan(1_000);
  });
});
