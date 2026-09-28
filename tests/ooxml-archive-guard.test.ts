import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import {
  assertOoxmlArchiveWithinLimits,
  DEFAULT_OOXML_ARCHIVE_LIMITS,
  OoxmlArchiveLimitError,
} from '../src/internal/features/work/work-ooxml-archive-guard';

async function zip(
  files: Record<string, string | Uint8Array>,
  compression: 'DEFLATE' | 'STORE' = 'DEFLATE',
): Promise<Uint8Array> {
  const archive = new JSZip();
  for (const [name, content] of Object.entries(files))
    archive.file(name, content);
  return archive.generateAsync({ type: 'uint8array', compression });
}

function limitCode(run: () => void): string | null {
  try {
    run();
    return null;
  } catch (error) {
    return error instanceof OoxmlArchiveLimitError ? error.code : 'unexpected';
  }
}

/** Rewrites every central-directory uncompressed size to simulate a lying header. */
function withDeclaredUncompressedSize(bytes: Uint8Array, size: number) {
  const copy = bytes.slice();
  const view = new DataView(copy.buffer);
  for (let offset = 0; offset + 46 <= copy.length; offset += 1) {
    if (view.getUint32(offset, true) === 0x02014b50) {
      view.setUint32(offset + 24, size, true);
    }
  }
  return copy;
}

describe('OOXML archive guard', () => {
  test('admits an ordinary package', async () => {
    const bytes = await zip({
      '[Content_Types].xml': '<Types/>',
      'word/document.xml': '<w:document/>',
    });
    expect(limitCode(() => assertOoxmlArchiveWithinLimits(bytes))).toBeNull();
  });

  test('rejects a deflate bomb by per-part compression ratio', async () => {
    const bomb = await zip({ 'word/document.xml': new Uint8Array(4 << 20) });
    expect(bomb.length).toBeLessThan(64 * 1024);
    expect(limitCode(() => assertOoxmlArchiveWithinLimits(bomb))).toBe(
      'package.archive.compression-ratio',
    );
  });

  test('rejects too many entries', async () => {
    const bytes = await zip({ a: 'a', b: 'b', c: 'c' });
    expect(
      limitCode(() =>
        assertOoxmlArchiveWithinLimits(bytes, {
          ...DEFAULT_OOXML_ARCHIVE_LIMITS,
          maxEntries: 2,
        }),
      ),
    ).toBe('package.archive.entries');
  });

  test('rejects an oversized declared part and total', async () => {
    const bytes = withDeclaredUncompressedSize(
      await zip({ a: 'a', b: 'b' }, 'STORE'),
      200 * 1024 * 1024,
    );
    expect(limitCode(() => assertOoxmlArchiveWithinLimits(bytes))).toBe(
      'package.archive.part-bytes',
    );
    expect(
      limitCode(() =>
        assertOoxmlArchiveWithinLimits(bytes, {
          ...DEFAULT_OOXML_ARCHIVE_LIMITS,
          maxPartBytes: 256 * 1024 * 1024,
          maxUncompressedBytes: 256 * 1024 * 1024,
          maxCompressionRatio: Number.MAX_SAFE_INTEGER,
        }),
      ),
    ).toBe('package.archive.uncompressed-bytes');
  });

  test('rejects an archive larger than the archive budget', async () => {
    const bytes = await zip({ a: 'a' });
    expect(
      limitCode(() =>
        assertOoxmlArchiveWithinLimits(bytes, {
          ...DEFAULT_OOXML_ARCHIVE_LIMITS,
          maxArchiveBytes: bytes.length - 1,
        }),
      ),
    ).toBe('package.archive.archive-bytes');
  });

  test('rejects a truncated central directory', async () => {
    const bytes = await zip({ 'word/document.xml': '<w:document/>' });
    expect(
      limitCode(() => assertOoxmlArchiveWithinLimits(bytes.slice(0, -30))),
    ).toBe('package.archive.malformed');
  });

  test('importOfficeFile rejects a bomb before any format parser runs', async () => {
    const bomb = await zip({
      '[Content_Types].xml': '<Types/>',
      'word/document.xml': new Uint8Array(4 << 20),
    });
    const codes: Record<string, string | null> = {};
    for (const name of ['bomb.docx', 'bomb.xlsx', 'bomb.pptx']) {
      try {
        await importOfficeFile(new File([bomb], name));
        codes[name] = null;
      } catch (error) {
        codes[name] =
          error instanceof OoxmlArchiveLimitError ? error.code : String(error);
      }
    }
    expect(codes).toEqual({
      'bomb.docx': 'package.archive.compression-ratio',
      'bomb.xlsx': 'package.archive.compression-ratio',
      'bomb.pptx': 'package.archive.compression-ratio',
    });
  });
});
