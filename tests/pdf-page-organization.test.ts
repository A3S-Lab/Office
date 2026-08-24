import { describe, expect, test } from '@rstest/core';
import { PDFDocument, PDFName, PDFNumber, PDFString } from 'pdf-lib';
import {
  PdfPageOrganizationError,
  applyPdfPageOrganizationJob,
  reorderedPdfPageIndexes,
  type PdfPageOrganizationJob,
} from '../src/internal/features/work/editors/pdf-page-organization';

describe('PDF page organization', () => {
  test('plans a stable multi-page move against an original insertion gap', () => {
    expect(reorderedPdfPageIndexes(6, [1, 3], 6)).toEqual([0, 2, 4, 5, 1, 3]);
    expect(reorderedPdfPageIndexes(6, [1, 3], 0)).toEqual([1, 3, 0, 2, 4, 5]);
    expect(reorderedPdfPageIndexes(6, [1, 3], 3)).toEqual([0, 2, 1, 3, 4, 5]);
  });

  test('inserts, rotates, deletes, and reorders pages with exact structural evidence', async () => {
    const source = await createPdf([301, 302, 303], 'Primary');

    const inserted = await mutate(source, {
      kind: 'insert-blank',
      index: 1,
      size: { height: 500, width: 400 },
    });
    expect(await pageEvidence(inserted.source)).toEqual([
      [301, 600, 0],
      [400, 500, 0],
      [302, 600, 0],
      [303, 600, 0],
    ]);

    const rotated = await mutate(inserted.source, {
      degrees: 90,
      kind: 'rotate',
      pageIndexes: [0, 2],
    });
    expect(await pageEvidence(rotated.source)).toEqual([
      [301, 600, 90],
      [400, 500, 0],
      [302, 600, 90],
      [303, 600, 0],
    ]);

    const reordered = await mutate(rotated.source, {
      kind: 'reorder',
      pageOrder: [3, 0, 2, 1],
    });
    expect(await pageEvidence(reordered.source)).toEqual([
      [303, 600, 0],
      [301, 600, 90],
      [302, 600, 90],
      [400, 500, 0],
    ]);

    const deleted = await mutate(reordered.source, {
      kind: 'delete',
      pageIndexes: [1, 3],
    });
    expect(deleted.pageCount).toBe(2);
    expect(await pageEvidence(deleted.source)).toEqual([
      [303, 600, 0],
      [302, 600, 90],
    ]);
  });

  test('merges an imported PDF at an exact page gap and preserves primary metadata', async () => {
    const primary = await createPdf([301, 302], 'Primary title');
    const imported = await createPdf([401, 402], 'Imported title');
    const result = await applyPdfPageOrganizationJob({
      importSource: imported,
      kind: 'mutate',
      mutation: { index: 1, kind: 'merge' },
      source: primary,
    });
    expect(result.kind).toBe('mutated');
    if (result.kind !== 'mutated')
      throw new Error('Expected a mutation result.');
    expect(await pageEvidence(result.source)).toEqual([
      [301, 600, 0],
      [401, 600, 0],
      [402, 600, 0],
      [302, 600, 0],
    ]);
    const reopened = await PDFDocument.load(result.source);
    expect(reopened.getTitle()).toBe('Primary title');
  });

  test('extracts an ordered selection and splits at multiple exact boundaries', async () => {
    const source = await createPdf([301, 302, 303, 304, 305], 'Exports');
    const extracted = await applyPdfPageOrganizationJob({
      kind: 'export',
      operation: { kind: 'extract', pageIndexes: [4, 1, 3] },
      source,
    });
    expect(extracted.kind).toBe('exported');
    if (extracted.kind !== 'exported')
      throw new Error('Expected export files.');
    expect(extracted.files).toHaveLength(1);
    expect(await pageWidths(extracted.files[0]?.source)).toEqual([
      305, 302, 304,
    ]);

    const split = await applyPdfPageOrganizationJob({
      kind: 'export',
      operation: { kind: 'split', splitAfterPageIndexes: [0, 3] },
      source,
    });
    expect(split.kind).toBe('exported');
    if (split.kind !== 'exported') throw new Error('Expected split files.');
    expect(split.files.map((file) => file.pageCount)).toEqual([1, 3, 1]);
    expect(
      await Promise.all(split.files.map((file) => pageWidths(file.source))),
    ).toEqual([[301], [302, 303, 304], [305]]);
  });

  test('rejects destructive or ambiguous input instead of normalizing silently', async () => {
    const source = await createPdf([301, 302], 'Bounded');
    await expect(
      mutate(source, { kind: 'delete', pageIndexes: [0, 1] }),
    ).rejects.toMatchObject({ code: 'pdf.pages.empty-document' });
    await expect(
      mutate(source, { kind: 'rotate', pageIndexes: [0, 0], degrees: 90 }),
    ).rejects.toMatchObject({ code: 'pdf.pages.duplicate-index' });
    await expect(
      mutate(source, { kind: 'reorder', pageOrder: [0, 0] }),
    ).rejects.toMatchObject({ code: 'pdf.pages.invalid-order' });
    await expect(
      applyPdfPageOrganizationJob({
        kind: 'export',
        operation: { kind: 'split', splitAfterPageIndexes: [1] },
        source,
      }),
    ).rejects.toMatchObject({ code: 'pdf.pages.invalid-split' });
    await expect(
      applyPdfPageOrganizationJob({
        kind: 'mutate',
        mutation: { index: 1, kind: 'merge' },
        source,
      }),
    ).rejects.toMatchObject({ code: 'pdf.pages.missing-import' });
  });

  test('fails closed for signed sources and risky imported catalog structures', async () => {
    const signed = await PDFDocument.create();
    signed.addPage([300, 600]);
    signed.catalog.set(PDFName.of('Perms'), signed.context.obj({}));
    await expect(
      mutate(await signed.save(), {
        kind: 'rotate',
        pageIndexes: [0],
        degrees: 90,
      }),
    ).rejects.toMatchObject({ code: 'pdf.pages.signed-source' });

    const imported = await PDFDocument.create();
    imported.addPage([400, 600]);
    imported.catalog.set(
      PDFName.of('AcroForm'),
      imported.context.obj({ Fields: [] }),
    );
    await expect(
      applyPdfPageOrganizationJob({
        importSource: await imported.save(),
        kind: 'mutate',
        mutation: { index: 1, kind: 'merge' },
        source: await createPdf([301], 'Primary'),
      }),
    ).rejects.toMatchObject({ code: 'pdf.pages.risky-import' });
  });

  test('exposes typed errors with stable codes and bounded user-safe messages', () => {
    const error = new PdfPageOrganizationError(
      'pdf.pages.invalid-index',
      'Page index 9 is outside the current 2-page PDF.',
    );
    expect(error).toMatchObject({
      code: 'pdf.pages.invalid-index',
      message: 'Page index 9 is outside the current 2-page PDF.',
      name: 'PdfPageOrganizationError',
    });
  });
});

async function mutate(
  source: ArrayBuffer | Uint8Array,
  mutation: Extract<PdfPageOrganizationJob, { kind: 'mutate' }>['mutation'],
) {
  const result = await applyPdfPageOrganizationJob({
    kind: 'mutate',
    mutation,
    source,
  });
  if (result.kind !== 'mutated') throw new Error('Expected a mutation result.');
  return result;
}

async function createPdf(
  pageWidths: readonly number[],
  title: string,
): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.setTitle(title);
  pdf.setAuthor('A3S Lab');
  pdf.setSubject('Deterministic page-organization fixture');
  pdf.setKeywords(['A3S Office', 'page organization']);
  pdf.setProducer('A3S Office tests');
  pdf.setCreator('A3S Office tests');
  pdf.setCreationDate(new Date('2026-01-01T00:00:00.000Z'));
  pdf.setModificationDate(new Date('2026-01-01T00:00:00.000Z'));
  for (let index = 0; index < pageWidths.length; index += 1) {
    const page = pdf.addPage([pageWidths[index] ?? 300, 600]);
    page.node.set(PDFName.of('A3SPage'), PDFNumber.of(index + 1));
    page.node.set(PDFName.of('A3SLabel'), PDFString.of(`page-${index + 1}`));
  }
  return pdf.save({ useObjectStreams: false });
}

async function pageEvidence(
  source: ArrayBuffer | Uint8Array | undefined,
): Promise<Array<[number, number, number]>> {
  if (!source) throw new Error('Missing PDF source.');
  const pdf = await PDFDocument.load(source);
  return pdf
    .getPages()
    .map((page) => [
      page.getWidth(),
      page.getHeight(),
      page.getRotation().angle,
    ]);
}

async function pageWidths(
  source: ArrayBuffer | Uint8Array | undefined,
): Promise<number[]> {
  return (await pageEvidence(source)).map(([width]) => width);
}
