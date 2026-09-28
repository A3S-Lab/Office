import { expect, test } from '@rstest/core';
import {
  listRealCorpusFixtures,
  roundTripRealCorpusFile,
} from './harness';

test('real corpus fixtures are present and licensed for redistribution', async () => {
  const fixtures = await listRealCorpusFixtures();
  expect(fixtures).toEqual(
    expect.arrayContaining([
      'SampleDoc.docx',
      'charts.xlsx',
      'SampleShow.pptx',
      'PDFBOX-4352-0.pdf',
      'FormulaEvalTestData.xlsx',
    ]),
  );
});

test(
  'real corpus OOXML fixtures survive import → export → reopen',
  async () => {
    const fixtures = (await listRealCorpusFixtures()).filter((name) =>
      /\.(docx|xlsx|pptx)$/i.test(name),
    );
    expect(fixtures.length).toBeGreaterThan(0);

    for (const fileName of fixtures) {
      const result = await roundTripRealCorpusFile(fileName);
      expect(result.kind).toMatch(/document|spreadsheet|presentation/);
      expect(result.firstPassParts.length).toBeGreaterThan(0);
      expect(result.secondPassParts.length).toBeGreaterThan(0);
      // Re-export must not drop the package root the first pass observed.
      for (const required of requiredPartsFor(result.kind)) {
        expect(result.firstPassParts).toContain(required);
        expect(result.secondPassParts).toContain(required);
      }
    }
  },
  60_000,
);

test(
  'real corpus PDF fixtures survive import → export → reopen',
  async () => {
    const fixtures = (await listRealCorpusFixtures()).filter((name) =>
      /\.pdf$/i.test(name),
    );
    expect(fixtures.length).toBeGreaterThan(0);
    for (const fileName of fixtures) {
      const result = await roundTripRealCorpusFile(fileName);
      expect(result.kind).toBe('pdf');
      expect(result.firstPassParts.some((part) => part.startsWith('pdf:pages=')))
        .toBe(true);
      expect(result.secondPassParts).toEqual(result.firstPassParts);
    }
  },
  30_000,
);

function requiredPartsFor(
  kind: 'document' | 'spreadsheet' | 'presentation' | 'pdf',
): string[] {
  switch (kind) {
    case 'document':
      return ['[Content_Types].xml', 'word/document.xml'];
    case 'spreadsheet':
      return ['[Content_Types].xml', 'xl/workbook.xml'];
    case 'presentation':
      return ['[Content_Types].xml', 'ppt/presentation.xml'];
    case 'pdf':
      return [];
  }
}
