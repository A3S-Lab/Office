import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  OoxmlPackage,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import {
  isIdentitySpreadsheetImageTransform,
  normalizeSpreadsheetImageTransform,
  readXlsxImageTransform,
  xlsxImageTransformAttributes,
} from '../src/internal/features/work/work-xlsx-image-transform';
import {
  patchXlsxWorksheetDrawings,
  readXlsxWorksheetImages,
} from '../src/internal/features/work/work-xlsx-images';

describe('R2 worksheet image transforms', () => {
  test('admits quadrant rotations and flips while rejecting arbitrary angles', () => {
    expect(
      normalizeSpreadsheetImageTransform({
        rotation: 90,
        flipHorizontal: true,
        flipVertical: false,
      }),
    ).toEqual({
      rotation: 90,
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(
      isIdentitySpreadsheetImageTransform({
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      }),
    ).toBe(true);

    const supported = parseXml(
      `<xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" rot="5400000" flipH="1"/>`,
      'xfrm',
    ).documentElement;
    expect(readXlsxImageTransform(supported)).toEqual({
      transform: {
        rotation: 90,
        flipHorizontal: true,
        flipVertical: false,
      },
      supported: true,
    });

    const unsupported = parseXml(
      `<xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" rot="2700000"/>`,
      'xfrm',
    ).documentElement;
    expect(readXlsxImageTransform(unsupported).supported).toBe(false);
    expect(
      xlsxImageTransformAttributes({
        rotation: 180,
        flipHorizontal: false,
        flipVertical: true,
      }),
    ).toBe(' rot="10800000" flipV="1"');
  });

  test('round-trips supported transforms through worksheet drawings', async () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          status: 1,
          data: [[{ v: 1 }]],
          images: [
            {
              id: 'image-1',
              src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
              contentType: 'image/png',
              name: 'Rotated',
              left: 0,
              top: 0,
              width: 64,
              height: 64,
              transform: {
                rotation: 90,
                flipHorizontal: false,
                flipVertical: true,
              },
            },
          ],
        },
      ],
    };

    const seeded = await seedWorkbookWithBlankSheet();
    const patched = await patchXlsxWorksheetDrawings(seeded, content);
    const archive = await OoxmlPackage.load(patched);
    const images = await readXlsxWorksheetImages(
      archive,
      'xl/worksheets/sheet1.xml',
      await archive.xml('xl/worksheets/sheet1.xml'),
      { bytes: 0 },
    );
    expect(images).toHaveLength(1);
    expect(images[0]?.transform).toEqual({
      rotation: 90,
      flipHorizontal: false,
      flipVertical: true,
    });
    expect(images[0]?.unsupportedTransform).toBeUndefined();
  });
});

async function seedWorkbookWithBlankSheet(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
