import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';
import {
  patchXlsxSheetFeatures,
  readXlsxSheetFeaturesFromPackage,
} from '../src/internal/features/work/work-xlsx-interop';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet data-validation XLSX interop', () => {
  test('imports decimal and 1904-date rules with Fortune-compatible semantics', async () => {
    const archive = await OoxmlPackage.load(await validationWorkbook());
    const rules = (await readXlsxSheetFeaturesFromPackage(archive)).get(
      'Rules',
    )?.validations;

    expect(rules).toEqual([
      {
        references: ['A1:A4'],
        item: expect.objectContaining({
          type: 'number',
          type2: 'moreThanThe',
          value1: '1.5',
        }),
      },
      {
        references: ['B1:B4'],
        item: expect.objectContaining({
          type: 'date',
          type2: 'noEarlierThan',
          value1: '1904-01-02',
        }),
      },
      {
        references: ['C1:C4'],
        item: expect.objectContaining({
          type: 'date',
          type2: 'between',
          value1: '2026-08-21',
          value2: '2026-12-31',
        }),
      },
      {
        references: ['D1:D4'],
        item: expect.objectContaining({
          type: 'dropdown',
          value1: 'Ready,Blocked',
          allowBlank: false,
          showDropdownArrow: false,
          prohibitInput: true,
          errorStyle: 'warning',
          errorTitle: 'Invalid state',
          errorMessage: 'Choose Ready or Blocked.',
          hintShow: true,
          hintTitle: 'Workflow state',
          hintValue: 'Choose a state.',
        }),
      },
    ]);
  });

  test('exports ISO date boundaries as valid OOXML formulas', async () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Rules',
          dataValidationRanges: [
            {
              ranges: [{ row: [0, 3], column: [2, 2] }],
              item: {
                type: 'date',
                type2: 'between',
                rangeTxt: 'C1:C4',
                value1: '2026-08-21',
                value2: '2026-12-31',
                validity: '',
                remote: false,
                allowBlank: false,
                showDropdownArrow: true,
                prohibitInput: true,
                errorStyle: 'information',
                errorTitle: 'Date required',
                errorMessage: 'Enter a date in 2026.',
                hintShow: true,
                hintTitle: '2026 date',
                hintValue: 'Use a 2026 date.',
              },
            },
          ],
        },
      ],
    };
    const patched = await patchXlsxSheetFeatures(
      await validationWorkbook(),
      content,
    );
    const zip = await JSZip.loadAsync(patched);
    const worksheet =
      (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? '';

    expect(worksheet).toContain('type="date"');
    expect(worksheet).toContain('operator="between"');
    expect(worksheet).toContain('sqref="C1:C4"');
    expect(worksheet).toContain('allowBlank="0"');
    expect(worksheet).toContain('showErrorMessage="1"');
    expect(worksheet).toContain('errorStyle="information"');
    expect(worksheet).toContain('errorTitle="Date required"');
    expect(worksheet).toContain('error="Enter a date in 2026."');
    expect(worksheet).toContain('showInputMessage="1"');
    expect(worksheet).toContain('promptTitle="2026 date"');
    expect(worksheet).toContain('prompt="Use a 2026 date."');
    expect(worksheet).toContain('<formula1>DATE(2026,8,21)</formula1>');
    expect(worksheet).toContain('<formula2>DATE(2026,12,31)</formula2>');
  });

  test('exports named list sources as formulas and single literals as inline lists', async () => {
    const dropdown = {
      type: 'dropdown',
      type2: '',
      rangeTxt: '',
      value2: '',
      validity: '',
      remote: false,
      allowBlank: true,
      showDropdownArrow: true,
      prohibitInput: true,
      errorStyle: 'stop' as const,
      errorTitle: '',
      errorMessage: '',
      hintShow: false,
      hintTitle: '',
      hintValue: '',
    };
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      namedRanges: [
        {
          id: 'name-1',
          name: 'WorkflowStates',
          reference: "'Rules'!$D$1:$D$4",
        },
      ],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Rules',
          dataValidationRanges: [
            {
              ranges: [{ row: [0, 3], column: [0, 0] }],
              item: {
                ...dropdown,
                rangeTxt: 'A1:A4',
                value1: 'WorkflowStates',
              },
            },
            {
              ranges: [{ row: [0, 3], column: [1, 1] }],
              item: {
                ...dropdown,
                rangeTxt: 'B1:B4',
                value1: 'Ready',
              },
            },
          ],
        },
      ],
    };

    const patched = await patchXlsxSheetFeatures(
      await validationWorkbook(),
      content,
    );
    const zip = await JSZip.loadAsync(patched);
    const worksheet =
      (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? '';

    expect(worksheet).toContain('<formula1>WorkflowStates</formula1>');
    expect(worksheet).toContain('<formula1>"Ready"</formula1>');
    expect(worksheet).not.toContain('<formula1>"WorkflowStates"</formula1>');
    expect(
      (await readXlsxSheetFeaturesFromPackage(await OoxmlPackage.load(patched)))
        .get('Rules')
        ?.validations.map((validation) => validation.item.value1),
    ).toEqual(['WorkflowStates', 'Ready']);
  });
});

async function validationWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    'xl/workbook.xml',
    [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<workbookPr date1904="1"/>',
      '<sheets><sheet name="Rules" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1"',
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"',
      ' Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1:D4"/><sheetData/>',
      '<dataValidations count="4">',
      '<dataValidation type="decimal" operator="greaterThan" sqref="A1:A4">',
      '<formula1>1.5</formula1></dataValidation>',
      '<dataValidation type="date" operator="greaterThanOrEqual" sqref="B1:B4">',
      '<formula1>1</formula1></dataValidation>',
      '<dataValidation type="date" operator="between" sqref="C1:C4">',
      '<formula1>DATE(2026,8,21)</formula1>',
      '<formula2>DATE(2026,12,31)</formula2></dataValidation>',
      '<dataValidation type="list" allowBlank="0" showDropDown="1"',
      ' showErrorMessage="1" errorStyle="warning" errorTitle="Invalid state"',
      ' error="Choose Ready or Blocked." showInputMessage="1"',
      ' promptTitle="Workflow state" prompt="Choose a state." sqref="D1:D4">',
      '<formula1>"Ready,Blocked"</formula1></dataValidation>',
      '</dataValidations>',
      '</worksheet>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
