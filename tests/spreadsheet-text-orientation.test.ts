import { describe, expect, test } from '@rstest/core';
import {
  spreadsheetTextOrientationCellStyle,
  spreadsheetTextOrientationChoiceFromCell,
  spreadsheetTextOrientationFromAngle,
  spreadsheetTextOrientationFromCell,
  spreadsheetTextOrientationFromXlsx,
  spreadsheetTextOrientationXlsxValue,
} from '../src/internal/features/work/work-spreadsheet-text-orientation';

describe('spreadsheet text orientation', () => {
  test('converts visible angles through Fortune and OOXML without changing direction', () => {
    for (const [angle, raw] of [
      [30, 30],
      [-30, 120],
      [-45, 135],
      [-90, 180],
    ] as const) {
      const orientation = spreadsheetTextOrientationFromAngle(angle);

      expect(orientation).not.toBeNull();
      expect(spreadsheetTextOrientationCellStyle(orientation)).toEqual({
        rt: raw,
      });
      expect(spreadsheetTextOrientationXlsxValue(orientation)).toBe(raw);
      expect(spreadsheetTextOrientationFromXlsx(raw)).toEqual(orientation);
      expect(spreadsheetTextOrientationFromCell({ rt: raw })).toEqual(
        orientation,
      );
    }
  });

  test('uses one stacked-text representation for Fortune and OOXML', () => {
    const orientation = spreadsheetTextOrientationFromXlsx(255);

    expect(orientation).toEqual({ kind: 'stacked' });
    expect(spreadsheetTextOrientationCellStyle(orientation)).toEqual({
      tr: '3',
    });
    expect(spreadsheetTextOrientationXlsxValue(orientation)).toBe(255);
    expect(spreadsheetTextOrientationFromCell({ tr: '3' })).toEqual(
      orientation,
    );
    expect(spreadsheetTextOrientationFromCell({ tr: '255' })).toEqual(
      orientation,
    );
    expect(spreadsheetTextOrientationChoiceFromCell({ tr: '3' })).toBe(
      'vertical',
    );
  });

  test('normalizes legacy Fortune presets and raw tr values', () => {
    expect(
      ['0', '1', '2', '3', '4', '5'].map((tr) =>
        spreadsheetTextOrientationFromCell({ tr }),
      ),
    ).toEqual([
      { angle: 0, kind: 'rotation' },
      { angle: 45, kind: 'rotation' },
      { angle: -45, kind: 'rotation' },
      { kind: 'stacked' },
      { angle: 90, kind: 'rotation' },
      { angle: -90, kind: 'rotation' },
    ]);
    expect(spreadsheetTextOrientationFromCell({ tr: '120' })).toEqual({
      angle: -30,
      kind: 'rotation',
    });
    expect(spreadsheetTextOrientationChoiceFromCell({ rt: 30 })).toBeNull();
  });

  test('rejects invalid angles and OOXML encodings', () => {
    expect(spreadsheetTextOrientationFromAngle(-91)).toBeNull();
    expect(spreadsheetTextOrientationFromAngle(30.5)).toBeNull();
    expect(spreadsheetTextOrientationFromXlsx(181)).toBeNull();
    expect(spreadsheetTextOrientationFromXlsx(Number.NaN)).toBeNull();
  });
});
