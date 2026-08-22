import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';
import {
  withXlsxGradientFill,
  type XlsxGradientFill,
} from '../../src/internal/features/work/work-xlsx-gradient-fill';

export const SPREADSHEET_GRADIENT_FILL_FIXTURE = 'spreadsheet-gradient-fill';
export const SPREADSHEET_GRADIENT_FILL_ARTIFACT_ID =
  'fixture-spreadsheet-gradient-fill';

const gradientColors = [
  '#1d4ed8',
  '#b42318',
  '#067647',
  '#7a5af8',
  '#c2410c',
  '#0e7490',
] as const;

const gradientGallery: Array<{ fill: XlsxGradientFill; label: string }> = [
  ...[0, 45, 90, 135].map((degree, index) => ({
    fill: {
      degree,
      stops: [
        { color: gradientColors[index] ?? '#1d4ed8', position: 0 },
        { color: '#ffffff', position: 1 },
      ],
      type: 'linear' as const,
    },
    label: `Linear ${degree}°`,
  })),
  {
    fill: {
      bottom: 0.75,
      left: 0.25,
      right: 0.75,
      stops: [
        { color: gradientColors[4], position: 0 },
        { color: '#ffffff', position: 1 },
      ],
      top: 0.25,
      type: 'path',
    },
    label: 'Path inset',
  },
  {
    fill: {
      bottom: 0.65,
      left: 0.1,
      right: 0.9,
      stops: [
        { color: gradientColors[5], position: 0 },
        { color: '#67e8f9', position: 0.45 },
        { color: '#ffffff', position: 1 },
      ],
      top: 0.35,
      type: 'path',
    },
    label: 'Path multistop',
  },
];

export function createSpreadsheetGradientFillArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data[0] = [
    {
      v: 'Native XLSX gradient fills',
      m: 'Native XLSX gradient fills',
      bl: 1,
      fc: '#172033',
      fs: 15,
      bg: '#eaf0ff',
    },
  ];
  for (const [index, entry] of gradientGallery.entries()) {
    const row = 2 + Math.floor(index / 3) * 2;
    const column = index % 3;
    data[row] ??= [];
    data[row][column] = withXlsxGradientFill(
      {
        v: entry.label,
        m: entry.label,
        bg: entry.fill.stops[0]?.color ?? '#ffffff',
        fc: '#172033',
        fs: 10,
        ht: 0,
        vt: 0,
      },
      entry.fill,
    );
  }

  return {
    id: SPREADSHEET_GRADIENT_FILL_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Native gradient-fill gallery',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'gradient-fill-sheet',
          name: 'Gradient Fills',
          status: 1,
          row: 10,
          column: 6,
          data,
          config: {
            columnlen: { 0: 180, 1: 180, 2: 180 },
            rowlen: { 0: 34, 2: 58, 4: 58 },
            merge: { '0_0': { r: 0, c: 0, rs: 1, cs: 3 } },
          },
          luckysheet_select_save: [
            {
              row: [2, 2],
              column: [0, 0],
              row_focus: 2,
              column_focus: 0,
            },
          ],
        },
      ],
    },
  };
}
