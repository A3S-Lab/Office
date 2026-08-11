import type { ITableCellOptions } from 'docx';
import {
  type DocumentTableBorder,
  documentTableBordersFromElement,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeDocumentTableVerticalAlign,
  normalizeTableColor,
} from './work-document-table-cell-formatting';
import { documentTableCellMarginOverridesFromElement } from './work-document-table-geometry';
import {
  type DocxThemePatchCollector,
  parseDocxThemeReference,
} from './work-docx-theme-reference';

type TableCellPresentationOptions = Partial<
  Pick<ITableCellOptions, 'borders' | 'margins' | 'shading' | 'verticalAlign'>
>;

const TWIPS_PER_PIXEL = 1440 / 96;

export function documentTableCellDocxOptions(
  cell: HTMLTableCellElement,
  docx: typeof import('docx'),
  themePatches?: DocxThemePatchCollector,
): TableCellPresentationOptions {
  const backgroundColor = normalizeTableColor(
    cell.dataset.officeCellFill || cell.style.backgroundColor,
  );
  const verticalAlign = normalizeDocumentTableVerticalAlign(
    cell.dataset.officeCellVerticalAlign || cell.style.verticalAlign,
  );
  const legacyBorderStyle = normalizeDocumentTableBorderStyle(
    cell.dataset.officeCellBorderStyle || cell.style.borderStyle,
  );
  const fallbackBorder: DocumentTableBorder = {
    color:
      normalizeTableColor(
        cell.dataset.officeCellBorderColor || cell.style.borderColor,
      ) ?? '#cfd5df',
    style: legacyBorderStyle ?? 'solid',
    width:
      legacyBorderStyle === 'none'
        ? 0
        : (normalizeDocumentTableBorderWidth(
            cell.dataset.officeCellBorderWidth || cell.style.borderWidth,
          ) ?? 1),
  };
  const borders = hasDocumentTableBorderPresentation(cell)
    ? documentTableBordersFromElement(cell, fallbackBorder)
    : null;
  const margins = documentTableCellMarginOverridesFromElement(cell);
  const fillMarker = themePatches?.marker(
    'fill',
    parseDocxThemeReference(cell.dataset.officeCellThemeFill),
    backgroundColor,
  );

  return {
    ...(backgroundColor
      ? {
          shading: {
            fill: fillMarker ?? docxColor(backgroundColor),
            type: docx.ShadingType.CLEAR,
          },
        }
      : {}),
    ...(verticalAlign
      ? {
          verticalAlign:
            verticalAlign === 'middle'
              ? docx.VerticalAlignTable.CENTER
              : verticalAlign === 'bottom'
                ? docx.VerticalAlignTable.BOTTOM
                : docx.VerticalAlignTable.TOP,
        }
      : {}),
    ...(borders
      ? {
          borders: {
            top: documentTableCellDocxBorder(
              borders.top,
              'top',
              cell,
              docx,
              themePatches,
            ),
            right: documentTableCellDocxBorder(
              borders.right,
              'right',
              cell,
              docx,
              themePatches,
            ),
            bottom: documentTableCellDocxBorder(
              borders.bottom,
              'bottom',
              cell,
              docx,
              themePatches,
            ),
            left: documentTableCellDocxBorder(
              borders.left,
              'left',
              cell,
              docx,
              themePatches,
            ),
          },
        }
      : {}),
    ...(margins
      ? {
          margins: Object.fromEntries(
            Object.entries(margins).map(([side, value]) => [
              side,
              pixelsToTwips(value),
            ]),
          ),
        }
      : {}),
  };
}

function pixelsToTwips(pixels: number): number {
  return Math.max(0, Math.round(pixels * TWIPS_PER_PIXEL));
}

function documentTableCellDocxBorder(
  border: DocumentTableBorder,
  edge: 'top' | 'right' | 'bottom' | 'left',
  cell: HTMLTableCellElement,
  docx: typeof import('docx'),
  themePatches?: DocxThemePatchCollector,
) {
  const marker = themePatches?.marker(
    'border',
    parseDocxThemeReference(
      cell.dataset[`officeCellBorderTheme${capitalize(edge)}`],
    ),
    border.color,
  );
  return {
    style: docxBorderStyle(border.style, docx),
    ...(border.style === 'none' || border.width === 0
      ? {}
      : {
          color: marker ?? docxColor(border.color),
          size: Math.max(1, Math.round(border.width * 6)),
        }),
  };
}

function capitalize(value: string): string {
  return `${value.slice(0, 1).toUpperCase()}${value.slice(1)}`;
}

function hasDocumentTableBorderPresentation(
  cell: HTMLTableCellElement,
): boolean {
  return Boolean(
    Object.keys(cell.dataset).some((key) =>
      key.startsWith('officeCellBorder'),
    ) ||
      cell.style.borderStyle ||
      cell.style.borderTopStyle ||
      cell.style.borderRightStyle ||
      cell.style.borderBottomStyle ||
      cell.style.borderLeftStyle,
  );
}

function docxBorderStyle(
  style: NonNullable<ReturnType<typeof normalizeDocumentTableBorderStyle>>,
  docx: typeof import('docx'),
) {
  if (style === 'dashed') return docx.BorderStyle.DASHED;
  if (style === 'dotted') return docx.BorderStyle.DOTTED;
  if (style === 'double') return docx.BorderStyle.DOUBLE;
  if (style === 'none') return docx.BorderStyle.NONE;
  return docx.BorderStyle.SINGLE;
}

function docxColor(color: string): string {
  return color.replace(/^#/, '').toUpperCase();
}
