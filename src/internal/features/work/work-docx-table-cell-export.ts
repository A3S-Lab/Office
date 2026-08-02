import type { ITableCellOptions } from 'docx';
import {
  type DocumentTableBorder,
  documentTableBordersFromElement,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeDocumentTableVerticalAlign,
  normalizeTableColor,
} from './work-document-table-cell-formatting';

type TableCellPresentationOptions = Partial<
  Pick<ITableCellOptions, 'borders' | 'shading' | 'verticalAlign'>
>;

export function documentTableCellDocxOptions(
  cell: HTMLTableCellElement,
  docx: typeof import('docx'),
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

  return {
    ...(backgroundColor
      ? {
          shading: {
            fill: docxColor(backgroundColor),
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
            top: documentTableCellDocxBorder(borders.top, docx),
            right: documentTableCellDocxBorder(borders.right, docx),
            bottom: documentTableCellDocxBorder(borders.bottom, docx),
            left: documentTableCellDocxBorder(borders.left, docx),
          },
        }
      : {}),
  };
}

function documentTableCellDocxBorder(
  border: DocumentTableBorder,
  docx: typeof import('docx'),
) {
  return {
    style: docxBorderStyle(border.style, docx),
    ...(border.style === 'none' || border.width === 0
      ? {}
      : {
          color: docxColor(border.color),
          size: Math.max(1, Math.round(border.width * 6)),
        }),
  };
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
