import type { ITableCellOptions } from 'docx';
import {
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
  const borderStyle = normalizeDocumentTableBorderStyle(
    cell.dataset.officeCellBorderStyle || cell.style.borderStyle,
  );
  const borderWidth = normalizeDocumentTableBorderWidth(
    cell.dataset.officeCellBorderWidth || cell.style.borderWidth,
  );
  const borderColor = normalizeTableColor(
    cell.dataset.officeCellBorderColor || cell.style.borderColor,
  );
  const border = borderStyle
    ? {
        style: docxBorderStyle(borderStyle, docx),
        ...(borderStyle === 'none'
          ? {}
          : {
              color: docxColor(borderColor ?? '#cfd5df'),
              size: Math.max(1, Math.round((borderWidth ?? 1) * 6)),
            }),
      }
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
    ...(border
      ? {
          borders: {
            top: border,
            right: border,
            bottom: border,
            left: border,
          },
        }
      : {}),
  };
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
