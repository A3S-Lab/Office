import type { Cell } from '@fortune-sheet/core';
import type { XlsxCellBorder } from './work-xlsx-cell-borders';
import {
  defaultXlsxBorder,
  defaultXlsxFill,
  ensureXlsxStyleCollection,
  setXlsxBorderLine,
  setXlsxColorChild,
  setXlsxToggleChild,
  setXlsxUnderlineChild,
  setXlsxValueChild,
  type XlsxDirectAlignmentStyle,
  writeXlsxAlignment,
  xlsxRgbColor,
} from './work-xlsx-cell-style-xml';
import { attribute, directChildren } from './work-ooxml-package';
import {
  spreadsheetUnderlineStyle,
  type SpreadsheetUnderlineStyle,
} from './work-spreadsheet-underline';

const fontChildOrder = [
  'name',
  'charset',
  'family',
  'b',
  'i',
  'strike',
  'outline',
  'shadow',
  'condense',
  'extend',
  'color',
  'sz',
  'u',
  'vertAlign',
  'scheme',
] as const;

interface XlsxDirectFontStyle {
  bold?: boolean;
  color?: string;
  italic?: boolean;
  name?: string;
  size?: number;
  strike?: boolean;
  underline?: SpreadsheetUnderlineStyle;
}

export class XlsxDirectCellStyleWriter {
  private readonly fonts: Element;
  private readonly fills: Element;
  private readonly borders: Element;
  private readonly cellXfs: Element;
  private readonly generatedFonts = new Map<string, number>();
  private readonly generatedFills = new Map<string, number>();
  private readonly generatedBorders = new Map<string, number>();
  private readonly generatedStyles = new Map<string, number>();
  changed = false;

  constructor(private readonly styles: Document) {
    const root = styles.documentElement;
    this.fonts = ensureXlsxStyleCollection(styles, 'fonts', [
      'fills',
      'borders',
      'cellStyleXfs',
      'cellXfs',
      'cellStyles',
      'dxfs',
      'tableStyles',
      'colors',
      'extLst',
    ]);
    if (!directChildren(this.fonts, 'font').length) {
      this.fonts.append(styles.createElementNS(root.namespaceURI, 'font'));
      this.changed = true;
    }
    this.fills = ensureXlsxStyleCollection(styles, 'fills', [
      'borders',
      'cellStyleXfs',
      'cellXfs',
      'cellStyles',
      'dxfs',
      'tableStyles',
      'colors',
      'extLst',
    ]);
    if (!directChildren(this.fills, 'fill').length) {
      this.fills.append(defaultXlsxFill(styles, 'none'));
      this.fills.append(defaultXlsxFill(styles, 'gray125'));
      this.changed = true;
    }
    this.borders = ensureXlsxStyleCollection(styles, 'borders', [
      'cellStyleXfs',
      'cellXfs',
      'cellStyles',
      'dxfs',
      'tableStyles',
      'colors',
      'extLst',
    ]);
    if (!directChildren(this.borders, 'border').length) {
      this.borders.append(defaultXlsxBorder(styles));
      this.changed = true;
    }
    this.cellXfs = ensureXlsxStyleCollection(styles, 'cellXfs', [
      'cellStyles',
      'dxfs',
      'tableStyles',
      'colors',
      'extLst',
    ]);
    if (!directChildren(this.cellXfs, 'xf').length) {
      const base = styles.createElementNS(root.namespaceURI, 'xf');
      base.setAttribute('numFmtId', '0');
      base.setAttribute('fontId', '0');
      base.setAttribute('fillId', '0');
      base.setAttribute('borderId', '0');
      base.setAttribute('xfId', '0');
      this.cellXfs.append(base);
      this.changed = true;
    }
    this.updateCounts();
  }

  styleId(baseStyleId: number, cell: Cell, border?: XlsxCellBorder): number {
    const styles = directChildren(this.cellXfs, 'xf');
    const baseIndex =
      Number.isInteger(baseStyleId) && styles[baseStyleId] ? baseStyleId : 0;
    const base = styles[baseIndex];
    const fontId = this.fontId(
      nonNegativeInteger(attribute(base, 'fontId')) ?? 0,
      cell,
    );
    const fillId = this.fillId(
      nonNegativeInteger(attribute(base, 'fillId')) ?? 0,
      cell,
    );
    const borderId = this.borderId(
      nonNegativeInteger(attribute(base, 'borderId')) ?? 0,
      border,
    );
    const alignment = directAlignment(cell);
    const key = `${baseIndex}:${fontId}:${fillId}:${borderId}:${JSON.stringify(alignment)}`;
    const cached = this.generatedStyles.get(key);
    if (cached !== undefined) return cached;

    const clone = base.cloneNode(true) as Element;
    clone.setAttribute('fontId', String(fontId));
    clone.setAttribute('fillId', String(fillId));
    clone.setAttribute('borderId', String(borderId));
    if (hasXlsxDirectFontStyle(cell)) clone.setAttribute('applyFont', '1');
    if (cell.bg !== undefined) clone.setAttribute('applyFill', '1');
    if (border) clone.setAttribute('applyBorder', '1');
    if (alignment) {
      writeXlsxAlignment(this.styles, clone, alignment);
      clone.setAttribute('applyAlignment', '1');
    }
    const index = styles.length;
    this.cellXfs.append(clone);
    this.generatedStyles.set(key, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private fontId(baseFontId: number, cell: Cell): number {
    if (!hasXlsxDirectFontStyle(cell)) return baseFontId;
    const fonts = directChildren(this.fonts, 'font');
    const baseIndex = fonts[baseFontId] ? baseFontId : 0;
    const style = directFontStyle(cell);
    const key = `${baseIndex}:${JSON.stringify(style)}`;
    const cached = this.generatedFonts.get(key);
    if (cached !== undefined) return cached;
    const font = fonts[baseIndex].cloneNode(true) as Element;
    if (style.name !== undefined)
      setXlsxValueChild(this.styles, font, 'name', style.name, fontChildOrder);
    if (style.size !== undefined)
      setXlsxValueChild(
        this.styles,
        font,
        'sz',
        String(style.size),
        fontChildOrder,
      );
    if (style.color !== undefined)
      setXlsxColorChild(this.styles, font, style.color, fontChildOrder);
    if (style.bold !== undefined)
      setXlsxToggleChild(this.styles, font, 'b', style.bold, fontChildOrder);
    if (style.italic !== undefined)
      setXlsxToggleChild(this.styles, font, 'i', style.italic, fontChildOrder);
    if (style.strike !== undefined)
      setXlsxToggleChild(
        this.styles,
        font,
        'strike',
        style.strike,
        fontChildOrder,
      );
    if (style.underline !== undefined)
      setXlsxUnderlineChild(this.styles, font, style.underline, fontChildOrder);
    const index = fonts.length;
    this.fonts.append(font);
    this.generatedFonts.set(key, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private fillId(baseFillId: number, cell: Cell): number {
    if (cell.bg === undefined) return baseFillId;
    const color = xlsxRgbColor(cell.bg);
    if (!color) return baseFillId;
    const cached = this.generatedFills.get(color);
    if (cached !== undefined) return cached;
    const fill = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'fill',
    );
    const pattern = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'patternFill',
    );
    pattern.setAttribute('patternType', 'solid');
    const foreground = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'fgColor',
    );
    foreground.setAttribute('rgb', color);
    const background = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'bgColor',
    );
    background.setAttribute('indexed', '64');
    pattern.append(foreground, background);
    fill.append(pattern);
    const index = directChildren(this.fills, 'fill').length;
    this.fills.append(fill);
    this.generatedFills.set(color, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private borderId(
    baseBorderId: number,
    update: XlsxCellBorder | undefined,
  ): number {
    if (!update) return baseBorderId;
    const borders = directChildren(this.borders, 'border');
    const baseIndex = borders[baseBorderId] ? baseBorderId : 0;
    const key = `${baseIndex}:${JSON.stringify(update)}`;
    const cached = this.generatedBorders.get(key);
    if (cached !== undefined) return cached;
    const border = borders[baseIndex].cloneNode(true) as Element;
    for (const [name, line] of [
      ['left', update.left],
      ['right', update.right],
      ['top', update.top],
      ['bottom', update.bottom],
      ['diagonal', update.diagonal],
    ] as const) {
      if (line !== undefined) {
        setXlsxBorderLine(this.styles, border, name, line);
      }
    }
    if (update.diagonalUp !== undefined) {
      border.setAttribute('diagonalUp', update.diagonalUp ? '1' : '0');
    }
    if (update.diagonalDown !== undefined) {
      border.setAttribute('diagonalDown', update.diagonalDown ? '1' : '0');
    }
    const index = borders.length;
    this.borders.append(border);
    this.generatedBorders.set(key, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private updateCounts(): void {
    this.fonts.setAttribute(
      'count',
      String(directChildren(this.fonts, 'font').length),
    );
    this.fills.setAttribute(
      'count',
      String(directChildren(this.fills, 'fill').length),
    );
    this.borders.setAttribute(
      'count',
      String(directChildren(this.borders, 'border').length),
    );
    this.cellXfs.setAttribute(
      'count',
      String(directChildren(this.cellXfs, 'xf').length),
    );
  }
}

export function hasXlsxDirectCellStyle(cell: Cell): boolean {
  return Boolean(
    hasXlsxDirectFontStyle(cell) ||
      cell.bg !== undefined ||
      cell.ht !== undefined ||
      cell.vt !== undefined ||
      cell.tb !== undefined ||
      cell.tr !== undefined,
  );
}

function hasXlsxDirectFontStyle(cell: Cell): boolean {
  return Boolean(
    cell.bl !== undefined ||
      cell.it !== undefined ||
      cell.un !== undefined ||
      cell.cl !== undefined ||
      cell.ff !== undefined ||
      cell.fs !== undefined ||
      cell.fc !== undefined,
  );
}

function directFontStyle(cell: Cell): XlsxDirectFontStyle {
  const color = cell.fc !== undefined ? xlsxRgbColor(cell.fc) : null;
  return {
    ...(cell.bl !== undefined ? { bold: Number(cell.bl) === 1 } : {}),
    ...(color ? { color } : {}),
    ...(cell.it !== undefined ? { italic: Number(cell.it) === 1 } : {}),
    ...(typeof cell.ff === 'string' && cell.ff.trim()
      ? { name: cell.ff.trim() }
      : {}),
    ...(typeof cell.fs === 'number' && Number.isFinite(cell.fs) && cell.fs > 0
      ? { size: cell.fs }
      : {}),
    ...(cell.cl !== undefined ? { strike: Number(cell.cl) === 1 } : {}),
    ...(cell.un !== undefined
      ? { underline: spreadsheetUnderlineStyle(cell.un) }
      : {}),
  };
}

function directAlignment(cell: Cell): XlsxDirectAlignmentStyle | null {
  const alignment: XlsxDirectAlignmentStyle = {};
  if (cell.ht !== undefined) {
    alignment.horizontal =
      Number(cell.ht) === 0
        ? 'center'
        : Number(cell.ht) === 2
          ? 'right'
          : 'left';
  }
  if (cell.vt !== undefined) {
    alignment.vertical =
      Number(cell.vt) === 0
        ? 'center'
        : Number(cell.vt) === 1
          ? 'top'
          : 'bottom';
  }
  if (cell.tb !== undefined) alignment.wrapText = cell.tb === '2';
  if (cell.tr !== undefined) {
    const rotation = Number(cell.tr);
    if (Number.isFinite(rotation) && rotation >= 0 && rotation <= 180) {
      alignment.textRotation = Math.round(rotation);
    }
  }
  return Object.keys(alignment).length ? alignment : null;
}

function nonNegativeInteger(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}
