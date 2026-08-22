import type { Cell } from '@fortune-sheet/core';
import { attribute, directChild, directChildren } from './work-ooxml-package';
import type { XlsxCellBorder } from './work-xlsx-cell-borders';
import {
  applyXlsxSemanticColorOrigin,
  type XlsxBorderColorOrigins,
  type XlsxSemanticColorOrigin,
  type XlsxSemanticPalette,
  xlsxCellStyleOrigin,
  xlsxColorElementMatchesOrigin,
  xlsxSemanticColorOriginKey,
} from './work-xlsx-cell-style-origin';
import {
  activeXlsxSemanticColorOrigin,
  directXlsxAlignment,
  directXlsxFontStyle,
  hasXlsxDirectFontStyle,
  xlsxAlignmentMatches,
  xlsxBooleanAttribute,
  xlsxBorderLineMatches,
  xlsxColorMatches,
  xlsxStyleCollectionIndex,
  xlsxToggleEnabled,
  xlsxUnderlineStyle,
} from './work-xlsx-cell-style-values';
import {
  defaultXlsxBorder,
  defaultXlsxFill,
  ensureXlsxStyleCollection,
  setXlsxBorderLine,
  setXlsxColorChild,
  setXlsxToggleChild,
  setXlsxUnderlineChild,
  setXlsxValueChild,
  writeXlsxAlignment,
  xlsxRgbColor,
} from './work-xlsx-cell-style-xml';
import {
  createXlsxColorResolver,
  type XlsxColorResolver,
} from './work-xlsx-colors';
import {
  activeXlsxPatternFill,
  type XlsxPatternFill,
} from './work-xlsx-pattern-fill';

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

export class XlsxDirectCellStyleWriter {
  private readonly fonts: Element;
  private readonly fills: Element;
  private readonly borders: Element;
  private readonly cellXfs: Element;
  private readonly generatedFonts = new Map<string, number>();
  private readonly generatedFills = new Map<string, number>();
  private readonly generatedBorders = new Map<string, number>();
  private readonly generatedStyles = new Map<string, number>();
  private readonly colors: XlsxColorResolver;
  changed = false;

  constructor(
    private readonly styles: Document,
    theme: Document | null = null,
    private readonly semanticPalette?: XlsxSemanticPalette,
  ) {
    this.colors = createXlsxColorResolver(styles, theme);
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
    const baseFontId = xlsxStyleCollectionIndex(
      this.fonts,
      'font',
      attribute(base, 'fontId'),
    );
    const baseFillId = xlsxStyleCollectionIndex(
      this.fills,
      'fill',
      attribute(base, 'fillId'),
    );
    const baseBorderId = xlsxStyleCollectionIndex(
      this.borders,
      'border',
      attribute(base, 'borderId'),
    );
    const origin = xlsxCellStyleOrigin(cell);
    const fontId = this.fontId(baseFontId, cell, origin?.fontColor);
    const fillId = this.fillId(
      baseFillId,
      cell,
      origin?.fillColor,
      activeXlsxPatternFill(cell),
    );
    const borderId = this.borderId(baseBorderId, border, origin?.borderColors);
    const alignment = directXlsxAlignment(cell);
    const alignmentChanged = Boolean(
      alignment && !xlsxAlignmentMatches(base, alignment),
    );
    if (
      fontId === baseFontId &&
      fillId === baseFillId &&
      borderId === baseBorderId &&
      !alignmentChanged
    ) {
      return baseIndex;
    }
    const key = `${baseIndex}:${fontId}:${fillId}:${borderId}:${alignmentChanged ? JSON.stringify(alignment) : ''}`;
    const cached = this.generatedStyles.get(key);
    if (cached !== undefined) return cached;

    const clone = base.cloneNode(true) as Element;
    clone.setAttribute('fontId', String(fontId));
    clone.setAttribute('fillId', String(fillId));
    clone.setAttribute('borderId', String(borderId));
    if (fontId !== baseFontId) clone.setAttribute('applyFont', '1');
    if (fillId !== baseFillId) clone.setAttribute('applyFill', '1');
    if (borderId !== baseBorderId) clone.setAttribute('applyBorder', '1');
    if (alignment && alignmentChanged) {
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

  private fontId(
    baseFontId: number,
    cell: Cell,
    colorOrigin?: XlsxSemanticColorOrigin,
  ): number {
    if (!hasXlsxDirectFontStyle(cell)) return baseFontId;
    const fonts = directChildren(this.fonts, 'font');
    const baseIndex = fonts[baseFontId] ? baseFontId : 0;
    const style = directXlsxFontStyle(cell);
    const base = fonts[baseIndex];
    const semanticColor = activeXlsxSemanticColorOrigin(
      colorOrigin,
      style.color,
      this.semanticPalette,
    );
    const nameChanged =
      style.name !== undefined &&
      attribute(directChild(base, 'name') ?? base, 'val')?.trim() !==
        style.name;
    const sizeChanged =
      style.size !== undefined &&
      Number(attribute(directChild(base, 'sz') ?? base, 'val')) !== style.size;
    const colorChanged =
      style.color !== undefined &&
      (semanticColor
        ? !xlsxColorElementMatchesOrigin(
            directChild(base, 'color'),
            semanticColor,
          )
        : !xlsxColorMatches(
            directChild(base, 'color'),
            style.color,
            this.colors,
            '#000000',
          ));
    const boldChanged =
      style.bold !== undefined &&
      xlsxToggleEnabled(directChild(base, 'b')) !== style.bold;
    const italicChanged =
      style.italic !== undefined &&
      xlsxToggleEnabled(directChild(base, 'i')) !== style.italic;
    const strikeChanged =
      style.strike !== undefined &&
      xlsxToggleEnabled(directChild(base, 'strike')) !== style.strike;
    const underlineChanged =
      style.underline !== undefined &&
      xlsxUnderlineStyle(directChild(base, 'u')) !== style.underline;
    if (
      !nameChanged &&
      !sizeChanged &&
      !colorChanged &&
      !boldChanged &&
      !italicChanged &&
      !strikeChanged &&
      !underlineChanged
    ) {
      return baseIndex;
    }
    const key = `${baseIndex}:${JSON.stringify(style)}:${xlsxSemanticColorOriginKey(semanticColor)}`;
    const cached = this.generatedFonts.get(key);
    if (cached !== undefined) return cached;
    const font = base.cloneNode(true) as Element;
    if (nameChanged && style.name !== undefined)
      setXlsxValueChild(this.styles, font, 'name', style.name, fontChildOrder);
    if (sizeChanged && style.size !== undefined)
      setXlsxValueChild(
        this.styles,
        font,
        'sz',
        String(style.size),
        fontChildOrder,
      );
    if (colorChanged && style.color !== undefined) {
      setXlsxColorChild(this.styles, font, style.color, fontChildOrder);
      const color = directChild(font, 'color');
      if (color && semanticColor)
        applyXlsxSemanticColorOrigin(color, semanticColor);
    }
    if (boldChanged && style.bold !== undefined)
      setXlsxToggleChild(this.styles, font, 'b', style.bold, fontChildOrder);
    if (italicChanged && style.italic !== undefined)
      setXlsxToggleChild(this.styles, font, 'i', style.italic, fontChildOrder);
    if (strikeChanged && style.strike !== undefined)
      setXlsxToggleChild(
        this.styles,
        font,
        'strike',
        style.strike,
        fontChildOrder,
      );
    if (underlineChanged && style.underline !== undefined)
      setXlsxUnderlineChild(this.styles, font, style.underline, fontChildOrder);
    const index = fonts.length;
    this.fonts.append(font);
    this.generatedFonts.set(key, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private fillId(
    baseFillId: number,
    cell: Cell,
    colorOrigin?: XlsxSemanticColorOrigin,
    patternFill?: XlsxPatternFill,
  ): number {
    if (!patternFill && cell.bg === undefined) return baseFillId;
    const solidColor = patternFill ? null : xlsxRgbColor(cell.bg);
    if (!patternFill && !solidColor) return baseFillId;
    const fills = directChildren(this.fills, 'fill');
    const baseIndex = fills[baseFillId] ? baseFillId : 0;
    const basePattern = directChild(fills[baseIndex], 'patternFill');
    const foregroundColor = patternFill?.foregroundColor ?? solidColor!;
    const backgroundColor = patternFill?.backgroundColor;
    const patternType = patternFill?.patternType ?? 'solid';
    const foregroundOrigin = activeXlsxSemanticColorOrigin(
      patternFill?.foregroundColorOrigin ?? colorOrigin,
      foregroundColor,
      this.semanticPalette,
    );
    const backgroundOrigin = patternFill
      ? activeXlsxSemanticColorOrigin(
          patternFill.backgroundColorOrigin,
          patternFill.backgroundColor,
          this.semanticPalette,
        )
      : undefined;
    if (
      attribute(basePattern ?? fills[baseIndex], 'patternType') ===
        patternType &&
      (foregroundOrigin
        ? xlsxColorElementMatchesOrigin(
            directChild(basePattern ?? fills[baseIndex], 'fgColor'),
            foregroundOrigin,
          )
        : xlsxColorMatches(
            directChild(basePattern ?? fills[baseIndex], 'fgColor'),
            foregroundColor,
            this.colors,
            '#000000',
          )) &&
      (!patternFill ||
        (backgroundOrigin
          ? xlsxColorElementMatchesOrigin(
              directChild(basePattern ?? fills[baseIndex], 'bgColor'),
              backgroundOrigin,
            )
          : xlsxColorMatches(
              directChild(basePattern ?? fills[baseIndex], 'bgColor'),
              backgroundColor!,
              this.colors,
              '#ffffff',
            )))
    ) {
      return baseIndex;
    }
    const key = [
      patternType,
      foregroundColor,
      xlsxSemanticColorOriginKey(foregroundOrigin),
      backgroundColor ?? '',
      xlsxSemanticColorOriginKey(backgroundOrigin),
    ].join(':');
    const cached = this.generatedFills.get(key);
    if (cached !== undefined) return cached;
    const fill = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'fill',
    );
    const pattern = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      'patternFill',
    );
    pattern.setAttribute('patternType', patternType);
    const foreground = this.fillColorElement(
      'fgColor',
      foregroundColor,
      foregroundOrigin,
    );
    const background = patternFill
      ? this.fillColorElement(
          'bgColor',
          patternFill.backgroundColor,
          backgroundOrigin,
        )
      : this.styles.createElementNS(
          this.styles.documentElement.namespaceURI,
          'bgColor',
        );
    if (!patternFill) background.setAttribute('indexed', '64');
    pattern.append(foreground, background);
    fill.append(pattern);
    const index = directChildren(this.fills, 'fill').length;
    this.fills.append(fill);
    this.generatedFills.set(key, index);
    this.changed = true;
    this.updateCounts();
    return index;
  }

  private fillColorElement(
    name: 'bgColor' | 'fgColor',
    color: string,
    origin?: XlsxSemanticColorOrigin,
  ): Element {
    const element = this.styles.createElementNS(
      this.styles.documentElement.namespaceURI,
      name,
    );
    const rgb = xlsxRgbColor(color);
    if (rgb) element.setAttribute('rgb', rgb);
    if (origin) applyXlsxSemanticColorOrigin(element, origin);
    return element;
  }

  private borderId(
    baseBorderId: number,
    update: XlsxCellBorder | undefined,
    colorOrigins?: XlsxBorderColorOrigins,
  ): number {
    if (!update) return baseBorderId;
    const borders = directChildren(this.borders, 'border');
    const baseIndex = borders[baseBorderId] ? baseBorderId : 0;
    const base = borders[baseIndex];
    const lineUpdates = [
      ['left', update.left],
      ['right', update.right],
      ['top', update.top],
      ['bottom', update.bottom],
      ['diagonal', update.diagonal],
    ] as const;
    const lineChanged = lineUpdates.some(
      ([name, line]) =>
        line !== undefined &&
        !xlsxBorderLineMatches(
          directChild(base, name),
          line,
          this.colors,
          line
            ? activeXlsxSemanticColorOrigin(
                colorOrigins?.[name],
                line.color,
                this.semanticPalette,
              )
            : undefined,
        ),
    );
    const diagonalUpChanged =
      update.diagonalUp !== undefined &&
      xlsxBooleanAttribute(base, 'diagonalUp') !== update.diagonalUp;
    const diagonalDownChanged =
      update.diagonalDown !== undefined &&
      xlsxBooleanAttribute(base, 'diagonalDown') !== update.diagonalDown;
    if (!lineChanged && !diagonalUpChanged && !diagonalDownChanged) {
      return baseIndex;
    }
    const key = `${baseIndex}:${JSON.stringify(update)}:${JSON.stringify(colorOrigins ?? {})}`;
    const cached = this.generatedBorders.get(key);
    if (cached !== undefined) return cached;
    const border = base.cloneNode(true) as Element;
    for (const [name, line] of lineUpdates) {
      if (
        line !== undefined &&
        !xlsxBorderLineMatches(
          directChild(base, name),
          line,
          this.colors,
          line
            ? activeXlsxSemanticColorOrigin(
                colorOrigins?.[name],
                line.color,
                this.semanticPalette,
              )
            : undefined,
        )
      ) {
        setXlsxBorderLine(this.styles, border, name, line);
        const semanticColor = line
          ? activeXlsxSemanticColorOrigin(
              colorOrigins?.[name],
              line.color,
              this.semanticPalette,
            )
          : undefined;
        const color = directChild(directChild(border, name) ?? border, 'color');
        if (color && semanticColor)
          applyXlsxSemanticColorOrigin(color, semanticColor);
      }
    }
    if (diagonalUpChanged && update.diagonalUp !== undefined) {
      border.setAttribute('diagonalUp', update.diagonalUp ? '1' : '0');
    }
    if (diagonalDownChanged && update.diagonalDown !== undefined) {
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
      cell.rt !== undefined ||
      cell.tr !== undefined,
  );
}
