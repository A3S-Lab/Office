import { describe, expect, test } from '@rstest/core';
import {
  applyDocumentTableSharedBorderPaint,
  projectDocumentTableBordersForPaint,
  resolveDocumentTableSharedBorder,
  type DocumentTableBorder,
  type DocumentTableCellBorders,
} from '../src/internal/features/work/work-document-table-borders';

const none: DocumentTableBorder = {
  color: '#000000',
  style: 'none',
  width: 0,
};

const solid1: DocumentTableBorder = {
  color: '#c00000',
  style: 'solid',
  width: 1,
};

const solid3: DocumentTableBorder = {
  color: '#4472c4',
  style: 'solid',
  width: 3,
};

const double1: DocumentTableBorder = {
  color: '#70ad47',
  style: 'double',
  width: 1,
};

const dotted1: DocumentTableBorder = {
  color: '#c00000',
  style: 'dotted',
  width: 1,
};

describe('document table shared-edge border conflict', () => {
  test('none loses to any painted border', () => {
    expect(resolveDocumentTableSharedBorder(solid1, none)).toEqual(solid1);
    expect(resolveDocumentTableSharedBorder(none, solid1)).toEqual(solid1);
  });

  test('thicker border wins when styles match', () => {
    expect(resolveDocumentTableSharedBorder(solid1, solid3)).toEqual(solid3);
    expect(resolveDocumentTableSharedBorder(solid3, solid1)).toEqual(solid3);
  });

  test('higher style weight wins at equal width', () => {
    expect(resolveDocumentTableSharedBorder(solid1, double1)).toEqual(double1);
    expect(resolveDocumentTableSharedBorder(dotted1, solid1)).toEqual(solid1);
  });

  test('equal borders keep a stable winner (first argument)', () => {
    const twin: DocumentTableBorder = { ...solid1 };
    expect(resolveDocumentTableSharedBorder(solid1, twin)).toEqual(solid1);
  });

  test('projects winners onto both sides of shared edges for paint', () => {
    const left: DocumentTableCellBorders = {
      top: none,
      right: dotted1,
      bottom: none,
      left: none,
    };
    const right: DocumentTableCellBorders = {
      top: none,
      right: none,
      bottom: none,
      left: none,
    };
    const bottomLeft: DocumentTableCellBorders = {
      top: solid3,
      right: none,
      bottom: none,
      left: none,
    };
    const bottomRight: DocumentTableCellBorders = {
      top: solid1,
      right: none,
      bottom: none,
      left: none,
    };

    const projected = projectDocumentTableBordersForPaint([
      [left, right],
      [bottomLeft, bottomRight],
    ]);

    expect(projected[0]?.[0]?.right).toEqual(dotted1);
    expect(projected[0]?.[1]?.left).toEqual(dotted1);
    expect(projected[0]?.[0]?.bottom).toEqual(solid3);
    expect(projected[1]?.[0]?.top).toEqual(solid3);
    expect(projected[0]?.[1]?.bottom).toEqual(solid1);
    expect(projected[1]?.[1]?.top).toEqual(solid1);
  });

  test('projects shared-edge winners across a colspan cell for paint', () => {
    document.body.innerHTML = `
      <table>
        <tr>
          <td id="span" colspan="2" style="border-right: 3px solid #4472c4"
            data-office-cell-border-top-style="none"
            data-office-cell-border-right-style="solid"
            data-office-cell-border-right-width="3"
            data-office-cell-border-right-color="#4472c4"
            data-office-cell-border-bottom-style="none"
            data-office-cell-border-left-style="none">A</td>
          <td id="side"
            data-office-cell-border-top-style="none"
            data-office-cell-border-right-style="none"
            data-office-cell-border-bottom-style="none"
            data-office-cell-border-left-style="dotted"
            data-office-cell-border-left-width="1"
            data-office-cell-border-left-color="#c00000">B</td>
        </tr>
      </table>
    `;
    const table = document.querySelector('table');
    const span = document.getElementById('span');
    const side = document.getElementById('side');
    if (
      !(table instanceof HTMLTableElement) ||
      !(span instanceof HTMLTableCellElement) ||
      !(side instanceof HTMLTableCellElement)
    ) {
      throw new Error('Expected colspan fixture.');
    }
    applyDocumentTableSharedBorderPaint(table);
    expect(span.dataset.officeCellBorderRightStyle).toBe('solid');
    expect(side.dataset.officeCellBorderLeftStyle).toBe('dotted');
    expect(span.style.borderRightStyle).toBe('solid');
    expect(span.style.borderRightWidth).toContain('3');
    expect(side.style.borderLeftStyle).toBe('solid');
    expect(side.style.borderLeftWidth).toContain('3');
  });

  test('projects shared-edge winners across a rowspan cell for paint', () => {
    document.body.innerHTML = `
      <table>
        <tr>
          <td id="tall" rowspan="2"
            data-office-cell-border-top-style="none"
            data-office-cell-border-right-style="none"
            data-office-cell-border-bottom-style="double"
            data-office-cell-border-bottom-width="1"
            data-office-cell-border-bottom-color="#70ad47"
            data-office-cell-border-left-style="none">A</td>
          <td>B</td>
        </tr>
        <tr>
          <td>C</td>
        </tr>
        <tr>
          <td id="below"
            data-office-cell-border-top-style="solid"
            data-office-cell-border-top-width="1"
            data-office-cell-border-top-color="#c00000"
            data-office-cell-border-right-style="none"
            data-office-cell-border-bottom-style="none"
            data-office-cell-border-left-style="none">D</td>
          <td>E</td>
        </tr>
      </table>
    `;
    const table = document.querySelector('table');
    const tall = document.getElementById('tall');
    const below = document.getElementById('below');
    if (
      !(table instanceof HTMLTableElement) ||
      !(tall instanceof HTMLTableCellElement) ||
      !(below instanceof HTMLTableCellElement)
    ) {
      throw new Error('Expected rowspan fixture.');
    }
    applyDocumentTableSharedBorderPaint(table);
    expect(tall.dataset.officeCellBorderBottomStyle).toBe('double');
    expect(below.dataset.officeCellBorderTopStyle).toBe('solid');
    expect(tall.style.borderBottomStyle).toBe('double');
    expect(below.style.borderTopStyle).toBe('double');
  });
});
