import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import {
  DocumentRuler,
  documentRulerPageHeight,
  documentRulerMaximumIndentLevel,
  documentRulerPageWidth,
} from '../src/internal/features/work/editors/document-ruler';
import { DocumentVerticalRuler } from '../src/internal/features/work/editors/document-vertical-ruler';
import type { DocumentParagraphIndent } from '../src/internal/features/work/work-document-paragraph-formatting';
import type { DocumentTabStop } from '../src/internal/features/work/work-document-tab-stops';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

const layout: WorkDocumentSectionLayout = {
  pageSize: 'a4',
  orientation: 'portrait',
  margins: { top: 25, right: 23, bottom: 25, left: 23 },
  columns: { count: 1, spacing: 12, separator: false },
  breakAfter: 'nextPage',
};

test('matches the active page orientation and available indent width', () => {
  expect(documentRulerPageWidth(layout)).toBeCloseTo(793.7, 1);
  expect(documentRulerPageHeight(layout)).toBeCloseTo(1122.5, 1);
  expect(
    documentRulerPageWidth({ ...layout, orientation: 'landscape' }),
  ).toBeCloseTo(1122.5, 1);
  expect(
    documentRulerPageHeight({ ...layout, orientation: 'landscape' }),
  ).toBeCloseTo(793.7, 1);
  expect(documentRulerMaximumIndentLevel(layout)).toBe(8);
});

test('updates margins and paragraph indent from accessible ruler handles', () => {
  const layoutChanges: WorkDocumentSectionLayout[] = [];
  const indentChanges: DocumentParagraphIndent[] = [];

  render(
    <DocumentRuler
      paragraphIndent={{ left: 48, right: 12, firstLine: -12 }}
      tabStops={[]}
      layout={layout}
      onParagraphIndentChange={(value) => indentChanges.push(value)}
      onTabStopsChange={() => undefined}
      onLayoutChange={(value) => layoutChanges.push(value)}
    />,
  );

  fireEvent.keyDown(screen.getByRole('slider', { name: '左页边距' }), {
    key: 'ArrowRight',
  });
  fireEvent.keyDown(screen.getByRole('slider', { name: '左缩进' }), {
    key: 'ArrowLeft',
  });
  fireEvent.keyDown(screen.getByRole('slider', { name: '首行/悬挂缩进' }), {
    key: 'ArrowRight',
  });
  fireEvent.keyDown(screen.getByRole('slider', { name: '右缩进' }), {
    key: 'ArrowRight',
  });

  expect(layoutChanges.at(-1)?.margins.left).toBe(24);
  expect(indentChanges).toEqual([
    { left: 42, right: 12, firstLine: -12 },
    { left: 48, right: 12, firstLine: -6 },
    { left: 48, right: 18, firstLine: -12 },
  ]);
  expect(screen.getByRole('slider', { name: '右页边距' })).toHaveAttribute(
    'aria-valuetext',
    '23 毫米',
  );
  expect(screen.getByRole('slider', { name: '首行/悬挂缩进' })).toHaveAttribute(
    'aria-valuetext',
    '悬挂缩进 0.32 厘米',
  );
});

test('adds, moves, changes, and removes accessible paragraph tab stops', () => {
  const tabStopChanges: DocumentTabStop[][] = [];
  const { rerender } = render(
    <DocumentRuler
      paragraphIndent={{ left: 0, right: 0, firstLine: 0 }}
      tabStops={[]}
      layout={layout}
      onParagraphIndentChange={() => undefined}
      onTabStopsChange={(value) => tabStopChanges.push(value)}
      onLayoutChange={() => undefined}
    />,
  );
  const ruler = screen.getByRole('group', { name: '水平标尺' });
  ruler.getBoundingClientRect = () =>
    ({
      bottom: 20,
      height: 20,
      left: 0,
      right: documentRulerPageWidth(layout),
      top: 0,
      width: documentRulerPageWidth(layout),
      x: 0,
      y: 0,
      toJSON: () => ({}),
    }) as DOMRect;

  fireEvent.click(screen.getByRole('button', { name: '在标尺上添加制表位' }), {
    clientX: 210,
  });
  expect(tabStopChanges.at(-1)).toEqual([
    { position: 123, alignment: 'left', leader: 'none' },
  ]);

  const tabStops: DocumentTabStop[] = [
    { position: 96, alignment: 'right', leader: 'dot' },
  ];
  rerender(
    <DocumentRuler
      paragraphIndent={{ left: 0, right: 0, firstLine: 0 }}
      tabStops={tabStops}
      layout={layout}
      onParagraphIndentChange={() => undefined}
      onTabStopsChange={(value) => tabStopChanges.push(value)}
      onLayoutChange={() => undefined}
    />,
  );
  const handle = screen.getByRole('slider', { name: '右对齐制表位 1' });
  expect(handle).toHaveAttribute('aria-valuetext', '2.54 厘米，点线前导符');

  fireEvent.keyDown(handle, { key: 'ArrowRight' });
  expect(tabStopChanges.at(-1)).toEqual([
    { position: 102, alignment: 'right', leader: 'dot' },
  ]);

  fireEvent.keyDown(handle, { key: 'Enter' });
  expect(tabStopChanges.at(-1)).toEqual([
    { position: 96, alignment: 'decimal', leader: 'dot' },
  ]);

  fireEvent.keyDown(handle, { key: 'Delete' });
  expect(tabStopChanges.at(-1)).toEqual([]);
});

test('updates vertical page margins with keyboard and pointer input', () => {
  const layoutChanges: WorkDocumentSectionLayout[] = [];

  render(
    <DocumentVerticalRuler
      layout={layout}
      onLayoutChange={(value) => layoutChanges.push(value)}
    />,
  );

  const top = screen.getByRole('slider', { name: '上页边距' });
  const bottom = screen.getByRole('slider', { name: '下页边距' });
  fireEvent.keyDown(top, { key: 'ArrowUp' });
  fireEvent.keyDown(bottom, { key: 'PageDown' });

  expect(layoutChanges[0]?.margins.top).toBe(26);
  expect(layoutChanges[1]?.margins.bottom).toBe(20);
  expect(top).toHaveAttribute('aria-orientation', 'vertical');
  expect(bottom).toHaveAttribute('aria-valuetext', '25 毫米');

  const ruler = screen.getByRole('group', { name: '垂直标尺' });
  const pageHeight = documentRulerPageHeight(layout);
  ruler.getBoundingClientRect = () =>
    ({
      bottom: pageHeight,
      height: pageHeight,
      left: 0,
      right: 20,
      top: 0,
      width: 20,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    }) as DOMRect;
  fireEvent.pointerDown(top, {
    button: 0,
    clientY: (30 * 96) / 25.4,
    pointerId: 7,
  });

  expect(layoutChanges.at(-1)?.margins.top).toBe(30);
});
