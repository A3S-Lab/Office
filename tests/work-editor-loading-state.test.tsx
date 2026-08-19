import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { WorkEditorLoadingState } from '../src/internal/features/work/components/work-editor-loading-state';

test('renders an accessible editor loading state with a decorative layout glyph', () => {
  render(<WorkEditorLoadingState title="正在加载文字编辑器" />);

  const status = screen.getByRole('status');
  expect(status).toHaveClass('work-editor-loading');
  expect(
    screen.getByRole('heading', { name: '正在加载文字编辑器' }),
  ).toBeVisible();
  expect(status.querySelector('.work-editor-loading-visual')).toHaveAttribute(
    'aria-hidden',
    'true',
  );
  expect(status.querySelectorAll('.work-editor-loading-line')).toHaveLength(3);
});

test('keeps the editor loading state centered and motion-safe', async () => {
  const styles = await readFile(
    path.resolve(import.meta.dirname, '../src/styles/work-editor.css'),
    'utf8',
  );

  expect(styles).toMatch(
    /\.work-editor-loading\.ds-state-view\s*\{[^}]*height:\s*100%[^}]*align-items:\s*center[^}]*justify-content:\s*center/s,
  );
  expect(styles).toMatch(
    /@media \(prefers-reduced-motion: reduce\)[\s\S]*?\.work-editor-loading-focus,[\s\S]*?animation:\s*none/,
  );
});
