import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { SiteSidebar } from '../playground/src/site-sidebar';

test('keeps Markdown last in the quick-create list', () => {
  const createdTemplates: string[] = [];

  render(
    <SiteSidebar
      route="office"
      onCollapse={() => undefined}
      onNavigate={() => undefined}
      onCreate={(templateId) => createdTemplates.push(templateId)}
      onOpenFile={() => undefined}
      onOpenPdf={() => undefined}
    />,
  );

  const quickCreate = screen.getByRole('region', { name: '快速新建' });
  const labels = within(quickCreate)
    .getAllByRole('button')
    .map((button) => button.textContent?.replaceAll(/\s/g, ''));

  expect(labels).toEqual(['文字', '表格', '演示', 'PDF打开', 'Markdown']);

  fireEvent.click(
    within(quickCreate).getByRole('button', { name: 'Markdown' }),
  );
  expect(createdTemplates).toEqual(['blank-markdown']);
});
