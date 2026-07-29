import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { useState } from 'react';
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

test('keeps one documentation entry in the product navigation', () => {
  const routes: string[] = [];

  render(
    <SiteSidebar
      route="guide"
      onCollapse={() => undefined}
      onNavigate={(route) => routes.push(route)}
      onCreate={() => undefined}
      onOpenFile={() => undefined}
      onOpenPdf={() => undefined}
    />,
  );

  const productNavigation = screen.getByRole('navigation', {
    name: '产品页面',
  });
  const items = within(productNavigation)
    .getAllByRole('button')
    .map((button) => button.textContent?.trim());

  expect(items).toEqual(['编辑器', '接入文档']);
  fireEvent.click(
    within(productNavigation).getByRole('button', { name: '接入文档' }),
  );
  expect(routes).toEqual(['guide']);
});

test('contains phone focus in the sidebar and restores its exact trigger', () => {
  render(<ModalSidebarHarness />);

  const trigger = screen.getByRole('button', { name: '展开办公侧边栏' });
  trigger.focus();
  fireEvent.click(trigger);

  const sidebar = screen.getByRole('dialog', { name: 'A3S Office 导航' });
  const close = screen.getByRole('button', { name: '收起办公侧边栏' });
  const background = screen.getByRole('button', { name: '页面操作' });
  expect(sidebar).toHaveAttribute('aria-modal', 'true');
  expect(background.closest<HTMLElement>('[inert]')).toBeInTheDocument();
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Tab' });
  expect(screen.getByRole('button', { name: '编辑器' })).toHaveFocus();
  fireEvent.keyDown(screen.getByRole('button', { name: '编辑器' }), {
    key: 'Tab',
    shiftKey: true,
  });
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: 'A3S Office 导航' })).toBeNull();
  expect(trigger).toHaveFocus();
});

function ModalSidebarHarness() {
  const [open, setOpen] = useState(false);
  return (
    <div>
      <button type="button" onClick={() => setOpen(true)}>
        展开办公侧边栏
      </button>
      {open && (
        <SiteSidebar
          modal
          route="office"
          onCollapse={() => setOpen(false)}
          onNavigate={() => undefined}
          onCreate={() => undefined}
          onOpenFile={() => undefined}
          onOpenPdf={() => undefined}
        />
      )}
      <main>
        <button type="button">页面操作</button>
      </main>
    </div>
  );
}
