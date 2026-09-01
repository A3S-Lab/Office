import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { useState } from 'react';
import { SiteSidebar } from '../playground/src/site-sidebar';
import { WorkspaceHome } from '../playground/src/workspace-home';

test('keeps examples in the template grid without a latest-capabilities promotion', () => {
  const createdTemplates: string[] = [];

  render(
    <WorkspaceHome
      artifacts={[]}
      collaborationDocsUrl="/docs/components/collaboration-server.html"
      sidebarOpen
      onOpenSidebar={() => undefined}
      onCreate={(templateId) => createdTemplates.push(templateId)}
      onOpen={() => undefined}
      onImport={() => undefined}
      onOpenCollaborationDemo={() => undefined}
      onOpenSuggestionDemo={() => undefined}
      onOpenFormattingReviewDemo={() => undefined}
      onOpenPdf={() => undefined}
    />,
  );

  expect(screen.queryByRole('region', { name: '最新能力' })).toBeNull();

  const templates = screen.getByRole('region', { name: '新建' });
  const dataValidation = within(templates).getByRole('button', {
    name: '从模板新建：数据验证',
  });
  const structuredReferences = within(templates).getByRole('button', {
    name: '从模板新建：结构化引用',
  });
  const objectAnimations = within(templates).getByRole('button', {
    name: '从模板新建：进入与退出动画',
  });
  fireEvent.click(dataValidation);

  expect(createdTemplates).toEqual(['data-validation']);
  fireEvent.click(structuredReferences);
  expect(createdTemplates).toEqual([
    'data-validation',
    'structured-references',
  ]);
  fireEvent.click(objectAnimations);
  expect(createdTemplates).toEqual([
    'data-validation',
    'structured-references',
    'animated-deck',
  ]);
});

test('keeps Markdown last in the quick-create list', () => {
  const createdTemplates: string[] = [];

  render(
    <SiteSidebar
      docsUrl="/docs/"
      homeUrl="/"
      logoUrl="/a3s-logo.png"
      onCollapse={() => undefined}
      onHome={() => undefined}
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
  let homeRequests = 0;

  render(
    <SiteSidebar
      docsUrl="/docs/"
      homeUrl="/"
      logoUrl="/a3s-logo.png"
      onCollapse={() => undefined}
      onHome={() => {
        homeRequests += 1;
      }}
      onCreate={() => undefined}
      onOpenFile={() => undefined}
      onOpenPdf={() => undefined}
    />,
  );

  const productNavigation = screen.getByRole('navigation', {
    name: '产品页面',
  });
  const playground = within(productNavigation).getByRole('button', {
    name: 'Playground',
  });
  expect(playground).toHaveAttribute('aria-current', 'page');
  expect(
    within(productNavigation).getByRole('link', { name: '文档' }),
  ).toHaveAttribute('href', '/docs/');
  fireEvent.click(playground);
  expect(homeRequests).toBe(1);
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
  expect(screen.getByRole('button', { name: 'Playground' })).toHaveFocus();
  fireEvent.keyDown(screen.getByRole('button', { name: 'Playground' }), {
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
          docsUrl="/docs/"
          homeUrl="/"
          logoUrl="/a3s-logo.png"
          modal
          onCollapse={() => setOpen(false)}
          onHome={() => undefined}
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
