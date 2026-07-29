import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { WorkOfficeRibbon } from '../src/internal/features/work/editors/work-office-chrome';

test('supports complete keyboard navigation in the shared file menu', async () => {
  render(
    <>
      <button type="button">Before ribbon</button>
      <WorkOfficeRibbon
        ariaLabel="Test ribbon"
        tabs={[
          { id: 'home', label: '开始' },
          { id: 'insert', label: '插入' },
        ]}
        defaultTab="home"
        panels={{
          home: <span>Home tools</span>,
          insert: <span>Insert tools</span>,
        }}
        fileActions={[
          { id: 'open', label: '打开', onSelect: () => undefined },
          { id: 'save', label: '保存', onSelect: () => undefined },
        ]}
      />
      <button type="button">After ribbon</button>
    </>,
  );

  const trigger = screen.getByRole('button', { name: '文件' });
  fireEvent.keyDown(trigger, { key: 'ArrowDown' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: '打开' })).toHaveFocus(),
  );

  const menu = screen.getByRole('menu', { name: '文件菜单' });
  fireEvent.keyDown(menu, { key: 'End' });
  expect(screen.getByRole('menuitem', { name: '保存' })).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Home' });
  expect(screen.getByRole('menuitem', { name: '打开' })).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Tab' });
  await waitFor(() => {
    expect(
      screen.queryByRole('menu', { name: '文件菜单' }),
    ).not.toBeInTheDocument();
    expect(screen.getByRole('tab', { name: '开始' })).toHaveFocus();
  });

  fireEvent.keyDown(trigger, { key: 'ArrowUp' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: '保存' })).toHaveFocus(),
  );
  fireEvent.keyDown(screen.getByRole('menu', { name: '文件菜单' }), {
    key: 'Tab',
    shiftKey: true,
  });
  await waitFor(() =>
    expect(screen.getByRole('button', { name: 'Before ribbon' })).toHaveFocus(),
  );

  fireEvent.keyDown(trigger, { key: 'ArrowDown' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: '打开' })).toHaveFocus(),
  );
  fireEvent.keyDown(screen.getByRole('menu', { name: '文件菜单' }), {
    key: 'Escape',
  });
  await waitFor(() => {
    expect(
      screen.queryByRole('menu', { name: '文件菜单' }),
    ).not.toBeInTheDocument();
    expect(trigger).toHaveFocus();
  });
});

test('pages compact ribbon tabs instead of compressing their labels', async () => {
  render(
    <WorkOfficeRibbon
      ariaLabel="Test ribbon"
      tabs={[
        { id: 'home', label: '开始' },
        { id: 'insert', label: '插入' },
        { id: 'design', label: '设计' },
        { id: 'transitions', label: '切换' },
        {
          id: 'slideshow',
          label: '幻灯片放映',
          compactLabel: '放映',
        },
      ]}
      defaultTab="home"
      panels={{
        home: <span>Home tools</span>,
        insert: <span>Insert tools</span>,
        design: <span>Design tools</span>,
        transitions: <span>Transition tools</span>,
        slideshow: <span>Slideshow tools</span>,
      }}
    />,
  );

  const tabList = screen.getByRole('tablist', { name: 'Test ribbon' });
  const slideshow = screen.getByRole('tab', { name: '幻灯片放映' });
  expect(slideshow.querySelector('.ds-tabs-label')).toHaveTextContent(
    '幻灯片放映',
  );
  expect(slideshow.querySelector('.ds-tabs-label-compact')).toHaveTextContent(
    '放映',
  );
  Object.defineProperty(tabList, 'clientWidth', {
    configurable: true,
    value: 160,
  });
  screen.getAllByRole('tab').forEach((tab, index) => {
    Object.defineProperties(tab, {
      offsetLeft: { configurable: true, value: index * 72 },
      offsetWidth: { configurable: true, value: 72 },
    });
  });

  fireEvent(window, new Event('resize'));
  const next = await waitFor(() =>
    screen.getByRole('button', { name: '向右查看更多功能区标签' }),
  );
  fireEvent.click(next);
  expect(tabList.scrollLeft).toBeGreaterThan(0);
  await waitFor(() =>
    expect(
      screen.getByRole('button', { name: '向左查看更多功能区标签' }),
    ).toBeInTheDocument(),
  );
});
