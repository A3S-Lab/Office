import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  isWorkspaceContextMenuKeyboardEvent,
  WorkspaceContextMenu,
  workspaceContextMenuPosition,
} from '../src/internal/features/workspace/components/workspace-context-menu';

test('keeps context-menu focus deterministic and exits in tab order', async () => {
  function Fixture() {
    const [open, setOpen] = useState(false);
    return (
      <section data-a3s-office>
        <button type="button">Before menu</button>
        <button type="button" onClick={() => setOpen(true)}>
          Open menu
        </button>
        <button type="button">After menu</button>
        {open && (
          <WorkspaceContextMenu
            label="Selection actions"
            x={24}
            y={24}
            items={[
              {
                id: 'disabled',
                label: 'Unavailable action',
                icon: <span />,
                disabled: true,
                onSelect: () => undefined,
              },
              {
                id: 'copy',
                label: 'Copy',
                icon: <span />,
                onSelect: () => undefined,
              },
              {
                id: 'delete',
                label: 'Delete',
                icon: <span />,
                onSelect: () => undefined,
              },
            ]}
            onClose={() => setOpen(false)}
          />
        )}
      </section>
    );
  }

  render(<Fixture />);
  const trigger = screen.getByRole('button', { name: 'Open menu' });
  trigger.focus();
  fireEvent.click(trigger);

  let menu = screen.getByRole('menu', { name: 'Selection actions' });
  expect(menu).toHaveClass('work-office-context-menu');
  expect(menu).toHaveAttribute('data-office-shortcuts', 'ignore');
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );
  expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveAttribute(
    'tabindex',
    '-1',
  );

  fireEvent.keyDown(menu, { key: 'End' });
  expect(screen.getByRole('menuitem', { name: 'Delete' })).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Escape' });
  await waitFor(() => {
    expect(screen.queryByRole('menu')).toBeNull();
    expect(trigger).toHaveFocus();
  });

  trigger.focus();
  fireEvent.click(trigger);
  menu = screen.getByRole('menu', { name: 'Selection actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );
  fireEvent.keyDown(menu, { key: 'Tab' });
  await waitFor(() =>
    expect(screen.getByRole('button', { name: 'After menu' })).toHaveFocus(),
  );

  trigger.focus();
  fireEvent.click(trigger);
  menu = screen.getByRole('menu', { name: 'Selection actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );
  fireEvent.keyDown(menu, { key: 'Tab', shiftKey: true });
  await waitFor(() =>
    expect(screen.getByRole('button', { name: 'Before menu' })).toHaveFocus(),
  );
});

test('anchors keyboard context menus to the focused control', () => {
  let keyboardPosition: { x: number; y: number } | null = null;
  let pointerPosition: { x: number; y: number } | null = null;
  render(
    <button
      type="button"
      onKeyDown={(event) => {
        if (isWorkspaceContextMenuKeyboardEvent(event)) {
          keyboardPosition = workspaceContextMenuPosition(event);
        }
      }}
      onContextMenu={(event) => {
        pointerPosition = workspaceContextMenuPosition(event);
      }}
    >
      Context target
    </button>,
  );

  const target = screen.getByRole('button', { name: 'Context target' });
  target.getBoundingClientRect = () =>
    ({
      bottom: 100,
      height: 20,
      left: 120,
      right: 160,
      top: 80,
      width: 40,
      x: 120,
      y: 80,
      toJSON: () => undefined,
    }) as DOMRect;

  fireEvent.keyDown(target, { key: 'F10', shiftKey: true });
  expect(keyboardPosition).toEqual({ x: 140, y: 90 });
  fireEvent.contextMenu(target, { clientX: 44, clientY: 66 });
  expect(pointerPosition).toEqual({ x: 44, y: 66 });
});

test('prefers an editor-provided keyboard anchor without moving pointer menus', () => {
  let keyboardPosition: { x: number; y: number } | null = null;
  let pointerPosition: { x: number; y: number } | null = null;
  render(
    <button
      type="button"
      onKeyDown={(event) => {
        if (isWorkspaceContextMenuKeyboardEvent(event)) {
          keyboardPosition = workspaceContextMenuPosition(event, {
            height: 24,
            left: 40,
            top: 60,
            width: 80,
          });
        }
      }}
      onContextMenu={(event) => {
        pointerPosition = workspaceContextMenuPosition(event, {
          height: 24,
          left: 40,
          top: 60,
          width: 80,
        });
      }}
    >
      Anchored target
    </button>,
  );

  const target = screen.getByRole('button', { name: 'Anchored target' });
  fireEvent.keyDown(target, { key: 'F10', shiftKey: true });
  expect(keyboardPosition).toEqual({ x: 80, y: 72 });
  fireEvent.contextMenu(target, { clientX: 18, clientY: 28 });
  expect(pointerPosition).toEqual({ x: 18, y: 28 });
});

test('executes advertised shortcuts while the context menu owns focus', async () => {
  const calls: string[] = [];

  function Fixture() {
    const [open, setOpen] = useState(false);
    return (
      <section data-a3s-office>
        <button type="button" onClick={() => setOpen(true)}>
          Open shortcut menu
        </button>
        {open && (
          <WorkspaceContextMenu
            label="Shortcut actions"
            x={24}
            y={24}
            items={[
              {
                id: 'copy',
                label: 'Copy',
                icon: <span />,
                shortcut: '⌘C',
                ariaKeyShortcut: 'Control+C Meta+C',
                onSelect: () => calls.push('copy'),
              },
              {
                id: 'clear',
                label: 'Clear',
                icon: <span />,
                shortcut: 'Delete',
                ariaKeyShortcut: 'Delete',
                onSelect: () => calls.push('clear'),
              },
            ]}
            onClose={() => setOpen(false)}
          />
        )}
      </section>
    );
  }

  render(<Fixture />);
  const trigger = screen.getByRole('button', {
    name: 'Open shortcut menu',
  });
  trigger.focus();
  fireEvent.click(trigger);
  let menu = screen.getByRole('menu', { name: 'Shortcut actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );

  fireEvent.keyDown(menu, { key: 'c', metaKey: true });
  await waitFor(() => {
    expect(calls).toEqual(['copy']);
    expect(screen.queryByRole('menu', { name: 'Shortcut actions' })).toBeNull();
    expect(trigger).toHaveFocus();
  });

  fireEvent.click(trigger);
  menu = screen.getByRole('menu', { name: 'Shortcut actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );
  fireEvent.keyDown(menu, { key: 'c', ctrlKey: true });
  await waitFor(() => {
    expect(calls).toEqual(['copy', 'copy']);
    expect(screen.queryByRole('menu', { name: 'Shortcut actions' })).toBeNull();
    expect(trigger).toHaveFocus();
  });

  fireEvent.click(trigger);
  menu = screen.getByRole('menu', { name: 'Shortcut actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Copy' })).toHaveFocus(),
  );
  fireEvent.keyDown(menu, { key: 'Delete' });
  await waitFor(() => {
    expect(calls).toEqual(['copy', 'copy', 'clear']);
    expect(screen.queryByRole('menu', { name: 'Shortcut actions' })).toBeNull();
    expect(trigger).toHaveFocus();
  });
});
