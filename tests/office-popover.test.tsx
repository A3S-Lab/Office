import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { Popover } from '../src/internal/design-system/primitives';

test('carries the editor control accent into a portal panel', () => {
  render(
    <>
      <style>{`.test-office-popover-accent { --work-office-control-accent: #159469; }`}</style>
      <Popover
        label="Choose a font"
        panelLabel="Fonts"
        panelRole="listbox"
        portal
        className="test-office-popover-accent"
        trigger={(triggerProps) => (
          <button {...triggerProps}>Choose a font</button>
        )}
      >
        <button type="button" role="option" aria-selected="true">
          Aptos
        </button>
      </Popover>
    </>,
  );

  fireEvent.click(screen.getByRole('button', { name: 'Choose a font' }));

  expect(
    screen
      .getByRole('listbox', { name: 'Fonts' })
      .style.getPropertyValue('--work-office-control-accent'),
  ).toBe('#159469');
});

test('moves portal menu focus to the tab stop beside its trigger', async () => {
  render(
    <>
      <button type="button">Before menu</button>
      <Popover
        label="More actions"
        panelLabel="More actions"
        panelRole="menu"
        portal
        focusFirstOnOpen
        trigger={(triggerProps) => (
          <button {...triggerProps}>More actions</button>
        )}
      >
        <button type="button" role="menuitem" tabIndex={-1}>
          Action
        </button>
      </Popover>
      <span style={{ display: 'none' }}>
        <button type="button">Hidden after menu</button>
      </span>
      <button type="button">After menu</button>
    </>,
  );

  const trigger = screen.getByRole('button', { name: 'More actions' });
  fireEvent.click(trigger);
  const menu = screen.getByRole('menu', { name: 'More actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Action' })).toHaveFocus(),
  );

  fireEvent.keyDown(menu, { key: 'Tab' });
  await waitFor(() => {
    expect(screen.queryByRole('menu', { name: 'More actions' })).toBeNull();
    expect(screen.getByRole('button', { name: 'After menu' })).toHaveFocus();
  });

  fireEvent.click(trigger);
  const reopenedMenu = screen.getByRole('menu', { name: 'More actions' });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: 'Action' })).toHaveFocus(),
  );
  fireEvent.keyDown(reopenedMenu, { key: 'Tab', shiftKey: true });

  await waitFor(() =>
    expect(screen.getByRole('button', { name: 'Before menu' })).toHaveFocus(),
  );
});

test('keeps Tab inside a portal dialog until its focus boundary', async () => {
  render(
    <>
      <button type="button">Before settings</button>
      <Popover
        label="Open settings"
        panelLabel="Settings"
        panelRole="dialog"
        portal
        focusFirstOnOpen
        trigger={(triggerProps) => (
          <button {...triggerProps}>Open settings</button>
        )}
      >
        <input aria-label="First setting" />
        <input aria-label="Last setting" />
      </Popover>
      <button type="button">After settings</button>
    </>,
  );

  const trigger = screen.getByRole('button', { name: 'Open settings' });
  fireEvent.click(trigger);
  const dialog = screen.getByRole('dialog', { name: 'Settings' });
  const first = screen.getByRole('textbox', { name: 'First setting' });
  const last = screen.getByRole('textbox', { name: 'Last setting' });
  await waitFor(() => expect(first).toHaveFocus());

  fireEvent.keyDown(dialog, { key: 'Tab' });
  expect(dialog).toBeInTheDocument();
  last.focus();
  fireEvent.keyDown(dialog, { key: 'Tab' });
  await waitFor(() => {
    expect(screen.queryByRole('dialog', { name: 'Settings' })).toBeNull();
    expect(
      screen.getByRole('button', { name: 'After settings' }),
    ).toHaveFocus();
  });

  fireEvent.click(trigger);
  const reopenedDialog = screen.getByRole('dialog', { name: 'Settings' });
  await waitFor(() =>
    expect(
      screen.getByRole('textbox', { name: 'First setting' }),
    ).toHaveFocus(),
  );
  fireEvent.keyDown(reopenedDialog, { key: 'Tab', shiftKey: true });
  await waitFor(() =>
    expect(
      screen.getByRole('button', { name: 'Before settings' }),
    ).toHaveFocus(),
  );
});

test('treats each named radio group as one Tab stop in a portal dialog', async () => {
  render(
    <>
      <Popover
        label="Open styles"
        panelLabel="Styles"
        panelRole="dialog"
        portal
        focusFirstOnOpen
        trigger={(triggerProps) => (
          <button {...triggerProps}>Open styles</button>
        )}
      >
        <fieldset>
          <legend>Opacity</legend>
          <label>
            <input type="radio" name="opacity" aria-label="50%" />
            50%
          </label>
          <label>
            <input
              type="radio"
              name="opacity"
              aria-label="100%"
              defaultChecked
            />
            100%
          </label>
        </fieldset>
        <fieldset>
          <legend>Width</legend>
          <label>
            <input type="radio" name="width" aria-label="2 px" />2 px
          </label>
          <label>
            <input type="radio" name="width" aria-label="4 px" defaultChecked />
            4 px
          </label>
        </fieldset>
      </Popover>
      <button type="button">After styles</button>
    </>,
  );

  fireEvent.click(screen.getByRole('button', { name: 'Open styles' }));
  const dialog = screen.getByRole('dialog', { name: 'Styles' });
  const currentWidth = screen.getByRole('radio', { name: '4 px' });
  await waitFor(() =>
    expect(screen.getByRole('radio', { name: '100%' })).toHaveFocus(),
  );

  currentWidth.focus();
  fireEvent.keyDown(dialog, { key: 'Tab' });

  await waitFor(() => {
    expect(screen.queryByRole('dialog', { name: 'Styles' })).toBeNull();
    expect(screen.getByRole('button', { name: 'After styles' })).toHaveFocus();
  });
});
