import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { Dialog } from '../src/internal/design-system/primitives';
import { officeOverlayPortalRoot } from '../src/internal/design-system/primitives/overlay/portal-root';
import { OfficeSelect } from '../src/internal/features/work/editors/office-select';

test('Dialog-hosted OfficeSelect portals into the dialog focus scope', () => {
  render(
    <Dialog title="Sort" onClose={() => undefined}>
      <OfficeSelect
        ariaLabel="Column"
        value="A"
        options={[
          { value: 'A', label: 'Column A' },
          { value: 'B', label: 'Column B' },
        ]}
        onChange={() => undefined}
      />
    </Dialog>,
  );

  fireEvent.click(screen.getByRole('combobox', { name: 'Column' }));
  const listbox = screen.getByRole('listbox');
  const dialog = screen.getByRole('dialog', { name: 'Sort' });
  expect(dialog.contains(listbox)).toBe(true);
  expect(listbox.closest('.ds-dialog')).toBe(dialog);
});

test('office overlay portal prefers the modal dialog focus root', () => {
  const host = document.createElement('div');
  document.body.append(host);
  host.innerHTML = `
    <dialog class="ds-dialog-backdrop" open>
      <section class="ds-dialog" role="dialog" aria-modal="true">
        <button type="button" id="anchor">Open</button>
      </section>
    </dialog>
  `;
  const anchor = host.querySelector('#anchor');
  expect(anchor).toBeTruthy();
  const root = officeOverlayPortalRoot(document, anchor as HTMLElement);
  expect(root.getAttribute('role')).toBe('dialog');
  expect(root.classList.contains('ds-dialog')).toBe(true);
  host.remove();
});

test('office overlay portal escapes a clipped office root onto document.body', () => {
  const host = document.createElement('div');
  host.setAttribute('data-a3s-office', '');
  host.style.overflow = 'hidden';
  const anchor = document.createElement('button');
  host.append(anchor);
  document.body.append(host);
  expect(officeOverlayPortalRoot(document, anchor)).toBe(document.body);
  host.remove();
});
