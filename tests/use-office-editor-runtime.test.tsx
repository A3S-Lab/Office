import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { useMemo } from 'react';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from '../src/internal/features/work/editors/office-editor-extension';
import { useOfficeEditorRuntime } from '../src/internal/features/work/editors/use-office-editor-runtime';

interface ToggleContext {
  enabled: boolean;
}

interface ToggleCommands {
  run: () => void;
}

test('renders command availability from the current controlled context', () => {
  const view = render(<RuntimeHarness enabled={false} />);
  expect(screen.getByRole('button', { name: '运行' })).toBeDisabled();

  view.rerender(<RuntimeHarness enabled />);

  expect(screen.getByRole('button', { name: '运行' })).toBeEnabled();
});

function RuntimeHarness({ enabled }: ToggleContext) {
  const extensions = useMemo<
    readonly OfficeEditorExtension<ToggleContext, ToggleCommands>[]
  >(
    () => [
      createOfficeEditorExtension<ToggleContext, ToggleCommands>({
        name: 'toggle',
        addCommands: () => ({
          run: {
            canExecute: (context) => context.enabled,
            execute: () => undefined,
          },
        }),
      }),
    ],
    [],
  );
  const runtime = useOfficeEditorRuntime({ enabled }, extensions);

  return (
    <button type="button" disabled={!runtime.can().run()}>
      运行
    </button>
  );
}
