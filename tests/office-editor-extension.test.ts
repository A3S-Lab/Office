import { describe, expect, test } from '@rstest/core';
import {
  createOfficeEditorExtension,
  createOfficeEditorRuntime,
} from '../src/internal/features/work/editors/office-editor-extension';

interface CounterContext {
  enabled: boolean;
  value: number;
}

interface CounterCommands {
  add: (amount: number) => number;
  reset: () => number;
}

describe('office editor extensions', () => {
  test('composes typed commands from ordered extensions', () => {
    const lifecycle: string[] = [];
    const counter = createOfficeEditorExtension<
      CounterContext,
      CounterCommands,
      { executions: number }
    >({
      name: 'counter',
      priority: 200,
      addStorage: () => ({ executions: 0 }),
      addCommands: ({ storage }) => ({
        add: {
          canExecute: (context) => context.enabled,
          execute: (context, amount) => {
            storage.executions += 1;
            context.value += amount;
            return context.value;
          },
        },
      }),
      onCreate: () => lifecycle.push('counter:create'),
      onDestroy: () => lifecycle.push('counter:destroy'),
    });
    const reset = createOfficeEditorExtension<CounterContext, CounterCommands>({
      name: 'reset',
      priority: 100,
      addCommands: () => ({
        reset: {
          execute: (context) => {
            context.value = 0;
            return context.value;
          },
        },
      }),
      onCreate: () => lifecycle.push('reset:create'),
      onDestroy: () => lifecycle.push('reset:destroy'),
    });
    const context = { enabled: true, value: 2 };
    const editor = createOfficeEditorRuntime(context, [reset, counter]);

    editor.mount();

    expect(editor.extensionNames).toEqual(['counter', 'reset']);
    expect(editor.can().add(3)).toBe(true);
    expect(editor.commands.add(3)).toBe(5);
    expect(editor.commands.reset()).toBe(0);
    expect(editor.storage.counter).toEqual({ executions: 1 });
    expect(lifecycle).toEqual(['counter:create', 'reset:create']);

    editor.unmount();
    expect(lifecycle).toEqual([
      'counter:create',
      'reset:create',
      'reset:destroy',
      'counter:destroy',
    ]);
  });

  test('commands read the latest controlled editor context', () => {
    const extension = createOfficeEditorExtension<
      CounterContext,
      CounterCommands
    >({
      name: 'counter',
      addCommands: () => ({
        add: {
          canExecute: (context) => context.enabled,
          execute: (context, amount) => {
            context.value += amount;
            return context.value;
          },
        },
        reset: {
          execute: (context) => {
            context.value = 0;
            return context.value;
          },
        },
      }),
    });
    const editor = createOfficeEditorRuntime({ enabled: false, value: 1 }, [
      extension,
    ]);

    expect(editor.can().add(2)).toBe(false);
    editor.updateContext({ enabled: true, value: 10 });
    expect(editor.can().add(2)).toBe(true);
    expect(editor.commands.add(2)).toBe(12);
  });

  test('rejects duplicate extension and command names', () => {
    const first = createOfficeEditorExtension<CounterContext, CounterCommands>({
      name: 'first',
      addCommands: () => ({
        reset: {
          execute: () => 0,
        },
      }),
    });
    const duplicateCommand = createOfficeEditorExtension<
      CounterContext,
      CounterCommands
    >({
      name: 'second',
      addCommands: () => ({
        reset: {
          execute: () => 1,
        },
      }),
    });

    expect(() =>
      createOfficeEditorRuntime({ enabled: true, value: 0 }, [first, first]),
    ).toThrow(/Duplicate editor extension "first"/);
    expect(() =>
      createOfficeEditorRuntime({ enabled: true, value: 0 }, [
        first,
        duplicateCommand,
      ]),
    ).toThrow(/Duplicate editor command "reset"/);
  });

  test('routes TipTap-style keyboard shortcuts through commands and can()', () => {
    const extension = createOfficeEditorExtension<
      CounterContext,
      CounterCommands
    >({
      name: 'counterKeyboard',
      addCommands: () => ({
        add: {
          canExecute: (context) => context.enabled,
          execute: (context, amount) => {
            context.value += amount;
            return context.value;
          },
        },
        reset: {
          execute: (context) => {
            context.value = 0;
            return context.value;
          },
        },
      }),
      addKeyboardShortcuts: () => ({
        'Mod-k': ({ can, commands }) =>
          can.add(2) ? commands.add(2) > 0 : false,
      }),
    });
    const context = { enabled: true, value: 1 };
    const editor = createOfficeEditorRuntime(context, [extension]);
    const handled = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'k',
      metaKey: true,
    });

    expect(editor.handleKeyDown(handled)).toBe(true);
    expect(handled.defaultPrevented).toBe(true);
    expect(context.value).toBe(3);

    editor.updateContext({ enabled: false, value: 5 });
    const blocked = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: 'k',
    });
    expect(editor.handleKeyDown(blocked)).toBe(false);
    expect(blocked.defaultPrevented).toBe(false);
  });
});
