import { matchesOfficeEditorKeyboardShortcut } from '../../../keyboard-shortcuts';
import { isOfficeCompositionKeyboardEvent } from './office-shortcuts';

type CommandArguments<Command> = Command extends (
  ...args: infer Arguments extends unknown[]
) => unknown
  ? Arguments
  : never;

type CommandResult<Command> = Command extends (...args: never[]) => infer Result
  ? Result
  : never;

export interface OfficeEditorCommandDefinition<Context, Command> {
  canExecute?: (
    context: Context,
    ...args: CommandArguments<Command>
  ) => boolean;
  execute: (
    context: Context,
    ...args: CommandArguments<Command>
  ) => CommandResult<Command>;
}

export type OfficeEditorCommandDefinitions<Context, Commands> = {
  [Name in keyof Commands]?: OfficeEditorCommandDefinition<
    Context,
    Commands[Name]
  >;
};

export type OfficeEditorCanCommands<Commands> = {
  [Name in keyof Commands]: (
    ...args: CommandArguments<Commands[Name]>
  ) => boolean;
};

export interface OfficeEditorExtensionLifecycle<Context, Storage> {
  context: Context;
  storage: Storage;
}

export interface OfficeEditorExtensionUpdate<Context, Storage>
  extends OfficeEditorExtensionLifecycle<Context, Storage> {
  previousContext: Context;
}

export interface OfficeEditorKeyboardShortcutProps<Context, Commands, Storage> {
  can: OfficeEditorCanCommands<Commands>;
  commands: Commands;
  context: Context;
  storage: Storage;
}

export type OfficeEditorKeyboardShortcutHandler<Context, Commands, Storage> = (
  props: OfficeEditorKeyboardShortcutProps<Context, Commands, Storage>,
  event: KeyboardEvent,
) => boolean;

export interface OfficeEditorExtensionConfig<
  Context,
  Commands,
  Storage = Record<string, never>,
> {
  name: string;
  priority?: number;
  addStorage?: () => Storage;
  addCommands?: (extension: {
    storage: Storage;
  }) => OfficeEditorCommandDefinitions<Context, Commands>;
  addKeyboardShortcuts?: (extension: {
    storage: Storage;
  }) => Record<
    string,
    OfficeEditorKeyboardShortcutHandler<Context, Commands, Storage>
  >;
  onCreate?: (
    extension: OfficeEditorExtensionLifecycle<Context, Storage>,
  ) => void;
  onUpdate?: (extension: OfficeEditorExtensionUpdate<Context, Storage>) => void;
  onDestroy?: (
    extension: OfficeEditorExtensionLifecycle<Context, Storage>,
  ) => void;
}

interface ErasedOfficeEditorCommand<Context> {
  canExecute?: (context: Context, args: readonly unknown[]) => boolean;
  execute: (context: Context, args: readonly unknown[]) => unknown;
}

interface OfficeEditorExtensionInstance<Context> {
  commands: ReadonlyMap<string, ErasedOfficeEditorCommand<Context>>;
  keyboardShortcuts: readonly ErasedOfficeEditorKeyboardShortcut<Context>[];
  onCreate: (context: Context) => void;
  onDestroy: (context: Context) => void;
  onUpdate: (context: Context, previousContext: Context) => void;
  storage: unknown;
}

interface ErasedOfficeEditorKeyboardShortcut<Context> {
  handler: (
    context: Context,
    commands: unknown,
    can: unknown,
    event: KeyboardEvent,
  ) => boolean;
  shortcut: string;
}

export interface OfficeEditorExtension<Context, Commands> {
  readonly name: string;
  readonly priority: number;
  createInstance: () => OfficeEditorExtensionInstance<Context>;
  /**
   * Type-only marker used to keep command maps aligned across extensions.
   * It is intentionally absent at runtime.
   */
  readonly commandTypes?: Commands;
}

export interface OfficeEditorRuntime<Context, Commands> {
  readonly commands: Commands;
  readonly extensionNames: readonly string[];
  readonly storage: Readonly<Record<string, unknown>>;
  can: () => OfficeEditorCanCommands<Commands>;
  handleKeyDown: (event: KeyboardEvent) => boolean;
  mount: () => void;
  unmount: () => void;
  updateContext: (context: Context) => void;
}

export interface OfficeEditorRuntimeOptions<Context> {
  /**
   * Supplies render-current controlled state to commands and can() checks.
   * Lifecycle callbacks still receive contexts through updateContext().
   */
  getCurrentContext?: () => Context;
}

const DEFAULT_EXTENSION_PRIORITY = 100;

export function createOfficeEditorExtension<
  Context,
  Commands,
  Storage = Record<string, never>,
>(
  config: OfficeEditorExtensionConfig<Context, Commands, Storage>,
): OfficeEditorExtension<Context, Commands> {
  const name = config.name.trim();
  if (!name) throw new Error('Editor extension name cannot be empty.');
  const priority = Number.isFinite(config.priority)
    ? Number(config.priority)
    : DEFAULT_EXTENSION_PRIORITY;

  return {
    name,
    priority,
    createInstance: () => {
      const storage = config.addStorage?.() ?? ({} as Storage);
      const definitions = config.addCommands?.({ storage }) ?? {};
      const commands = new Map<string, ErasedOfficeEditorCommand<Context>>();
      const keyboardShortcuts = Object.entries(
        config.addKeyboardShortcuts?.({ storage }) ?? {},
      ).map(([shortcut, handler]) => {
        const normalizedShortcut = shortcut.trim();
        if (!normalizedShortcut) {
          throw new Error(
            `Keyboard shortcut in extension "${name}" cannot be empty.`,
          );
        }
        if (typeof handler !== 'function') {
          throw new Error(
            `Keyboard shortcut "${shortcut}" in extension "${name}" must define a handler.`,
          );
        }
        return {
          shortcut: normalizedShortcut,
          handler: (
            context: Context,
            commands: unknown,
            can: unknown,
            event: KeyboardEvent,
          ) =>
            handler(
              {
                context,
                commands: commands as Commands,
                can: can as OfficeEditorCanCommands<Commands>,
                storage,
              },
              event,
            ),
        };
      });

      for (const [commandName, candidate] of Object.entries(
        definitions as Record<string, unknown>,
      )) {
        const definition = candidate as {
          canExecute?: unknown;
          execute?: unknown;
        };
        if (typeof definition.execute !== 'function') {
          throw new Error(
            `Editor command "${commandName}" in extension "${name}" must define execute().`,
          );
        }
        const execute = definition.execute as (
          context: Context,
          ...args: unknown[]
        ) => unknown;
        const canExecute =
          typeof definition.canExecute === 'function'
            ? (definition.canExecute as (
                context: Context,
                ...args: unknown[]
              ) => boolean)
            : undefined;
        commands.set(commandName, {
          canExecute: canExecute
            ? (context, args) => canExecute(context, ...args)
            : undefined,
          execute: (context, args) => execute(context, ...args),
        });
      }

      return {
        commands,
        keyboardShortcuts,
        onCreate: (context) => config.onCreate?.({ context, storage }),
        onDestroy: (context) => config.onDestroy?.({ context, storage }),
        onUpdate: (context, previousContext) =>
          config.onUpdate?.({ context, previousContext, storage }),
        storage,
      };
    },
  };
}

export function createOfficeEditorRuntime<Context, Commands>(
  initialContext: Context,
  extensions: readonly OfficeEditorExtension<Context, Commands>[],
  options: OfficeEditorRuntimeOptions<Context> = {},
): OfficeEditorRuntime<Context, Commands> {
  const orderedExtensions = extensions
    .map((extension, index) => ({ extension, index }))
    .sort(
      (left, right) =>
        right.extension.priority - left.extension.priority ||
        left.index - right.index,
    )
    .map(({ extension }) => extension);
  const extensionNames = new Set<string>();
  const commandDefinitions = new Map<
    string,
    ErasedOfficeEditorCommand<Context>
  >();
  const storage: Record<string, unknown> = {};
  const instances = orderedExtensions.map((extension) => {
    if (extensionNames.has(extension.name)) {
      throw new Error(`Duplicate editor extension "${extension.name}".`);
    }
    extensionNames.add(extension.name);
    const instance = extension.createInstance();
    storage[extension.name] = instance.storage;
    for (const [commandName, definition] of instance.commands) {
      if (commandDefinitions.has(commandName)) {
        throw new Error(`Duplicate editor command "${commandName}".`);
      }
      commandDefinitions.set(commandName, definition);
    }
    return instance;
  });

  let context = initialContext;
  const readContext = () => options.getCurrentContext?.() ?? context;
  let mounted = false;
  const commandApi: Record<string, (...args: unknown[]) => unknown> = {};
  const canApi: Record<string, (...args: unknown[]) => boolean> = {};
  for (const [commandName, definition] of commandDefinitions) {
    commandApi[commandName] = (...args) =>
      definition.execute(readContext(), args);
    canApi[commandName] = (...args) =>
      definition.canExecute?.(readContext(), args) ?? true;
  }

  return {
    commands: commandApi as unknown as Commands,
    extensionNames: Object.freeze([...extensionNames]),
    storage: Object.freeze(storage),
    can: () => canApi as unknown as OfficeEditorCanCommands<Commands>,
    handleKeyDown: (event) => {
      if (event.defaultPrevented || isOfficeCompositionKeyboardEvent(event)) {
        return false;
      }
      for (const instance of instances) {
        for (const shortcut of instance.keyboardShortcuts) {
          if (
            matchesOfficeEditorKeyboardShortcut(event, shortcut.shortcut) &&
            shortcut.handler(readContext(), commandApi, canApi, event)
          ) {
            event.preventDefault();
            return true;
          }
        }
      }
      return false;
    },
    mount: () => {
      if (mounted) return;
      mounted = true;
      for (const instance of instances) instance.onCreate(readContext());
    },
    unmount: () => {
      if (!mounted) return;
      for (const instance of [...instances].reverse()) {
        instance.onDestroy(readContext());
      }
      mounted = false;
    },
    updateContext: (nextContext) => {
      const previousContext = context;
      context = nextContext;
      if (!mounted || Object.is(previousContext, nextContext)) return;
      for (const instance of instances) {
        instance.onUpdate(context, previousContext);
      }
    },
  };
}

export { matchesOfficeEditorKeyboardShortcut } from '../../../keyboard-shortcuts';
