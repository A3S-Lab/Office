export function matchesOfficeEditorKeyboardShortcut(
  event: KeyboardEvent,
  shortcut: string,
): boolean {
  return matchesKeyboardShortcut(event, shortcutParts(shortcut, '-'));
}

export function matchesAriaKeyShortcuts(
  event: KeyboardEvent,
  shortcuts: string | undefined,
): boolean {
  if (!shortcuts) return false;
  return shortcuts
    .split(/\s+/)
    .filter(Boolean)
    .some((shortcut) =>
      matchesKeyboardShortcut(event, shortcutParts(shortcut, '+')),
    );
}

function shortcutParts(shortcut: string, separator: '-' | '+'): string[] {
  const normalized = shortcut.trim();
  if (separator === '+' && normalized.endsWith('++')) {
    return [
      ...normalized
        .slice(0, -1)
        .split(separator)
        .map((part) => part.trim().toLocaleLowerCase())
        .filter(Boolean),
      'plus',
    ];
  }
  return normalized
    .split(separator)
    .map((part) => part.trim().toLocaleLowerCase())
    .filter(Boolean);
}

function matchesKeyboardShortcut(
  event: KeyboardEvent,
  parts: readonly string[],
): boolean {
  const key = parts.at(-1);
  if (!key) return false;
  const modifiers = new Set(parts.slice(0, -1));
  const mod = modifiers.has('mod');
  const control = modifiers.has('control') || modifiers.has('ctrl');
  const meta = modifiers.has('meta') || modifiers.has('cmd');
  const shift = modifiers.has('shift');
  const alt = modifiers.has('alt') || modifiers.has('option');

  if (mod && !(event.metaKey || event.ctrlKey)) return false;
  if (!mod && event.ctrlKey !== control) return false;
  if (!mod && event.metaKey !== meta) return false;
  if (event.shiftKey !== shift) return false;
  if (event.altKey !== alt) return false;

  const eventKey = event.key.toLocaleLowerCase();
  return (
    eventKey === key ||
    shortcutKeyAlias(eventKey) === key ||
    shortcutCodeKey(event.code) === key
  );
}

function shortcutKeyAlias(key: string): string {
  if (key === ' ') return 'space';
  if (key === '+') return 'plus';
  if (key === '-') return 'minus';
  if (key === '=') return 'equal';
  if (key === 'esc') return 'escape';
  if (key === 'del') return 'delete';
  return key;
}

function shortcutCodeKey(code: string): string | null {
  if (/^Key[A-Z]$/.test(code)) return code.slice(3).toLocaleLowerCase();
  if (/^Digit[0-9]$/.test(code)) return code.slice(5);
  if (code === 'Space') return 'space';
  if (code === 'Equal') return 'equal';
  if (code === 'Minus') return 'minus';
  if (code === 'Comma') return ',';
  if (code === 'Period') return '.';
  return null;
}
