import { ClipboardPaste, Copy, Eraser, Scissors } from 'lucide-react';
import { showToast } from '../../../state/app-state';
import type { WorkspaceContextMenuItem } from '../../workspace/components/workspace-context-menu';
import type { WorkSpreadsheetAgentSelection } from '../work-spreadsheet-agent-context';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';

type SpreadsheetContextCan = Pick<
  SpreadsheetEditorCanCommands,
  'clearSelectedCells' | 'pasteCells'
>;

type SpreadsheetContextCommands = Pick<
  SpreadsheetEditorCommands,
  'clearSelectedCells' | 'pasteCells'
>;

export interface SpreadsheetClipboardPort {
  readText(): Promise<string>;
  writeText(value: string): Promise<void>;
}

const spreadsheetSystemClipboardTimeoutMs = 800;

export function spreadsheetCoreContextMenuItems({
  can,
  clipboard = browserSpreadsheetClipboard,
  commands,
  selection,
}: {
  can: SpreadsheetContextCan;
  clipboard?: SpreadsheetClipboardPort;
  commands: SpreadsheetContextCommands;
  selection: Pick<WorkSpreadsheetAgentSelection, 'clipboard' | 'reference'>;
}): WorkspaceContextMenuItem[] {
  const copy = (cut: boolean) => {
    void copySpreadsheetSelection(clipboard, selection.clipboard, cut).then(
      (copied) => {
        if (copied && cut && !commands.clearSelectedCells()) {
          showToast('选区已复制，但无法清除原内容。', 'error');
        }
      },
    );
  };
  return [
    {
      id: 'cut-cells',
      label: '剪切',
      icon: <Scissors size={14} />,
      shortcut: '⌘X',
      ariaKeyShortcut: 'Control+X Meta+X',
      disabled: !can.clearSelectedCells(),
      onSelect: () => copy(true),
    },
    {
      id: 'copy-cells',
      label: '复制',
      icon: <Copy size={14} />,
      shortcut: '⌘C',
      ariaKeyShortcut: 'Control+C Meta+C',
      onSelect: () => copy(false),
    },
    {
      id: 'paste-cells',
      label: '粘贴',
      icon: <ClipboardPaste size={14} />,
      shortcut: '⌘V',
      ariaKeyShortcut: 'Control+V Meta+V',
      disabled: !can.pasteCells([['']]),
      onSelect: () =>
        void pasteSpreadsheetSelection(clipboard, commands.pasteCells),
    },
    {
      id: 'clear-cells',
      label: '清除内容',
      icon: <Eraser size={14} />,
      shortcut: 'Delete',
      ariaKeyShortcut: 'Delete',
      separatorBefore: true,
      disabled: !can.clearSelectedCells(),
      onSelect: () => {
        if (!commands.clearSelectedCells()) {
          showToast(`无法清除选区 ${selection.reference}。`, 'error');
        }
      },
    },
  ];
}

export function parseSpreadsheetClipboardText(value: string): string[][] {
  if (!value) return [];
  const normalized = value.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (!normalized) return [];
  const rows = normalized.split('\n').map((row) => row.split('\t'));
  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => [
    ...row,
    ...Array.from({ length: width - row.length }, () => ''),
  ]);
}

export const browserSpreadsheetClipboard: SpreadsheetClipboardPort = {
  readText: () =>
    readSpreadsheetClipboardText(
      typeof navigator !== 'undefined' && navigator.clipboard?.readText
        ? () => navigator.clipboard.readText()
        : undefined,
      readLocalClipboard,
      spreadsheetSystemClipboardTimeoutMs,
      canReadBrowserSystemClipboard,
    ),
  writeText: (value) =>
    writeSpreadsheetClipboardText(
      value,
      typeof navigator !== 'undefined' && navigator.clipboard?.writeText
        ? (next) => navigator.clipboard.writeText(next)
        : undefined,
      writeLocalClipboard,
    ),
};

export async function readSpreadsheetClipboardText(
  readSystem: (() => Promise<string>) | undefined,
  readLocal: () => string,
  timeoutMs = spreadsheetSystemClipboardTimeoutMs,
  canReadSystem: () => boolean | Promise<boolean> = () => true,
): Promise<string> {
  if (readSystem && (await canReadSystem())) {
    try {
      const value = await withSpreadsheetClipboardTimeout(
        readSystem(),
        timeoutMs,
      );
      if (value) return value;
    } catch {
      // The editor-local clipboard remains available when browser access is
      // denied or does not settle promptly.
    }
  }
  return readLocal();
}

async function canReadBrowserSystemClipboard(): Promise<boolean> {
  if (typeof navigator === 'undefined' || !navigator.permissions?.query) {
    return true;
  }
  try {
    const permission = await navigator.permissions.query({
      name: 'clipboard-read' as PermissionName,
    });
    return permission.state === 'granted';
  } catch {
    return true;
  }
}

export async function writeSpreadsheetClipboardText(
  value: string,
  writeSystem: ((value: string) => Promise<void>) | undefined,
  writeLocal: (value: string) => void,
  timeoutMs = spreadsheetSystemClipboardTimeoutMs,
): Promise<void> {
  writeLocal(value);
  if (!writeSystem) return;
  try {
    await withSpreadsheetClipboardTimeout(writeSystem(value), timeoutMs);
  } catch {
    // Keep the editor-local copy usable when the system clipboard is blocked.
  }
}

async function copySpreadsheetSelection(
  clipboard: SpreadsheetClipboardPort,
  value: string,
  cut: boolean,
): Promise<boolean> {
  try {
    await clipboard.writeText(value);
    showToast(cut ? '选区已剪切' : '选区已复制', 'success');
    return true;
  } catch {
    showToast('无法访问剪贴板，请使用系统快捷键。', 'error');
    return false;
  }
}

async function pasteSpreadsheetSelection(
  clipboard: SpreadsheetClipboardPort,
  paste: SpreadsheetContextCommands['pasteCells'],
): Promise<void> {
  try {
    const values = parseSpreadsheetClipboardText(await clipboard.readText());
    if (!values.length) {
      showToast('剪贴板中没有可粘贴的表格内容。', 'error');
      return;
    }
    if (!paste(values)) {
      showToast('当前选区无法粘贴这些内容。', 'error');
    }
  } catch {
    showToast('无法读取剪贴板，请使用系统粘贴快捷键。', 'error');
  }
}

function readLocalClipboard(): string {
  try {
    return sessionStorage.getItem('a3s-office-spreadsheet-clipboard') ?? '';
  } catch {
    return '';
  }
}

function writeLocalClipboard(value: string): void {
  try {
    sessionStorage.setItem('a3s-office-spreadsheet-clipboard', value);
  } catch {
    // System clipboard access may still succeed.
  }
}

async function withSpreadsheetClipboardTimeout<T>(
  request: Promise<T>,
  timeoutMs: number,
): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  const expired = new Promise<never>((_resolve, reject) => {
    timeout = setTimeout(
      () => reject(new Error('Spreadsheet clipboard request timed out.')),
      Math.max(0, timeoutMs),
    );
  });
  try {
    return await Promise.race([request, expired]);
  } finally {
    if (timeout) clearTimeout(timeout);
  }
}
