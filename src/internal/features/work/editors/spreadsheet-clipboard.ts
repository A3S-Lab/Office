import { showToast } from '../../../state/app-state';
import {
  createSpreadsheetTextClipboardSnapshot,
  normalizeSpreadsheetClipboardText,
  type SpreadsheetClipboardSnapshot,
} from './spreadsheet-paste-special';

export interface SpreadsheetClipboardPort {
  readText(): Promise<string>;
  writeText(value: string): Promise<void>;
}

const spreadsheetSystemClipboardTimeoutMs = 800;
const spreadsheetLocalClipboardKey = 'a3s-office-spreadsheet-clipboard';

let richSpreadsheetClipboard: SpreadsheetClipboardSnapshot | null = null;
let browserClipboardWriteRevision = 0;
let preferLocalSpreadsheetClipboardRead = false;

export function parseSpreadsheetClipboardText(value: string): string[][] {
  const normalized = normalizeSpreadsheetClipboardText(value);
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
    preferLocalSpreadsheetClipboardRead
      ? Promise.resolve(readLocalSpreadsheetClipboardText())
      : readSpreadsheetClipboardText(
          typeof navigator !== 'undefined' && navigator.clipboard?.readText
            ? () => navigator.clipboard.readText()
            : undefined,
          readLocalSpreadsheetClipboardText,
          spreadsheetSystemClipboardTimeoutMs,
          canReadBrowserSystemClipboard,
        ),
  writeText: (value) => {
    const revision = ++browserClipboardWriteRevision;
    preferLocalSpreadsheetClipboardRead = true;
    return writeSpreadsheetClipboardText(
      value,
      typeof navigator !== 'undefined' && navigator.clipboard?.writeText
        ? (next) => navigator.clipboard.writeText(next)
        : undefined,
      writeLocalSpreadsheetClipboardText,
      spreadsheetSystemClipboardTimeoutMs,
      (systemWriteSucceeded) => {
        if (revision === browserClipboardWriteRevision) {
          preferLocalSpreadsheetClipboardRead = !systemWriteSucceeded;
        }
      },
    );
  },
};

export function storeRichSpreadsheetClipboard(
  snapshot: SpreadsheetClipboardSnapshot,
): void {
  richSpreadsheetClipboard = structuredClone(snapshot);
  writeLocalSpreadsheetClipboardText(snapshot.plainText);
}

export function spreadsheetClipboardSnapshotForText(
  plainText: string,
): SpreadsheetClipboardSnapshot | null {
  const normalized = normalizeSpreadsheetClipboardText(plainText);
  if (
    richSpreadsheetClipboard &&
    richSpreadsheetClipboard.plainText === normalized
  ) {
    return structuredClone(richSpreadsheetClipboard);
  }
  if (richSpreadsheetClipboard) {
    richSpreadsheetClipboard = null;
    browserClipboardWriteRevision += 1;
    preferLocalSpreadsheetClipboardRead = false;
  }
  return createSpreadsheetTextClipboardSnapshot(normalized);
}

export function clearRichSpreadsheetClipboard(): void {
  richSpreadsheetClipboard = null;
  browserClipboardWriteRevision += 1;
  preferLocalSpreadsheetClipboardRead = false;
}

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

export async function writeSpreadsheetClipboardText(
  value: string,
  writeSystem: ((value: string) => Promise<void>) | undefined,
  writeLocal: (value: string) => void,
  timeoutMs = spreadsheetSystemClipboardTimeoutMs,
  onSystemWriteResult?: (succeeded: boolean) => void,
): Promise<void> {
  writeLocal(value);
  if (!writeSystem) {
    onSystemWriteResult?.(false);
    return;
  }
  try {
    await withSpreadsheetClipboardTimeout(writeSystem(value), timeoutMs);
    onSystemWriteResult?.(true);
  } catch {
    onSystemWriteResult?.(false);
    // Keep the editor-local copy usable when the system clipboard is blocked.
  }
}

export async function copySpreadsheetSelection(
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

function readLocalSpreadsheetClipboardText(): string {
  try {
    return sessionStorage.getItem(spreadsheetLocalClipboardKey) ?? '';
  } catch {
    return '';
  }
}

function writeLocalSpreadsheetClipboardText(value: string): void {
  try {
    sessionStorage.setItem(spreadsheetLocalClipboardKey, value);
  } catch {
    // System clipboard access may still succeed.
  }
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

async function withSpreadsheetClipboardTimeout<T>(
  operation: Promise<T>,
  timeoutMs: number,
): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      operation,
      new Promise<T>((_resolve, reject) => {
        timeout = setTimeout(
          () => reject(new Error('Spreadsheet clipboard request timed out.')),
          timeoutMs,
        );
      }),
    ]);
  } finally {
    if (timeout !== undefined) clearTimeout(timeout);
  }
}
