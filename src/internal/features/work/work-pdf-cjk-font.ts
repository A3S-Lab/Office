import type { jsPDF as JsPdf } from 'jspdf';

export interface WorkPdfCjkFontRegistration {
  /** Logical PDF font family name registered with jsPDF. */
  family: string;
  /** Base64-encoded TrueType font bytes (no data: URI prefix). */
  base64: string;
  /** VFS file name passed to jsPDF addFileToVFS / addFont. */
  vfsName: string;
}

const MAX_FONT_BYTES = 32 * 1024 * 1024;
const DEFAULT_FAMILY = 'WorkCjk';
const DEFAULT_VFS_NAME = 'WorkCjk.ttf';

let registered: WorkPdfCjkFontRegistration | null = null;
const documentsWithFont = new WeakSet<object>();

/**
 * Registers an optional host-supplied CJK TrueType face for Writer PDF export.
 * Without a registration, searchable vector text stays Latin/Latin-1 only.
 * Call {@link clearWorkPdfCjkFont} in tests after use.
 */
export function registerWorkPdfCjkFont(options: {
  data: ArrayBuffer | Uint8Array | string;
  family?: string;
  vfsName?: string;
}): WorkPdfCjkFontRegistration {
  const base64 = normalizeFontBase64(options.data);
  const family = (options.family?.trim() || DEFAULT_FAMILY).slice(0, 64);
  const vfsName = (options.vfsName?.trim() || DEFAULT_VFS_NAME).slice(0, 128);
  if (!family || !vfsName.toLowerCase().endsWith('.ttf')) {
    throw new Error('Work PDF CJK fonts require a TrueType (.ttf) face.');
  }
  registered = { family, base64, vfsName };
  return registered;
}

export function clearWorkPdfCjkFont(): void {
  registered = null;
}

export function workPdfCjkFontRegistration(): WorkPdfCjkFontRegistration | null {
  return registered;
}

export function workPdfCjkFontRegistered(): boolean {
  return registered !== null;
}

/**
 * Ensures the registered CJK face is present on a jsPDF document once.
 * Returns the family name when available, otherwise null.
 */
export function ensureWorkPdfCjkFontOnDocument(pdf: JsPdf): string | null {
  if (!registered) return null;
  if (!documentsWithFont.has(pdf)) {
    pdf.addFileToVFS(registered.vfsName, registered.base64);
    pdf.addFont(registered.vfsName, registered.family, 'normal');
    documentsWithFont.add(pdf);
  }
  return registered.family;
}

export function workPdfTextNeedsCjkFont(text: string): boolean {
  return /[\u3000-\u303f\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(
    text,
  );
}

function normalizeFontBase64(data: ArrayBuffer | Uint8Array | string): string {
  if (typeof data === 'string') {
    const trimmed = data.trim();
    const raw = trimmed.includes(',')
      ? trimmed.slice(trimmed.indexOf(',') + 1)
      : trimmed;
    if (!raw || raw.length > MAX_FONT_BYTES * 2) {
      throw new Error('Work PDF CJK font payload is empty or too large.');
    }
    return raw.replace(/\s+/g, '');
  }
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (!bytes.byteLength || bytes.byteLength > MAX_FONT_BYTES) {
    throw new Error('Work PDF CJK font payload is empty or too large.');
  }
  if (!isTrueTypeFontBytes(bytes)) {
    throw new Error(
      'Work PDF CJK fonts require TrueType (.ttf) bytes, not OpenType CFF.',
    );
  }
  return bytesToBase64(bytes);
}

function isTrueTypeFontBytes(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false;
  // sfnt version 0x00010000
  if (
    bytes[0] === 0x00 &&
    bytes[1] === 0x01 &&
    bytes[2] === 0x00 &&
    bytes[3] === 0x00
  ) {
    return true;
  }
  // Apple TrueType 'true'
  if (
    bytes[0] === 0x74 &&
    bytes[1] === 0x72 &&
    bytes[2] === 0x75 &&
    bytes[3] === 0x65
  ) {
    return true;
  }
  return false;
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(
      ...bytes.subarray(index, Math.min(bytes.length, index + chunk)),
    );
  }
  return btoa(binary);
}
