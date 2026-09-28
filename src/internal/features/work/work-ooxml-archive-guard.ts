/** Resource limits for ZIP-based Office packages, matching `PackageLimits` in the native core. */
export interface OoxmlArchiveLimits {
  maxArchiveBytes: number;
  maxEntries: number;
  maxPartBytes: number;
  maxUncompressedBytes: number;
  maxCompressionRatio: number;
}

export const DEFAULT_OOXML_ARCHIVE_LIMITS: Readonly<OoxmlArchiveLimits> =
  Object.freeze({
    maxArchiveBytes: 512 * 1024 * 1024,
    maxEntries: 16_384,
    maxPartBytes: 128 * 1024 * 1024,
    maxUncompressedBytes: 1024 * 1024 * 1024,
    maxCompressionRatio: 250,
  });

// Small parts may compress extremely well (empty XML); only large expansions
// are treated as bombs.
const COMPRESSION_RATIO_MINIMUM_BYTES = 1024 * 1024;

export type OoxmlArchiveLimitCode =
  | 'package.archive.archive-bytes'
  | 'package.archive.entries'
  | 'package.archive.part-bytes'
  | 'package.archive.uncompressed-bytes'
  | 'package.archive.compression-ratio'
  | 'package.archive.malformed';

export class OoxmlArchiveLimitError extends Error {
  readonly code: OoxmlArchiveLimitCode;

  constructor(code: OoxmlArchiveLimitCode, message: string) {
    super(message);
    this.name = 'OoxmlArchiveLimitError';
    this.code = code;
  }
}

const LOCAL_FILE_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const ZIP64_END_SIGNATURE = 0x06064b50;
const ZIP64_EXTRA_FIELD_ID = 0x0001;
const MAX_EOCD_SEARCH_BYTES = 22 + 0xffff;

export function isZipArchive(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    new DataView(bytes.buffer, bytes.byteOffset, 4).getUint32(0, true) ===
      LOCAL_FILE_SIGNATURE
  );
}

/**
 * Validates entry count and declared sizes from the ZIP central directory
 * without inflating any part, so a decompression bomb is rejected before
 * JSZip or SheetJS allocate its expanded bytes.
 */
export function assertOoxmlArchiveWithinLimits(
  bytes: Uint8Array,
  limits: OoxmlArchiveLimits = DEFAULT_OOXML_ARCHIVE_LIMITS,
): void {
  if (bytes.length > limits.maxArchiveBytes) {
    throw new OoxmlArchiveLimitError(
      'package.archive.archive-bytes',
      `The Office package exceeds the ${limits.maxArchiveBytes}-byte archive limit.`,
    );
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const directory = readCentralDirectoryLocation(view);
  if (directory.entryCount > limits.maxEntries) {
    throw entriesError(limits);
  }

  let offset = directory.offset;
  let totalUncompressed = 0;
  for (let index = 0; index < directory.entryCount; index += 1) {
    requireRange(view, offset, 46);
    if (view.getUint32(offset, true) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw malformed('A central-directory entry has an invalid signature.');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const entryLength = 46 + nameLength + extraLength + commentLength;
    requireRange(view, offset, entryLength);
    const sizes = entrySizes(view, offset, nameLength, extraLength);

    if (sizes.uncompressed > limits.maxPartBytes) {
      throw new OoxmlArchiveLimitError(
        'package.archive.part-bytes',
        `An Office package part exceeds the ${limits.maxPartBytes}-byte part limit.`,
      );
    }
    if (
      sizes.uncompressed >= COMPRESSION_RATIO_MINIMUM_BYTES &&
      (sizes.compressed === 0 ||
        sizes.uncompressed / sizes.compressed > limits.maxCompressionRatio)
    ) {
      throw new OoxmlArchiveLimitError(
        'package.archive.compression-ratio',
        `An Office package part exceeds the ${limits.maxCompressionRatio}x compression-ratio limit.`,
      );
    }
    totalUncompressed += sizes.uncompressed;
    if (totalUncompressed > limits.maxUncompressedBytes) {
      throw new OoxmlArchiveLimitError(
        'package.archive.uncompressed-bytes',
        `The Office package expands beyond the ${limits.maxUncompressedBytes}-byte limit.`,
      );
    }
    offset += entryLength;
  }
}

interface CentralDirectoryLocation {
  entryCount: number;
  offset: number;
}

function readCentralDirectoryLocation(
  view: DataView,
): CentralDirectoryLocation {
  const end = findEndOfCentralDirectory(view);
  let entryCount = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true);
  if (entryCount === 0xffff || offset === 0xffffffff) {
    const locator = end - 20;
    if (
      locator < 0 ||
      view.getUint32(locator, true) !== ZIP64_LOCATOR_SIGNATURE
    ) {
      throw malformed('The ZIP64 end-of-central-directory locator is missing.');
    }
    const zip64End = safeNumber(view.getBigUint64(locator + 8, true));
    requireRange(view, zip64End, 56);
    if (view.getUint32(zip64End, true) !== ZIP64_END_SIGNATURE) {
      throw malformed('The ZIP64 end-of-central-directory record is invalid.');
    }
    entryCount = safeNumber(view.getBigUint64(zip64End + 32, true));
    offset = safeNumber(view.getBigUint64(zip64End + 48, true));
  }
  if (offset > end) {
    throw malformed('The central directory starts outside the archive.');
  }
  return { entryCount, offset };
}

function findEndOfCentralDirectory(view: DataView): number {
  const lowest = Math.max(0, view.byteLength - MAX_EOCD_SEARCH_BYTES);
  for (let offset = view.byteLength - 22; offset >= lowest; offset -= 1) {
    if (view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      return offset;
    }
  }
  throw malformed('The end-of-central-directory record is missing.');
}

function entrySizes(
  view: DataView,
  offset: number,
  nameLength: number,
  extraLength: number,
): { compressed: number; uncompressed: number } {
  let compressed = view.getUint32(offset + 20, true);
  let uncompressed = view.getUint32(offset + 24, true);
  if (compressed !== 0xffffffff && uncompressed !== 0xffffffff) {
    return { compressed, uncompressed };
  }
  let extra = offset + 46 + nameLength;
  const extraEnd = extra + extraLength;
  while (extra + 4 <= extraEnd) {
    const id = view.getUint16(extra, true);
    const size = view.getUint16(extra + 2, true);
    let field = extra + 4;
    if (id === ZIP64_EXTRA_FIELD_ID) {
      if (uncompressed === 0xffffffff) {
        requireRange(view, field, 8);
        uncompressed = safeNumber(view.getBigUint64(field, true));
        field += 8;
      }
      if (compressed === 0xffffffff) {
        requireRange(view, field, 8);
        compressed = safeNumber(view.getBigUint64(field, true));
      }
      return { compressed, uncompressed };
    }
    extra += 4 + size;
  }
  throw malformed('A ZIP64 entry is missing its size extra field.');
}

function requireRange(view: DataView, offset: number, length: number): void {
  if (offset < 0 || offset + length > view.byteLength) {
    throw malformed('The ZIP central directory is truncated.');
  }
}

function safeNumber(value: bigint): number {
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw malformed('A ZIP64 size or offset is out of range.');
  }
  return Number(value);
}

function entriesError(limits: OoxmlArchiveLimits): OoxmlArchiveLimitError {
  return new OoxmlArchiveLimitError(
    'package.archive.entries',
    `The Office package exceeds the ${limits.maxEntries}-entry limit.`,
  );
}

function malformed(message: string): OoxmlArchiveLimitError {
  return new OoxmlArchiveLimitError('package.archive.malformed', message);
}
