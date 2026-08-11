export function decodeXmlBytes(bytes: Uint8Array, label: string): string {
  let encoding: 'utf-8' | 'utf-16le' | 'utf-16be' = 'utf-8';
  let offset = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    offset = 3;
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = 'utf-16le';
    offset = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = 'utf-16be';
    offset = 2;
  } else if (
    bytes[0] === 0x3c &&
    bytes[1] === 0 &&
    bytes[2] === 0x3f &&
    bytes[3] === 0
  ) {
    encoding = 'utf-16le';
  } else if (
    bytes[0] === 0 &&
    bytes[1] === 0x3c &&
    bytes[2] === 0 &&
    bytes[3] === 0x3f
  ) {
    encoding = 'utf-16be';
  }
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(
      bytes.subarray(offset),
    );
  } catch {
    throw new Error(`${label} uses an invalid ${encoding} XML encoding.`);
  }
}

export function serializeUtf8Xml(document: Document): string {
  const serialized = new XMLSerializer().serializeToString(document);
  const body = serialized.replace(/^\s*<\?xml[^?]*\?>\s*/i, '');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
}
