export function stableJson(value: unknown): string {
  return JSON.stringify(sortJsonValue(value));
}

export function stableHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

export function wordParagraphId(value: string): string {
  const number = Number.parseInt(stableHash(value), 16) & 0x7fff_ffff || 1;
  return number.toString(16).toUpperCase().padStart(8, '0');
}

export function normalizedDate(value: string): string {
  const timestamp = Date.parse(value);
  return Number.isFinite(timestamp)
    ? new Date(timestamp).toISOString()
    : '1970-01-01T00:00:00.000Z';
}

export function boundedMetadata(value: string, maximum: number): string {
  return value.trim().slice(0, maximum);
}

function sortJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortJsonValue);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([name, child]) => [name, sortJsonValue(child)]),
  );
}
