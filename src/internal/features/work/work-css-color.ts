export function normalizeCssColor(
  source: string | null | undefined,
): 'transparent' | `#${string}` | null {
  const value = source?.trim().toLowerCase();
  if (!value) return null;
  if (value === 'transparent') return 'transparent';
  const shortHex = /^#([0-9a-f]{3})$/i.exec(value);
  if (shortHex?.[1]) {
    return `#${Array.from(shortHex[1])
      .map((channel) => `${channel}${channel}`)
      .join('')}`;
  }
  if (/^#[0-9a-f]{6}$/i.test(value)) return value as `#${string}`;
  const rgb =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([\d.]+))?\s*\)$/i.exec(
      value,
    );
  if (!rgb) return null;
  const channels = rgb.slice(1, 4).map(Number);
  if (channels.some((channel) => channel < 0 || channel > 255)) return null;
  if (rgb[4] !== undefined && Number(rgb[4]) === 0) return 'transparent';
  if (rgb[4] !== undefined && Number(rgb[4]) !== 1) return null;
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}
