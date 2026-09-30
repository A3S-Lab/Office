const TASK_LINE = /^(\s*(?:[-*+]|\d+[.)])\s+)\[([ xX])\](.*)$/;

/** Set the nth GFM task in document order. An unknown index leaves the source unchanged. */
export function setMarkdownTaskChecked(
  markdown: string,
  index: number,
  checked: boolean,
): string {
  if (!Number.isInteger(index) || index < 0) return markdown;
  let seen = 0;
  const lines = markdown.split('\n');
  const mark = checked ? 'x' : ' ';
  for (let line = 0; line < lines.length; line += 1) {
    const match = TASK_LINE.exec(lines[line] ?? '');
    if (!match) continue;
    if (seen === index) {
      lines[line] = `${match[1]}[${mark}]${match[3] ?? ''}`;
      return lines.join('\n');
    }
    seen += 1;
  }
  return markdown;
}

/** Flip the nth GFM task in document order. An unknown index leaves the source unchanged. */
export function toggleMarkdownTask(markdown: string, index: number): string {
  if (!Number.isInteger(index) || index < 0) return markdown;
  let seen = 0;
  const lines = markdown.split('\n');
  for (let line = 0; line < lines.length; line += 1) {
    const match = TASK_LINE.exec(lines[line] ?? '');
    if (!match) continue;
    if (seen === index) {
      return setMarkdownTaskChecked(
        markdown,
        index,
        match[2]?.toLowerCase() !== 'x',
      );
    }
    seen += 1;
  }
  return markdown;
}
