import type { Extensions } from '@tiptap/core';

export function mergeOfficeTiptapExtensions(
  editorName: string,
  builtInExtensions: Extensions,
  additionalExtensions: Extensions,
): Extensions {
  const names = new Set<string>();
  for (const extension of builtInExtensions) names.add(extension.name);
  for (const extension of additionalExtensions) {
    if (names.has(extension.name)) {
      throw new Error(
        `${editorName} extension "${extension.name}" is already registered.`,
      );
    }
    names.add(extension.name);
  }
  return [...builtInExtensions, ...additionalExtensions];
}
