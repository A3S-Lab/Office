import JSZip from 'jszip';
import {
  attribute,
  directChild,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function patchDocxNumberingRestartRules(
  buffer: ArrayBuffer,
  rules: ReadonlyArray<ReadonlyMap<number, number>>,
): Promise<ArrayBuffer> {
  if (!rules.some((levels) => levels.size)) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/numbering.xml');
  if (!entry) return buffer;
  const document = parseXml(await entry.async('text'), 'word/numbering.xml');
  const abstractNumbering = directChildren(
    document.documentElement,
    'abstractNum',
  );
  if (abstractNumbering.length < rules.length) return buffer;
  const generated = abstractNumbering.slice(-rules.length);
  for (const [index, levels] of rules.entries()) {
    const abstract = generated[index];
    if (!abstract) continue;
    for (const [level, restartAfterLevel] of levels) {
      const definition = directChildren(abstract, 'lvl').find(
        (candidate) => attribute(candidate, 'ilvl') === String(level),
      );
      if (!definition) continue;
      let restart = directChild(definition, 'lvlRestart');
      if (!restart) {
        restart = document.createElementNS(
          WORD_NAMESPACE,
          wordQualifiedName(document.documentElement, 'lvlRestart'),
        );
        definition.insertBefore(
          restart,
          directChild(definition, 'pPr') ?? null,
        );
      }
      restart.setAttributeNS(
        WORD_NAMESPACE,
        wordQualifiedName(document.documentElement, 'val'),
        String(restartAfterLevel),
      );
    }
  }
  archive.file(
    'word/numbering.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function wordQualifiedName(root: Element, localName: string): string {
  return `${xmlNamespacePrefix(root, WORD_NAMESPACE) ?? root.prefix ?? 'w'}:${localName}`;
}
