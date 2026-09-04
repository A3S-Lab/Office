import JSZip from 'jszip';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  descendants,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STORY_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const MAX_MOVE_REVISION_PATCHES = 65_536;
const MAX_MOVE_TEXT_LENGTH = 1_000_000;

export type DocxMoveRevisionRole = 'from' | 'to';

export interface DocxMoveRevisionRegistration {
  kind: 'move-from' | 'move-to';
  wireId: number;
  id: number;
  author: string;
  date: string;
}

export interface DocxMoveRevisionPatch {
  wireId: number;
  id: number;
  role: DocxMoveRevisionRole;
  author: string;
  date: string;
}

/**
 * Allocates transient revision IDs while the `docx` package creates its
 * ordinary `w:ins`/`w:del` wrappers. The post-pack patcher changes those
 * wrappers to native `w:moveFrom`/`w:moveTo` elements and restores the shared
 * logical revision ID. Negative wire IDs cannot collide with the positive
 * IDs emitted for ordinary tracked changes.
 */
export class DocxMoveRevisionPatchCollector {
  readonly patches: DocxMoveRevisionPatch[] = [];
  private readonly registrations = new WeakMap<
    HTMLElement,
    DocxMoveRevisionRegistration
  >();
  private nextWireId = -1;

  register(
    element: HTMLElement,
    id: number,
    revisionDate = normalizedRevisionDate(element.dataset.changeDate),
  ): DocxMoveRevisionRegistration | null {
    if (
      !element.hasAttribute('data-document-change') ||
      element.dataset.changeKind !== 'move'
    ) {
      return null;
    }
    const existing = this.registrations.get(element);
    if (existing) return existing;
    const role = moveRole(element.dataset.changeMoveRole);
    const author = element.dataset.changeAuthor?.trim() ?? '';
    const date = revisionDate;
    const key = element.dataset.changeId?.trim() ?? '';
    const text = element.textContent ?? '';
    if (
      !role ||
      !key ||
      key.length > 255 ||
      /[\u0000-\u001f\u007f]/.test(key) ||
      !Number.isSafeInteger(id) ||
      id < 0 ||
      !author ||
      author.length > 255 ||
      /[\u0000-\u001f\u007f]/.test(author) ||
      !text ||
      text.length > MAX_MOVE_TEXT_LENGTH ||
      element.querySelector('[data-document-change]') ||
      element.querySelector(
        '[data-document-equation], [data-document-field], [data-document-note-reference], [data-document-content-control], img, br',
      )
    ) {
      throw new Error('Document contains an invalid move revision.');
    }
    if (this.patches.length >= MAX_MOVE_REVISION_PATCHES) {
      throw new Error('Document exceeds the move revision limit.');
    }
    const wireId = this.nextWireId;
    this.nextWireId -= 1;
    const registration: DocxMoveRevisionRegistration = {
      kind: role === 'from' ? 'move-from' : 'move-to',
      wireId,
      id,
      author,
      date,
    };
    this.registrations.set(element, registration);
    this.patches.push({ wireId, id, role, author, date });
    return registration;
  }
}

/**
 * Converts the transient wrappers emitted by `docx` into native Word move
 * revisions. Every registered side must be emitted at least once; a missing
 * side is treated as a programming or package-generation error rather than
 * silently producing a half-move that Word would interpret differently.
 */
export async function patchDocxMoveRevisions(
  buffer: ArrayBuffer,
  patches: readonly DocxMoveRevisionPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_MOVE_REVISION_PATCHES) {
    throw new Error('Document exceeds the move revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const byWireId = new Map<number, DocxMoveRevisionPatch>();
  const logicalSides = new Map<
    string,
    { from: DocxMoveRevisionPatch[]; to: DocxMoveRevisionPatch[] }
  >();
  for (const patch of patches) {
    if (byWireId.has(patch.wireId)) {
      throw new Error('Generated DOCX move revision wire IDs are duplicated.');
    }
    byWireId.set(patch.wireId, patch);
    const logicalKey = `${patch.id}\u0000${patch.author}\u0000${patch.date}`;
    const sides = logicalSides.get(logicalKey) ?? { from: [], to: [] };
    sides[patch.role].push(patch);
    logicalSides.set(logicalKey, sides);
  }
  for (const sides of logicalSides.values()) {
    if (!sides.from.length || !sides.to.length) {
      throw new Error(
        'A DOCX move revision must contain both moveFrom and moveTo sides.',
      );
    }
  }
  const applied = new Map<number, number>();
  const emittedText = new Map<number, string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !STORY_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changed = false;
    const wrappers = [
      ...descendants(document, 'ins'),
      ...descendants(document, 'del'),
    ];
    for (const wrapper of wrappers) {
      if (!wrapper.parentNode) continue;
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(wrapper.namespaceURI ?? '')) {
        continue;
      }
      const wireId = numericWordAttribute(wrapper, 'id');
      if (wireId === null) continue;
      const patch = byWireId.get(wireId);
      if (!patch) continue;
      const expectedName = patch.role === 'from' ? 'del' : 'ins';
      if (wrapper.localName !== expectedName) {
        throw new Error(
          `Generated DOCX move revision ${patch.wireId} has the wrong wrapper kind.`,
        );
      }
      const replacement = document.createElementNS(
        wrapper.namespaceURI,
        `${xmlNamespacePrefix(wrapper, wrapper.namespaceURI) ?? 'w'}:${
          patch.role === 'from' ? 'moveFrom' : 'moveTo'
        }`,
      );
      for (const child of Array.from(wrapper.attributes)) {
        const namespace = xmlAttributeNamespace(wrapper, child);
        if (namespace) {
          replacement.setAttributeNS(namespace, child.name, child.value);
        } else {
          replacement.setAttribute(child.name, child.value);
        }
      }
      setWordAttribute(replacement, 'id', String(patch.id));
      setWordAttribute(replacement, 'author', patch.author);
      setWordAttribute(replacement, 'date', patch.date);
      const text = Array.from(wrapper.querySelectorAll('*'))
        .filter((element) =>
          patch.role === 'from'
            ? element.localName === 'delText'
            : element.localName === 't',
        )
        .map((element) => element.textContent ?? '')
        .join('');
      emittedText.set(
        patch.wireId,
        `${emittedText.get(patch.wireId) ?? ''}${text}`,
      );
      while (wrapper.firstChild) replacement.append(wrapper.firstChild);
      wrapper.replaceWith(replacement);
      applied.set(patch.wireId, (applied.get(patch.wireId) ?? 0) + 1);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.wireId));
  if (missing.length) {
    throw new Error(
      `DOCX move revision markers were not emitted: ${missing
        .map((patch) => patch.wireId)
        .join(', ')}.`,
    );
  }
  for (const sides of logicalSides.values()) {
    const from = sides.from
      .map((patch) => emittedText.get(patch.wireId) ?? '')
      .join('');
    const to = sides.to
      .map((patch) => emittedText.get(patch.wireId) ?? '')
      .join('');
    if (!from || !to || from !== to) {
      throw new Error(
        'A DOCX move revision source and destination must contain the same text.',
      );
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function moveRole(value: string | undefined): DocxMoveRevisionRole | null {
  return value === 'from' || value === 'to' ? value : null;
}

function numericWordAttribute(
  element: Element,
  localName: string,
): number | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === namespace,
  );
  if (attributes.length !== 1) return null;
  const value = Number(attributes[0]?.value);
  return Number.isSafeInteger(value) ? value : null;
}

function setWordAttribute(
  element: Element,
  localName: string,
  value: string,
): void {
  const namespace = element.namespaceURI ?? WORD_NAMESPACE;
  const prefix =
    xmlNamespacePrefix(element, namespace) ?? element.prefix ?? 'w';
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function normalizedRevisionDate(value: string | undefined): string {
  const time = Date.parse(value ?? '');
  return Number.isFinite(time)
    ? new Date(time).toISOString()
    : new Date().toISOString();
}
