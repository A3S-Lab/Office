import { Buffer } from 'node:buffer';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import JSZip from 'jszip';

const REVISION_AUTHOR = 'WPS Reference';
const REVISION_DATE = '2026-09-05T01:00:00.000Z';
const FIXTURE_DATE = new Date(REVISION_DATE);

export const PARAGRAPH_MARK_INSERTION_TEXT =
  'Inserted whole paragraph from the WPS reference.';
export const PARAGRAPH_MARK_DELETION_TEXT =
  'Deleted whole paragraph from the WPS reference.';
export const PARAGRAPH_MARK_STABLE_TEXT =
  'Stable paragraph outside both native revisions.';

/** Builds the exact text-only paragraph-mark shape emitted by local WPS. */
export async function createWordParagraphMarkRevisionFixture(): Promise<Buffer> {
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic WPS paragraph-mark revision fixture',
    title: 'A3S Office WPS paragraph-mark revision fixture',
    sections: [
      {
        children: [
          new Paragraph({
            children: [new TextRun(PARAGRAPH_MARK_INSERTION_TEXT)],
          }),
          new Paragraph({
            children: [new TextRun(PARAGRAPH_MARK_DELETION_TEXT)],
          }),
          new Paragraph({
            children: [new TextRun(PARAGRAPH_MARK_STABLE_TEXT)],
          }),
        ],
      },
    ],
  });
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document));
  const documentEntry = archive.file('word/document.xml');
  const settingsEntry = archive.file('word/settings.xml');
  if (!documentEntry || !settingsEntry) {
    throw new Error('Generated paragraph-mark fixture is missing package XML.');
  }

  let documentXml = await documentEntry.async('string');
  documentXml = patchWholeParagraphRevision(
    documentXml,
    PARAGRAPH_MARK_INSERTION_TEXT,
    'ins',
    1,
    2,
  );
  documentXml = patchWholeParagraphRevision(
    documentXml,
    PARAGRAPH_MARK_DELETION_TEXT,
    'del',
    3,
    4,
  );
  archive.file('word/document.xml', documentXml, { date: FIXTURE_DATE });

  const settingsXml = await settingsEntry.async('string');
  archive.file(
    'word/settings.xml',
    settingsXml.includes('<w:trackRevisions')
      ? settingsXml
      : settingsXml.replace(
          '</w:settings>',
          '<w:trackRevisions/></w:settings>',
        ),
    { date: FIXTURE_DATE },
  );
  return Buffer.from(
    await archive.generateAsync({
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      type: 'uint8array',
    }),
  );
}

function patchWholeParagraphRevision(
  documentXml: string,
  text: string,
  tag: 'ins' | 'del',
  paragraphMarkId: number,
  bodyId: number,
): string {
  const textElement = `<w:t xml:space="preserve">${text}</w:t>`;
  const run = `<w:r>${textElement}</w:r>`;
  const runOffset = documentXml.indexOf(run);
  if (runOffset < 0) {
    throw new Error(`Generated paragraph-mark fixture is missing '${text}'.`);
  }
  const paragraphStart = documentXml.lastIndexOf('<w:p>', runOffset);
  const paragraphEnd = documentXml.indexOf('</w:p>', runOffset);
  if (paragraphStart < 0 || paragraphEnd < 0) {
    throw new Error(
      'Generated paragraph-mark fixture has an invalid paragraph.',
    );
  }
  const paragraphXml = documentXml.slice(paragraphStart, paragraphEnd + 6);
  if (paragraphXml.includes('<w:pPr>')) {
    throw new Error(
      'Generated paragraph-mark fixture has unexpected properties.',
    );
  }
  const revisionAttributes = `w:author="${REVISION_AUTHOR}" w:date="${REVISION_DATE}"`;
  const paragraphProperties = `<w:pPr><w:rPr><w:${tag} w:id="${paragraphMarkId}" ${revisionAttributes}/></w:rPr></w:pPr>`;
  const revisionText =
    tag === 'del'
      ? textElement
          .replace('<w:t ', '<w:delText ')
          .replace('</w:t>', '</w:delText>')
      : textElement;
  const revisionRun = `<w:r>${revisionText}</w:r>`;
  const patched = paragraphXml
    .replace('<w:p>', `<w:p>${paragraphProperties}`)
    .replace(
      run,
      `<w:${tag} w:id="${bodyId}" ${revisionAttributes}>${revisionRun}</w:${tag}>`,
    );
  if (patched === paragraphXml) {
    throw new Error('Generated paragraph-mark fixture could not be patched.');
  }
  return (
    documentXml.slice(0, paragraphStart) +
    patched +
    documentXml.slice(paragraphEnd + 6)
  );
}
