import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { preserveDocxNoteCommentContentControls } from '../src/internal/features/work/work-docx-note-comment-content-control-preservation';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
import type {
  WorkDocumentComment,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-content-control';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

describe('DOCX note and comment content-control preservation', () => {
  test('restores a static inline comment control with safe metadata and runs', async () => {
    const content = commentContent('Alpha controlled tail');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Review field',
      id: '7401',
      lock: 'sdtLocked',
      selectedText: 'controlled',
      tag: 'review-field',
      type: 'richText',
    });

    const output = await exportWithSource(content, source);
    const comment = elementById(
      await xmlEntry(output, 'word/comments.xml'),
      'comment',
      '42',
    );
    const control = controlContaining(comment, 'controlled');
    const properties = directChild(control, 'sdtPr');
    expect(wordValue(directChild(properties ?? control, 'alias'))).toBe(
      'Review field',
    );
    expect(wordValue(directChild(properties ?? control, 'tag'))).toBe(
      'review-field',
    );
    expect(wordValue(directChild(properties ?? control, 'id'))).toBe('7401');
    expect(wordValue(directChild(properties ?? control, 'lock'))).toBe(
      'sdtLocked',
    );
    expect(directChild(properties ?? control, 'richText')).toBeDefined();
    expect(descendants(control, 'smallCaps')).toHaveLength(2);
    expect(vendorAttribute(control, 'token')).toBe('wrapper');
    expect(vendorAttribute(properties, 'token')).toBe('properties');
    expect(vendorAttribute(directChild(control, 'sdtContent'), 'token')).toBe(
      'content',
    );
    expect(descendants(control, 'passiveControl')).toHaveLength(3);
    expect(descendants(control, 'relationshipBound')).toHaveLength(0);
    expect(
      directChildren(descendants(comment, 'p')[0], 'r').map(runText),
    ).toEqual(['Alpha ', ' tail']);
  });

  test('restores a block control around stable note paragraphs', async () => {
    const content = noteContent(
      '<p>First controlled paragraph</p><p>Second controlled paragraph</p><p>Outside</p>',
    );
    const source = await seedArchive(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const controlledParagraphs = descendants(note, 'p').filter((paragraph) =>
      ['First controlled paragraph', 'Second controlled paragraph'].includes(
        textOf(paragraph),
      ),
    );
    wrapBlockControl(footnotes, note, controlledParagraphs, {
      alias: 'Address block',
      id: '7402',
      lock: 'contentLocked',
      tag: 'address-block',
      type: 'richText',
    });
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const outputNote = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/footnotes.xml',
      ),
      'footnote',
      '27',
    );
    const control = directChildren(outputNote, 'sdt')[0];
    expect(control).toBeDefined();
    expect(
      directChildren(directChild(control, 'sdtContent') ?? control, 'p').map(
        textOf,
      ),
    ).toEqual(['First controlled paragraph', 'Second controlled paragraph']);
    expect(wordValue(directChild(directChild(control, 'sdtPr'), 'id'))).toBe(
      '7402',
    );
    expect(vendorAttribute(control, 'token')).toBe('wrapper');
    expect(
      directChildren(outputNote, 'p').some(
        (paragraph) => textOf(paragraph) === 'Outside',
      ),
    ).toBe(true);
  });

  test('retains a valid plain-text control and its multiline declaration', async () => {
    const content = noteContent('<p>Plain control</p>', 'endnote');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/endnotes.xml', 'endnote', '41', {
      alias: 'Plain field',
      id: '7403',
      selectedText: 'Plain control',
      tag: 'plain-field',
      type: 'text',
    });

    const output = await exportWithSource(content, source);
    const control = controlContaining(
      elementById(await xmlEntry(output, 'word/endnotes.xml'), 'endnote', '41'),
      'Plain control',
    );
    const textProperty = directChild(
      directChild(control, 'sdtPr') ?? control,
      'text',
    );
    expect(textProperty).toBeDefined();
    expect(wordAttribute(textProperty, 'multiLine')).toBe('1');
    expect(descendants(control, 'r')).toHaveLength(1);
  });

  test('restores a control on the correct reply after thread reordering', async () => {
    const sourceContent = threadedCommentContent(false);
    const source = await seedArchive(sourceContent);
    await addInlineControl(source, 'word/comments.xml', 'comment', '91', {
      alias: 'Reply field',
      id: '7404',
      selectedText: 'Reply control',
      tag: 'reply-field',
      type: 'richText',
    });

    const output = await exportWithSource(threadedCommentContent(true), source);
    const comments = await xmlEntry(output, 'word/comments.xml');
    expect(
      controlContaining(
        elementById(comments, 'comment', '91'),
        'Reply control',
      ),
    ).toBeDefined();
    expect(
      descendants(elementById(comments, 'comment', '7'), 'sdt'),
    ).toHaveLength(0);
  });

  test('drops a content-control wrapper when its visible text changes', async () => {
    const source = await seedArchive(commentContent('Original control'));
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Stale field',
      id: '7405',
      selectedText: 'control',
      tag: 'stale-field',
      type: 'richText',
    });

    const output = await exportWithSource(
      commentContent('Edited control text'),
      source,
    );
    expect(
      descendants(
        elementById(
          await xmlEntry(output, 'word/comments.xml'),
          'comment',
          '42',
        ),
        'sdt',
      ),
    ).toHaveLength(0);
  });

  test('restores multiple disjoint controls and accepts default rich text', async () => {
    const content = commentContent('First field and second field');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'First field',
      id: '7406',
      selectedText: 'First field',
      tag: 'first-field',
      type: 'richText',
    });
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Second field',
      id: '7407',
      selectedText: 'second field',
      tag: 'second-field',
      type: 'richText',
    });
    const comments = await xmlEntry(source, 'word/comments.xml');
    const second = controlContaining(
      elementById(comments, 'comment', '42'),
      'second field',
    );
    directChild(directChild(second, 'sdtPr') ?? second, 'richText')?.remove();
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const outputComment = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/comments.xml',
      ),
      'comment',
      '42',
    );
    const controls = descendants(outputComment, 'sdt');
    expect(controls.map(textOf)).toEqual(['First field', 'second field']);
    expect(
      controls.map((control) =>
        wordValue(directChild(directChild(control, 'sdtPr'), 'id')),
      ),
    ).toEqual(['7406', '7407']);
    expect(
      directChild(directChild(controls[1], 'sdtPr') ?? controls[1], 'richText'),
    ).toBeUndefined();
  });

  test('retains validated Word 2013 appearance and declares it ignorable', async () => {
    const content = commentContent('Colored control');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Colored field',
      id: '7408',
      selectedText: 'Colored control',
      tag: 'colored-field',
      type: 'richText',
    });
    const comments = await xmlEntry(source, 'word/comments.xml');
    const properties = directChild(
      controlContaining(
        elementById(comments, 'comment', '42'),
        'Colored control',
      ),
      'sdtPr',
    );
    if (!properties) throw new Error('Missing colored properties.');
    for (const [localName, value] of [
      ['appearance', 'tags'],
      ['color', '2F73D9'],
    ] as const) {
      const property = comments.createElementNS(
        WORD_2012_NAMESPACE,
        `w15:${localName}`,
      );
      property.setAttributeNS(WORD_2012_NAMESPACE, 'w15:val', value);
      properties.append(property);
    }
    comments.documentElement.setAttributeNS(
      'http://www.w3.org/2000/xmlns/',
      'xmlns:w15',
      WORD_2012_NAMESPACE,
    );
    const ignorable = attribute(comments.documentElement, 'mc:Ignorable');
    comments.documentElement.setAttributeNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Ignorable',
      `${ignorable ?? ''} w15`.trim(),
    );
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const output = await exportWithSource(content, source);
    const outputComments = await xmlEntry(output, 'word/comments.xml');
    const outputProperties = directChild(
      controlContaining(
        elementById(outputComments, 'comment', '42'),
        'Colored control',
      ),
      'sdtPr',
    );
    const appearance = Array.from(outputProperties?.children ?? []).find(
      (element) =>
        element.localName === 'appearance' &&
        element.namespaceURI === WORD_2012_NAMESPACE,
    );
    const color = Array.from(outputProperties?.children ?? []).find(
      (element) =>
        element.localName === 'color' &&
        element.namespaceURI === WORD_2012_NAMESPACE,
    );
    expect(namespacedValue(appearance, WORD_2012_NAMESPACE)).toBe('tags');
    expect(namespacedValue(color, WORD_2012_NAMESPACE)).toBe('2F73D9');
    const outputIgnorable = attribute(
      outputComments.documentElement,
      'mc:Ignorable',
    );
    const prefix = appearance?.prefix;
    expect(prefix).toBeTruthy();
    expect(outputIgnorable?.split(/\s+/u)).toContain(prefix);
    expect(outputIgnorable?.split(/\s+/u)).toContain('vendor');
  });

  test('rewrites only colliding IDs while reserving stable source IDs', () => {
    const generated = parseXml(
      `<w:comments xmlns:w="${WORD_NAMESPACE}"><w:comment w:id="42"><w:p><w:r><w:t>Alpha Beta</w:t></w:r></w:p></w:comment><w:comment w:id="43"><w:p><w:sdt><w:sdtPr><w:id w:val="7419"/></w:sdtPr><w:sdtContent><w:r><w:t>Existing</w:t></w:r></w:sdtContent></w:sdt></w:p></w:comment></w:comments>`,
      'generated comments',
    );
    const source = parseXml(
      `<w:comments xmlns:w="${WORD_NAMESPACE}"><w:comment w:id="42"><w:p><w:sdt><w:sdtPr><w:id w:val="7419"/></w:sdtPr><w:sdtContent><w:r><w:t>Alpha</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:r><w:t>Beta</w:t></w:r></w:sdtContent></w:sdt></w:p></w:comment></w:comments>`,
      'source comments',
    );
    const generatedComment = elementById(generated, 'comment', '42');
    const sourceComment = elementById(source, 'comment', '42');
    preserveDocxNoteCommentContentControls(
      generated,
      source,
      [{ generated: generatedComment, source: sourceComment }],
      'comment',
    );

    const controls = descendants(generatedComment, 'sdt');
    expect(controls.map(textOf)).toEqual(['Alpha', 'Beta']);
    expect(
      controls.map((control) =>
        wordValue(directChild(directChild(control, 'sdtPr'), 'id')),
      ),
    ).toEqual(['2', '1']);
    expect(
      wordValue(
        directChild(
          directChild(
            controlContaining(
              elementById(generated, 'comment', '43'),
              'Existing',
            ),
            'sdtPr',
          ),
          'id',
        ),
      ),
    ).toBe('7419');
  });
});

describe('DOCX note and comment content-control boundaries', () => {
  test('translates static controls from strict UTF-16 source XML', async () => {
    const content = noteContent('<p>Strict controlled value</p>');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/footnotes.xml', 'footnote', '27', {
      alias: 'Strict field',
      id: '7410',
      selectedText: 'controlled',
      tag: 'strict-field',
      type: 'richText',
    });
    const xml = await source.file('word/footnotes.xml')?.async('string');
    if (!xml) throw new Error('Missing source footnotes.');
    source.file(
      'word/footnotes.xml',
      utf16Xml(
        xml.replaceAll(
          WORD_NAMESPACE,
          'http://purl.oclc.org/ooxml/wordprocessingml/main',
        ),
      ),
    );

    const output = await exportWithSource(content, source);
    const control = controlContaining(
      elementById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      ),
      'controlled',
    );
    expect(control.namespaceURI).toBe(WORD_NAMESPACE);
    expect(wordValue(directChild(directChild(control, 'sdtPr'), 'id'))).toBe(
      '7410',
    );
  });

  test('fails a paragraph atomically for active binding and placeholder state', async () => {
    const content = commentContent('Safe field and bound field');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Safe field',
      id: '7411',
      selectedText: 'Safe field',
      tag: 'safe-field',
      type: 'richText',
    });
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Bound field',
      id: '7412',
      selectedText: 'bound field',
      tag: 'bound-field',
      type: 'richText',
    });
    const comments = await xmlEntry(source, 'word/comments.xml');
    const bound = controlContaining(
      elementById(comments, 'comment', '42'),
      'bound field',
    );
    const properties = directChild(bound, 'sdtPr');
    if (!properties) throw new Error('Missing bound properties.');
    const binding = wordElement(comments, 'dataBinding');
    binding.setAttributeNS(WORD_NAMESPACE, 'w:storeItemID', '{BAD-BINDING}');
    binding.setAttributeNS(WORD_NAMESPACE, 'w:xpath', '/root/value');
    properties.append(binding, wordElement(comments, 'showingPlcHdr'));
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const outputComment = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/comments.xml',
      ),
      'comment',
      '42',
    );
    expect(descendants(outputComment, 'sdt')).toHaveLength(0);
    expect(textOf(outputComment)).toContain('Safe field and bound field');
  });

  test('rejects duplicate IDs without suppressing an unrelated control', async () => {
    const content = noteContent(
      '<p>Duplicate one</p><p>Duplicate two</p><p>Unique three</p>',
    );
    const source = await seedArchive(content);
    for (const [selectedText, id] of [
      ['Duplicate one', '7413'],
      ['Duplicate two', '7413'],
      ['Unique three', '7414'],
    ] as const) {
      await addInlineControl(source, 'word/footnotes.xml', 'footnote', '27', {
        alias: selectedText,
        id,
        selectedText,
        tag: selectedText.toLowerCase().replaceAll(' ', '-'),
        type: 'richText',
      });
    }

    const note = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/footnotes.xml',
      ),
      'footnote',
      '27',
    );
    expect(descendants(note, 'sdt').map(textOf)).toEqual(['Unique three']);
    expect(
      wordValue(
        directChild(directChild(descendants(note, 'sdt')[0], 'sdtPr'), 'id'),
      ),
    ).toBe('7414');
  });

  test('does not guess which duplicate paragraph owned an inline control', async () => {
    const content = noteContent('<p>Repeat</p><p>Repeat</p>');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/footnotes.xml', 'footnote', '27', {
      alias: 'Ambiguous field',
      id: '7415',
      selectedText: 'Repeat',
      tag: 'ambiguous-field',
      type: 'richText',
    });

    const note = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/footnotes.xml',
      ),
      'footnote',
      '27',
    );
    expect(descendants(note, 'sdt')).toHaveLength(0);
  });

  test('rejects block controls containing tables or mixed semantic content', async () => {
    const content = noteContent('<p>Block text</p><p>Outside</p>');
    const source = await seedArchive(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const paragraph = descendants(note, 'p').find(
      (item) => textOf(item) === 'Block text',
    );
    if (!paragraph) throw new Error('Missing block paragraph.');
    wrapBlockControl(footnotes, note, [paragraph], {
      alias: 'Mixed block',
      id: '7416',
      tag: 'mixed-block',
      type: 'richText',
    });
    const control = directChildren(note, 'sdt')[0];
    const contentElement = directChild(control, 'sdtContent');
    if (!contentElement) throw new Error('Missing block content.');
    const table = wordElement(footnotes, 'tbl');
    const row = wordElement(footnotes, 'tr');
    const cell = wordElement(footnotes, 'tc');
    const tableParagraph = wordElement(footnotes, 'p');
    tableParagraph.append(cloneTextRun(descendants(paragraph, 'r')[0], 'Cell'));
    cell.append(tableParagraph);
    row.append(cell);
    table.append(row);
    contentElement.append(table);
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const outputNote = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/footnotes.xml',
      ),
      'footnote',
      '27',
    );
    expect(descendants(outputNote, 'sdt')).toHaveLength(0);
    expect(descendants(outputNote, 'tbl')).toHaveLength(0);
  });

  test('rejects malformed, namespace-spoofed, and relationship-bound properties', async () => {
    const cases = [
      'duplicate-alias',
      'spoofed-type',
      'relationship-bound',
    ] as const;
    for (const testCase of cases) {
      const content = commentContent(`Case ${testCase}`);
      const source = await seedArchive(content);
      await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
        alias: 'Boundary field',
        id: '7417',
        selectedText: testCase,
        tag: 'boundary-field',
        type: 'richText',
      });
      const comments = await xmlEntry(source, 'word/comments.xml');
      const control = controlContaining(
        elementById(comments, 'comment', '42'),
        testCase,
      );
      const properties = directChild(control, 'sdtPr');
      if (!properties) throw new Error('Missing boundary properties.');
      if (testCase === 'duplicate-alias') {
        properties.append(
          wordValueElement(comments, 'alias', 'Duplicate alias'),
        );
      } else if (testCase === 'spoofed-type') {
        directChild(properties, 'richText')?.remove();
        properties.append(
          comments.createElementNS(RELATIONSHIP_NAMESPACE, 'r:richText'),
        );
      } else {
        directChild(properties, 'tag')?.setAttributeNS(
          RELATIONSHIP_NAMESPACE,
          'r:id',
          'rIdUnsafeProperty',
        );
      }
      source.file(
        'word/comments.xml',
        new XMLSerializer().serializeToString(comments),
      );

      const output = await exportWithSource(content, source);
      expect(
        descendants(
          elementById(
            await xmlEntry(output, 'word/comments.xml'),
            'comment',
            '42',
          ),
          'sdt',
        ),
      ).toHaveLength(0);
    }
  });

  test('rejects a plain-text declaration whose cache has multiple runs', async () => {
    const content = commentContent('Plain split');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Plain field',
      id: '7418',
      selectedText: 'Plain split',
      tag: 'plain-field',
      type: 'text',
    });
    const comments = await xmlEntry(source, 'word/comments.xml');
    const control = controlContaining(
      elementById(comments, 'comment', '42'),
      'Plain split',
    );
    const contentElement = directChild(control, 'sdtContent');
    const original = descendants(contentElement ?? control, 'r')[0];
    original.remove();
    contentElement?.append(
      cloneTextRun(original, 'Plain '),
      cloneTextRun(original, 'split'),
    );
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const output = await exportWithSource(content, source);
    expect(
      descendants(
        elementById(
          await xmlEntry(output, 'word/comments.xml'),
          'comment',
          '42',
        ),
        'sdt',
      ),
    ).toHaveLength(0);
  });

  test('rejects nested controls rather than reconnecting an ambiguous cache', async () => {
    const content = commentContent('Outer inner value');
    const source = await seedArchive(content);
    await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
      alias: 'Outer field',
      id: '7420',
      selectedText: 'Outer inner value',
      tag: 'outer-field',
      type: 'richText',
    });
    const comments = await xmlEntry(source, 'word/comments.xml');
    const outer = controlContaining(
      elementById(comments, 'comment', '42'),
      'Outer inner value',
    );
    const outerContent = directChild(outer, 'sdtContent');
    const outerRun = descendants(outerContent ?? outer, 'r')[0];
    const nested = createControl(comments, {
      alias: 'Nested field',
      id: '7421',
      selectedText: 'inner',
      tag: 'nested-field',
      type: 'richText',
    });
    directChild(nested, 'sdtContent')?.append(cloneTextRun(outerRun, 'inner'));
    outerRun.remove();
    outerContent?.append(
      cloneTextRun(outerRun, 'Outer '),
      nested,
      cloneTextRun(outerRun, ' value'),
    );
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const outputComment = elementById(
      await xmlEntry(
        await exportWithSource(content, source),
        'word/comments.xml',
      ),
      'comment',
      '42',
    );
    expect(descendants(outputComment, 'sdt')).toHaveLength(0);
  });

  test('rejects controls mixed with hyperlinks, math, drawings, or form types', async () => {
    for (const semantic of ['hyperlink', 'oMath', 'drawing', 'date'] as const) {
      const content = commentContent(`Mixed ${semantic}`);
      const source = await seedArchive(content);
      await addInlineControl(source, 'word/comments.xml', 'comment', '42', {
        alias: 'Mixed field',
        id: '7422',
        selectedText: `Mixed ${semantic}`,
        tag: 'mixed-field',
        type: 'richText',
      });
      const comments = await xmlEntry(source, 'word/comments.xml');
      const control = controlContaining(
        elementById(comments, 'comment', '42'),
        `Mixed ${semantic}`,
      );
      if (semantic === 'date') {
        const properties = directChild(control, 'sdtPr');
        directChild(properties ?? control, 'richText')?.remove();
        properties?.append(wordElement(comments, 'date'));
      } else {
        directChild(control, 'sdtContent')?.append(
          wordElement(comments, semantic),
        );
      }
      source.file(
        'word/comments.xml',
        new XMLSerializer().serializeToString(comments),
      );

      const output = await exportWithSource(content, source);
      expect(
        descendants(
          elementById(
            await xmlEntry(output, 'word/comments.xml'),
            'comment',
            '42',
          ),
          'sdt',
        ),
      ).toHaveLength(0);
    }
  });
});

interface StaticControlOptions {
  alias: string;
  id: string;
  lock?: 'contentLocked' | 'sdtContentLocked' | 'sdtLocked' | 'unlocked';
  selectedText: string;
  tag: string;
  type: 'richText' | 'text';
}

async function addInlineControl(
  archive: JSZip,
  path: string,
  itemName: string,
  itemId: string,
  options: StaticControlOptions,
): Promise<void> {
  const document = await xmlEntry(archive, path);
  const item = elementById(document, itemName, itemId);
  const run = descendants(item, 'r').find((candidate) =>
    runText(candidate).includes(options.selectedText),
  );
  if (!run?.parentElement || run.parentElement.localName !== 'p') {
    throw new Error(`Missing direct run for ${options.selectedText}.`);
  }
  const paragraph = run.parentElement;
  const text = runText(run);
  const start = text.indexOf(options.selectedText);
  const before = text.slice(0, start);
  const after = text.slice(start + options.selectedText.length);
  const control = createControl(document, options);
  const selectedRun = cloneTextRun(run, options.selectedText);
  const runProperties = wordElement(document, 'rPr');
  runProperties.append(wordElement(document, 'smallCaps'));
  selectedRun.insertBefore(runProperties, selectedRun.firstChild);
  directChild(control, 'sdtContent')?.append(selectedRun);
  if (before) paragraph.insertBefore(cloneTextRun(run, before), run);
  paragraph.insertBefore(control, run);
  if (after) paragraph.insertBefore(cloneTextRun(run, after), run);
  run.remove();
  archive.file(path, new XMLSerializer().serializeToString(document));
}

function wrapBlockControl(
  document: Document,
  owner: Element,
  paragraphs: readonly Element[],
  options: Omit<StaticControlOptions, 'selectedText'>,
): void {
  const first = paragraphs[0];
  if (!first || first.parentElement !== owner) {
    throw new Error('Missing direct block-control paragraph.');
  }
  const control = createControl(document, {
    ...options,
    selectedText: paragraphs.map(textOf).join('\n'),
  });
  owner.insertBefore(control, first);
  const content = directChild(control, 'sdtContent');
  if (!content) throw new Error('Missing content-control content.');
  for (const paragraph of paragraphs) content.append(paragraph);
}

function createControl(
  document: Document,
  options: StaticControlOptions,
): Element {
  declareVendor(document.documentElement);
  const control = wordElement(document, 'sdt');
  const properties = wordElement(document, 'sdtPr');
  properties.append(
    wordValueElement(document, 'alias', options.alias),
    wordValueElement(document, 'tag', options.tag),
    wordValueElement(document, 'id', options.id),
  );
  if (options.lock) {
    properties.append(wordValueElement(document, 'lock', options.lock));
  }
  const type = wordElement(document, options.type);
  if (options.type === 'text') {
    type.setAttributeNS(WORD_NAMESPACE, 'w:multiLine', '1');
  }
  properties.append(type);
  const endProperties = wordElement(document, 'sdtEndPr');
  const endRunProperties = wordElement(document, 'rPr');
  endRunProperties.append(wordElement(document, 'smallCaps'));
  endProperties.append(endRunProperties);
  const content = wordElement(document, 'sdtContent');
  decorate(control, document, 'wrapper');
  decorate(properties, document, 'properties');
  decorate(content, document, 'content');
  control.append(properties, endProperties, content);
  return control;
}

function decorate(element: Element, document: Document, token: string): void {
  element.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  const passive = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:passiveControl',
  );
  passive.textContent = token;
  element.append(passive);
  const relationshipBound = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:relationshipBound',
  );
  relationshipBound.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rIdUnsafe');
  element.append(relationshipBound);
}

function declareVendor(root: Element): void {
  root.setAttributeNS(
    'http://www.w3.org/2000/xmlns/',
    'xmlns:vendor',
    VENDOR_NAMESPACE,
  );
  root.setAttributeNS(
    'http://www.w3.org/2000/xmlns/',
    'xmlns:mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  root.setAttributeNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Ignorable', 'vendor');
}

function cloneTextRun(run: Element, text: string): Element {
  const clone = run.cloneNode(true) as Element;
  for (const child of Array.from(clone.children)) {
    if (child.localName !== 'rPr') child.remove();
  }
  const textElement = wordElement(clone.ownerDocument, 't');
  textElement.setAttributeNS(
    'http://www.w3.org/XML/1998/namespace',
    'xml:space',
    'preserve',
  );
  textElement.textContent = text;
  clone.append(textElement);
  return clone;
}

function noteContent(
  html: string,
  kind: 'endnote' | 'footnote' = 'footnote',
): WorkDocumentContent {
  const nativeId = kind === 'footnote' ? '27' : '41';
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      `<p>Body<sup data-document-note-reference="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">1</sup></p>`,
      `<aside data-document-note="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">${html}</aside>`,
      '</section>',
    ].join(''),
  };
}

function commentContent(text: string): WorkDocumentContent {
  return documentWithComments([
    {
      id: 'docx-comment-42',
      author: 'Reviewer',
      date: '2026-08-12T00:00:00.000Z',
      text,
      resolved: false,
    },
  ]);
}

function threadedCommentContent(reordered: boolean): WorkDocumentContent {
  const root: WorkDocumentComment = {
    id: 'docx-comment-7',
    author: 'Root',
    date: '2026-08-12T00:00:00.000Z',
    text: 'Root comment',
    resolved: false,
    replies: [
      {
        id: 'docx-comment-reply-91',
        author: 'Reply',
        date: '2026-08-12T00:01:00.000Z',
        text: 'Reply control',
      },
    ],
  };
  const other: WorkDocumentComment = {
    id: 'docx-comment-8',
    author: 'Other',
    date: '2026-08-12T00:02:00.000Z',
    text: 'Other comment',
    resolved: false,
  };
  return documentWithComments(reordered ? [other, root] : [root, other]);
}

function documentWithComments(
  comments: WorkDocumentComment[],
): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: `<section data-document-section="true"><p><span data-document-comment="true" data-comment-id="${comments.find((item) => item.id === 'docx-comment-42')?.id ?? 'docx-comment-7'}">Anchor</span></p></section>`,
    comments,
  };
}

async function seedArchive(content: WorkDocumentContent): Promise<JSZip> {
  return JSZip.loadAsync(await (await createDocxBlob(content)).arrayBuffer());
}

async function exportWithSource(
  content: WorkDocumentContent,
  source: JSZip,
): Promise<JSZip> {
  const output = await createDocxBlob(
    content,
    await source.generateAsync({ type: 'arraybuffer' }),
  );
  return JSZip.loadAsync(await output.arrayBuffer());
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const bytes = await archive.file(path)?.async('uint8array');
  if (!bytes) throw new Error(`Missing ${path}.`);
  return parseXml(new TextDecoder().decode(bytes), path);
}

function elementById(
  document: Document,
  localName: string,
  id: string,
): Element {
  const element = descendants(document, localName).find(
    (candidate) => attribute(candidate, 'id') === id,
  );
  if (!element) throw new Error(`Missing ${localName} ${id}.`);
  return element;
}

function controlContaining(scope: Element, text: string): Element {
  const control = descendants(scope, 'sdt').find(
    (candidate) => textOf(candidate) === text,
  );
  if (!control) throw new Error(`Missing content control ${text}.`);
  return control;
}

function runText(run: Element): string {
  return directChildren(run, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

function textOf(element: Element): string {
  return descendants(element, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

function wordElement(document: Document, localName: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
}

function wordValueElement(
  document: Document,
  localName: string,
  value: string,
): Element {
  const element = wordElement(document, localName);
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', value);
  return element;
}

function wordValue(element: Element | undefined): string | null {
  return wordAttribute(element, 'val');
}

function namespacedValue(
  element: Element | undefined,
  namespace: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === 'val' &&
        xmlAttributeNamespace(element, item) === namespace,
    )?.value ?? null
  );
}

function wordAttribute(
  element: Element | undefined,
  localName: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === WORD_NAMESPACE,
    )?.value ?? null
  );
}

function vendorAttribute(
  element: Element | undefined,
  localName: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

function utf16Xml(source: string): Uint8Array {
  const normalized = source.replace(/^<\?xml[^>]*\?>/, '').replace(/^\s+/, '');
  const value = `<?xml version="1.0" encoding="UTF-16"?>${normalized}`;
  const bytes = new Uint8Array(2 + value.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes[2 + index * 2] = code & 0xff;
    bytes[3 + index * 2] = code >> 8;
  }
  return bytes;
}
