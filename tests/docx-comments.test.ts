import {
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  Packer,
  Paragraph,
  TextRun,
} from 'docx';
import { expect, test } from '@rstest/core';
import { importOfficeFile } from '../src/core';

test('imports a large native DOCX comment collection without losing anchors', async () => {
  const commentCount = 120;
  const source = new Document({
    comments: {
      children: Array.from({ length: commentCount }, (_, index) => ({
        id: index,
        author: 'A3S Test',
        date: new Date('2026-08-02T00:00:00.000Z'),
        children: [
          new Paragraph({
            children: [new TextRun({ text: `Review comment ${index + 1}.` })],
          }),
        ],
      })),
    },
    sections: [
      {
        children: Array.from(
          { length: commentCount },
          (_, index) =>
            new Paragraph({
              children: [
                new TextRun({ text: `Context ${index + 1}: ` }),
                new CommentRangeStart(index),
                new TextRun({ text: `Comment marker ${index + 1}.` }),
                new CommentRangeEnd(index),
                new CommentReference(index),
              ],
            }),
        ),
      },
    ],
  });
  const buffer = await Packer.toBuffer(source);
  const imported = await importOfficeFile(
    new File([new Uint8Array(buffer)], 'comments.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
  );

  expect(imported.content.type).toBe('document');
  if (imported.content.type !== 'document') return;
  expect(imported.content.comments).toHaveLength(commentCount);
  expect(imported.content.html.match(/data-document-comment=/g)).toHaveLength(
    commentCount,
  );
  expect(imported.content.comments?.at(-1)).toMatchObject({
    author: 'A3S Test',
    text: 'Review comment 120.',
  });
});
