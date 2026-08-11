import { describe, expect, test } from '@rstest/core';
import { diagnoseDocxCaptions } from '../src/internal/features/work/work-docx-caption-diagnostics';
import { markDocxBodyFields } from '../src/internal/features/work/work-docx-field-import';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX body field structure', () => {
  test('accepts complete inline simple and complex fields', () => {
    const document = wordDocument(
      [
        '<w:p>',
        '<w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>DATE \\@ &quot;yyyy-MM-dd&quot;</w:instrText></w:r>',
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        '<w:r><w:t>2026-08-11</w:t></w:r>',
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '</w:p>',
      ].join(''),
    );

    expect(diagnoseDocxCaptions(document)).toMatchObject({
      hasUnsupportedFields: false,
      issues: [expect.objectContaining({ code: 'docx.fields.body' })],
    });
    expect(markDocxBodyFields(document).fields).toHaveLength(2);
  });

  test('reports and avoids atomizing nested fields', () => {
    const document = wordDocument(
      [
        '<w:p>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>PAGE</w:instrText></w:r>',
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>NUMPAGES</w:instrText></w:r>',
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        '<w:r><w:t>12</w:t></w:r>',
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '</w:p>',
      ].join(''),
    );

    const diagnostics = diagnoseDocxCaptions(document);
    expect(diagnostics.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );
    expect(diagnostics.hasUnsupportedFields).toBe(true);
    expect(markDocxBodyFields(document).fields).toHaveLength(0);
  });

  test('reports incomplete, cross-paragraph, deleted, and instructionless fields', () => {
    const document = wordDocument(
      [
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>SECTIONPAGES</w:instrText></w:r></w:p>',
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>TIME</w:instrText></w:r></w:p>',
        '<w:p><w:del><w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText>DATE</w:instrText></w:r>',
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:del></w:p>',
        '<w:p><w:fldSimple><w:r><w:t>Missing</w:t></w:r></w:fldSimple></w:p>',
      ].join(''),
    );

    const diagnostics = diagnoseDocxCaptions(document);
    expect(diagnostics.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );
    expect(diagnostics.hasUnsupportedFields).toBe(true);
    expect(markDocxBodyFields(document).fields).toHaveLength(0);
  });
});

function wordDocument(body: string): Document {
  return new DOMParser().parseFromString(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
    'application/xml',
  );
}
