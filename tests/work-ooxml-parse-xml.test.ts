import { expect, test } from '@rstest/core';
import {
  attribute,
  descendants,
  neutralizePrefixedAttributeCollisions,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

test('neutralizePrefixedAttributeCollisions moves r:id ahead of bare id', () => {
  expect(
    neutralizePrefixedAttributeCollisions(
      '<p:sldId id="256" foo="x" r:id="rId2"/>',
    ),
  ).toBe('<p:sldId r:id="rId2" id="256" foo="x"/>');
});

test('parseXml preserves relationship r:id when bare id precedes it', () => {
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
     <p:presentation
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
       <p:sldIdLst>
         <p:sldId id="256" r:id="rId2"/>
         <p:sldId id="257" foo="x" r:id="rId3"/>
       </p:sldIdLst>
     </p:presentation>`,
    'presentation.xml',
  );

  const slideIds = descendants(document, 'sldId');
  expect(slideIds).toHaveLength(2);
  expect(attribute(slideIds[0], 'id')).toBe('256');
  expect(attribute(slideIds[0], 'r:id')).toBe('rId2');
  expect(attribute(slideIds[1], 'r:id')).toBe('rId3');
});

test('parseXml rejects malformed XML with a labeled error', () => {
  expect(() => parseXml('<', 'broken.xml')).toThrow(/broken\.xml/);
});
