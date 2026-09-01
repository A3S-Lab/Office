import { describe, expect, test } from '@rstest/core';
import {
  DOCUMENT_OPEN_TYPE_ATTRIBUTE,
  documentOpenTypeDomAttributes,
  documentOpenTypeFeaturesFromElement,
  normalizeDocumentOpenTypeStylisticSets,
  parseDocumentOpenTypeFeatures,
  serializeDocumentOpenTypeFeatures,
} from '../src/internal/features/work/work-document-opentype';

describe('Document OpenType typography', () => {
  test('serializes one canonical bounded feature model', () => {
    const serialized = serializeDocumentOpenTypeFeatures({
      ligatures: 'standardContextual',
      numberForm: 'oldStyle',
      numberSpacing: 'tabular',
      stylisticSets: [4, 1, 4, 20],
      contextualAlternates: false,
    });

    expect(serialized).toBe(
      '{"ligatures":"standardContextual","numberForm":"oldStyle","numberSpacing":"tabular","stylisticSets":[4,1,20],"contextualAlternates":false}',
    );
    expect(parseDocumentOpenTypeFeatures(serialized)).toEqual({
      ligatures: 'standardContextual',
      numberForm: 'oldStyle',
      numberSpacing: 'tabular',
      stylisticSets: [4, 1, 20],
      contextualAlternates: false,
    });
    expect(
      parseDocumentOpenTypeFeatures(
        '{"ligatures":"standardContextual","stylisticSets":[4,4]}',
      ),
    ).toBeNull();
  });

  test('projects the native model through one data attribute and bounded CSS', () => {
    const attributes = documentOpenTypeDomAttributes({
      ligatures: 'standardContextual',
      numberForm: 'lining',
      numberSpacing: 'proportional',
      stylisticSets: [1, 20],
      contextualAlternates: true,
    });

    expect(attributes[DOCUMENT_OPEN_TYPE_ATTRIBUTE]).toContain(
      'standardContextual',
    );
    expect(attributes.style).toContain(
      'font-feature-settings: "liga" 1, "clig" 1, "hlig" 0, "dlig" 0, "ss01" 1, "ss20" 1',
    );
    expect(attributes.style).toContain(
      'font-variant-numeric: lining-nums proportional-nums',
    );
    expect(attributes.style).toContain('font-variant-ligatures: contextual');

    const element = document.createElement('span');
    element.setAttribute(
      DOCUMENT_OPEN_TYPE_ATTRIBUTE,
      attributes[DOCUMENT_OPEN_TYPE_ATTRIBUTE] ?? '',
    );
    expect(documentOpenTypeFeaturesFromElement(element)).toEqual({
      ligatures: 'standardContextual',
      numberForm: 'lining',
      numberSpacing: 'proportional',
      stylisticSets: [1, 20],
      contextualAlternates: true,
    });
  });

  test('keeps explicit native defaults distinct from omission', () => {
    const serialized = serializeDocumentOpenTypeFeatures({
      ligatures: 'none',
      numberForm: 'default',
      numberSpacing: 'default',
      stylisticSets: [],
      contextualAlternates: false,
    });

    expect(parseDocumentOpenTypeFeatures(serialized)).toEqual({
      ligatures: 'none',
      numberForm: 'default',
      numberSpacing: 'default',
      stylisticSets: [],
      contextualAlternates: false,
    });
    expect(documentOpenTypeDomAttributes({})).toEqual({});
  });

  test('bounds raw stylistic-set work before canonicalizing duplicates', () => {
    expect(
      normalizeDocumentOpenTypeStylisticSets(
        Array.from({ length: 4_096 }, () => 1),
      ),
    ).toEqual([1]);
    expect(
      normalizeDocumentOpenTypeStylisticSets(
        Array.from({ length: 4_097 }, () => 1),
      ),
    ).toBeNull();
  });
});
