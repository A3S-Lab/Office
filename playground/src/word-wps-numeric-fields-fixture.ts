import type { OfficeArtifact } from '@a3s-lab/office/core';
import { createWorkArtifact } from '../../src/internal/features/work/work-templates';

export const WORD_WPS_NUMERIC_FIELDS_FIXTURE = 'word-wps-numeric-fields';
export const WORD_WPS_NUMERIC_FIELDS_ARTIFACT_ID =
  'e2e-word-wps-numeric-fields';

export function createWordWpsNumericFieldsArtifact(): OfficeArtifact {
  const artifact = createWorkArtifact('blank-document');
  artifact.id = WORD_WPS_NUMERIC_FIELDS_ARTIFACT_ID;
  artifact.title = 'WPS 数字字段';
  if (artifact.content.type !== 'document') return artifact;
  artifact.content.html = [
    '<section data-document-section="true"><p>Page: ',
    '<span data-document-field="true" data-field-id="roman-page" data-field-kind="page" data-field-instruction="PAGE \\* ROMAN \\* MERGEFORMAT" data-field-display="XLII">XLII</span>',
    ' / total: ',
    '<span data-document-field="true" data-field-id="alpha-pages" data-field-kind="numPages" data-field-instruction="NUMPAGES \\* ALPHABETIC \\* MERGEFORMAT" data-field-display="AP">AP</span>',
    ' / section: ',
    '<span data-document-field="true" data-field-id="ordinal-section" data-field-kind="section" data-field-instruction="SECTION \\* Ordinal \\* MERGEFORMAT" data-field-display="1st">1st</span>',
    '</p><p>',
    '<span data-document-bookmark-boundary="true" data-bookmark-kind="start" data-bookmark-id="bookmark-target" data-bookmark-name="Fields_target"></span>',
    'Target section',
    '<span data-document-bookmark-boundary="true" data-bookmark-kind="end" data-bookmark-id="bookmark-target" data-bookmark-name="Fields_target"></span>',
    ' on page ',
    '<span data-document-field="true" data-field-id="ordinal-page" data-field-kind="pageReference" data-field-instruction="PAGEREF Fields_target \\h \\* Ordinal \\* MERGEFORMAT" data-field-display="7th" data-field-target-id="bookmark-target" data-field-target-name="Fields_target">7th</span>',
    '</p></section>',
  ].join('');
  return artifact;
}
