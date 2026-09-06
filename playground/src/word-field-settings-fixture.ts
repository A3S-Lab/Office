import type { OfficeArtifact } from '@a3s-lab/office/core';
import { createWorkArtifact } from '../../src/internal/features/work/work-templates';

export const WORD_FIELD_SETTINGS_FIXTURE = 'word-field-settings';
export const WORD_FIELD_SETTINGS_ARTIFACT_ID = 'e2e-word-field-settings';

export function createWordFieldSettingsArtifact(): OfficeArtifact {
  const artifact = createWorkArtifact('blank-document');
  artifact.id = WORD_FIELD_SETTINGS_ARTIFACT_ID;
  artifact.title = '字段设置';
  return artifact;
}
