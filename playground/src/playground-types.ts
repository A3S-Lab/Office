import type { OfficeArtifactKind } from '@a3s-lab/office/core';

export type SiteRoute = 'office' | 'guide' | 'cli' | 'skill';

export type EditableOfficeKind = Exclude<OfficeArtifactKind, 'pdf'>;

export type NoticeTone = 'neutral' | 'success' | 'danger';

export interface PlaygroundNotice {
  id: number;
  message: string;
  tone: NoticeTone;
}
