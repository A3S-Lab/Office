import type { OfficeTemplate } from '@a3s-lab/office/core';
import { WORK_TEMPLATES as officeTemplates } from '../../src/internal/features/work/work-templates';

export type LatestCapabilityLaunch =
  | { type: 'template'; templateId: string }
  | { type: 'pdf-page-organization' };

export type LatestCapability = Pick<
  OfficeTemplate,
  'description' | 'id' | 'kind' | 'name'
> & {
  launch: LatestCapabilityLaunch;
  release: string;
};

type LatestTemplateRegistration = {
  release: string;
  templateId: string;
};

const LATEST_CAPABILITY_REGISTRATIONS = [
  { templateId: 'animated-deck', release: '0.34.0' },
  {
    id: 'pdf-page-organization',
    kind: 'pdf',
    name: '组织 PDF 页面',
    description: '插入、删除、旋转、重排、抽取、合并与拆分',
    launch: { type: 'pdf-page-organization' },
    release: '0.33.0',
  },
  { templateId: 'document-comparison', release: '0.32.0' },
  { templateId: 'table-of-contents', release: '0.30.0' },
  { templateId: 'document-index', release: '0.31.0' },
  { templateId: 'run-shading', release: '0.30.0' },
  { templateId: 'proofing-languages', release: '0.30.0' },
  { templateId: 'data-validation', release: '0.30.0' },
] satisfies readonly (LatestTemplateRegistration | LatestCapability)[];

function resolveLatestCapability(
  registration: LatestTemplateRegistration | LatestCapability,
): LatestCapability[] {
  if ('launch' in registration) return [registration];

  const template = officeTemplates.find(
    ({ id }) => id === registration.templateId,
  );
  if (!template) return [];

  return [
    {
      id: template.id,
      kind: template.kind,
      name: template.name,
      description: template.description,
      launch: { type: 'template', templateId: template.id },
      release: registration.release,
    },
  ];
}

export const LATEST_CAPABILITIES: readonly LatestCapability[] =
  LATEST_CAPABILITY_REGISTRATIONS.flatMap(resolveLatestCapability);
