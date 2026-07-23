export {
  createWorkArtifact as createArtifact,
  createWorkId as createOfficeId,
  WORK_TEMPLATES as officeTemplates,
} from './internal/features/work/work-templates';
export {
  createWorkArtifactBlob as createArtifactBlob,
  exportWorkArtifact as downloadArtifact,
  importWorkFile as importOfficeFile,
  WORK_IMPORT_ACCEPT as OFFICE_FILE_ACCEPT,
  workKindForFile as officeKindForFile,
} from './internal/features/work/work-file-io';
export type { WorkArtifactExportOptions as ArtifactExportOptions } from './internal/features/work/work-file-io';
export { defaultPptxRuntimeUrl } from './internal/features/work/work-presentation-file-io';
export {
  forgetWorkSourceBlob as forgetSourceBlob,
  readWorkSourceBlob as readSourceBlob,
  rememberWorkSourceBlob as registerSourceBlob,
} from './internal/features/work/work-repository';
export {
  exportWorkArtifactPdf as downloadArtifactPdf,
  type WorkPdfExportOptions as PdfExportOptions,
  workPdfPagesForExport as pdfPagesForExport,
} from './internal/features/work/work-pdf-export';
export {
  OFFICE_NOTIFICATION_EVENT,
  type OfficeNotification,
  type OfficeNotificationTone,
  subscribeOfficeNotifications,
} from './internal/state/app-state';
export {
  createWorkAgentProposalRequest as createAgentProposalRequest,
  WORK_AGENT_PROPOSAL_PROTOCOL as AGENT_PROPOSAL_PROTOCOL,
  workAgentProposalInstruction as createAgentProposalInstruction,
  workAgentProposalStatus as getAgentProposalStatus,
} from './internal/features/work/work-agent-proposal';
export type {
  WorkAgentProposal as AgentProposal,
  WorkAgentProposalApplyResult as AgentProposalApplyResult,
  WorkAgentProposalChange as AgentProposalChange,
  WorkAgentProposalConflict as AgentProposalConflict,
  WorkAgentProposalMessage as AgentProposalMessage,
  WorkAgentProposalRequest as AgentProposalRequest,
  WorkAgentProposalStatus as AgentProposalStatus,
  WorkAgentProposalTarget as AgentProposalTarget,
} from './internal/features/work/work-agent-proposal';
export type { WorkEditorAgentRequest as EditorAgentRequest } from './internal/features/work/work-agent-request';
export {
  workArtifactExtension as artifactExtension,
  workArtifactKindLabel as artifactKindLabel,
} from './internal/features/work/work-types';
export type {
  WorkArtifact as OfficeArtifact,
  WorkArtifactContent as OfficeArtifactContent,
  WorkArtifactKind as OfficeArtifactKind,
  WorkCompatibilityIssue as CompatibilityIssue,
  WorkCompatibilityReport as CompatibilityReport,
  WorkDocumentContent as DocumentContent,
  WorkMarkdownContent as MarkdownContent,
  WorkPdfContent as PdfContent,
  WorkPresentationContent as PresentationContent,
  WorkSourceFile as SourceFile,
  WorkSpreadsheetContent as SpreadsheetContent,
  WorkTemplate as OfficeTemplate,
} from './internal/features/work/work-types';
export * from './internal/features/work/work-types';
