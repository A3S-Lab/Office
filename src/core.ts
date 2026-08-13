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
export type {
  WorkDocumentSelectionCommandFailure as DocumentSelectionCommandFailure,
  WorkDocumentSelectionCommandResult as DocumentSelectionCommandResult,
  WorkDocumentSelectionCommands as DocumentSelectionCommands,
  WorkDocumentSelectionContext as DocumentSelectionContext,
  WorkDocumentSelectionMenuIcon as DocumentSelectionMenuIcon,
  WorkDocumentSelectionMenuItem as DocumentSelectionMenuItem,
  WorkDocumentSelectionSnapshot as DocumentSelectionSnapshot,
  WorkGetDocumentSelectionMenuItems as GetDocumentSelectionMenuItems,
} from './internal/features/work/work-document-selection-menu';
export type {
  WorkGetMarkdownSelectionMenuItems as GetMarkdownSelectionMenuItems,
  WorkMarkdownSelectionCommandFailure as MarkdownSelectionCommandFailure,
  WorkMarkdownSelectionCommandResult as MarkdownSelectionCommandResult,
  WorkMarkdownSelectionCommands as MarkdownSelectionCommands,
  WorkMarkdownSelectionContext as MarkdownSelectionContext,
  WorkMarkdownSelectionMenuIcon as MarkdownSelectionMenuIcon,
  WorkMarkdownSelectionMenuItem as MarkdownSelectionMenuItem,
  WorkMarkdownSelectionSnapshot as MarkdownSelectionSnapshot,
} from './internal/features/work/work-markdown-selection-menu';
export type {
  WorkDocumentReviewConflict as DocumentReviewConflict,
  WorkDocumentReviewConflictEvent as DocumentReviewConflictEvent,
  WorkDocumentReviewConflictReason as DocumentReviewConflictReason,
  WorkDocumentReviewKind as DocumentReviewKind,
} from './internal/features/work/work-document-review-conflicts';
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
export {
  createWorkOfficeCollaborationSession as createOfficeCollaborationSession,
  readWorkOfficeCollaborationMetadata as readOfficeCollaborationMetadata,
  WORK_OFFICE_COLLABORATION_NAMESPACE as OFFICE_COLLABORATION_NAMESPACE,
  WORK_OFFICE_COLLABORATION_PROTOCOL as OFFICE_COLLABORATION_PROTOCOL,
  WORK_OFFICE_COLLABORATION_VERSION as OFFICE_COLLABORATION_VERSION,
  WorkOfficeCollaborationError as OfficeCollaborationError,
} from './internal/collaboration/office-collaboration';
export type {
  WorkOfficeCollaborationActor as OfficeCollaborationActor,
  WorkOfficeCollaborationActorKind as OfficeCollaborationActorKind,
  WorkOfficeCollaborationAwareness as OfficeCollaborationAwareness,
  WorkOfficeCollaborationErrorCode as OfficeCollaborationErrorCode,
  WorkOfficeCollaborationMetadata as OfficeCollaborationMetadata,
  WorkOfficeCollaborationMode as OfficeCollaborationMode,
  WorkOfficeCollaborationOrigin as OfficeCollaborationOrigin,
  WorkOfficeCollaborationOriginKind as OfficeCollaborationOriginKind,
  WorkOfficeCollaborationSession as OfficeCollaborationSession,
  WorkOfficeCollaborationSessionOptions as OfficeCollaborationSessionOptions,
} from './internal/collaboration/office-collaboration';
export {
  createWorkOfficeMarkdownCollaborationBinding as createOfficeMarkdownCollaborationBinding,
  initializeWorkOfficeMarkdownCollaboration as initializeOfficeMarkdownCollaboration,
  readWorkOfficeMarkdownCollaboration as readOfficeMarkdownCollaboration,
  replaceWorkOfficeMarkdownCollaboration as replaceOfficeMarkdownCollaboration,
} from './internal/collaboration/office-markdown-collaboration';
export type {
  WorkOfficeMarkdownCollaborationBinding as OfficeMarkdownCollaborationBinding,
  WorkOfficeMarkdownCollaborationBindingOptions as OfficeMarkdownCollaborationBindingOptions,
  WorkOfficeMarkdownCollaborationChange as OfficeMarkdownCollaborationChange,
} from './internal/collaboration/office-markdown-collaboration';
