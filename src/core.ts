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
  createWorkOfficeCollaborationSession as createOfficeCollaborationSession,
  readWorkOfficeCollaborationMetadata as readOfficeCollaborationMetadata,
  WORK_OFFICE_COLLABORATION_NAMESPACE as OFFICE_COLLABORATION_NAMESPACE,
  WORK_OFFICE_COLLABORATION_PROTOCOL as OFFICE_COLLABORATION_PROTOCOL,
  WORK_OFFICE_COLLABORATION_VERSION as OFFICE_COLLABORATION_VERSION,
  WorkOfficeCollaborationError as OfficeCollaborationError,
} from './internal/collaboration/office-collaboration';
export type {
  WorkOfficeCollaborationParticipant as OfficeCollaborationParticipant,
  WorkOfficeCollaborationPresence as OfficeCollaborationPresence,
  WorkOfficeCollaborationPresenceActivity as OfficeCollaborationPresenceActivity,
  WorkOfficeCollaborationPresenceActor as OfficeCollaborationPresenceActor,
  WorkOfficeCollaborationPresenceLocation as OfficeCollaborationPresenceLocation,
  WorkOfficeCollaborationPresenceSnapshot as OfficeCollaborationPresenceSnapshot,
  WorkOfficeCollaborationPresenceState as OfficeCollaborationPresenceState,
  WorkOfficeCollaborationPresenceUpdate as OfficeCollaborationPresenceUpdate,
  WorkOfficeDocumentPresenceLocation as OfficeDocumentPresenceLocation,
  WorkOfficeMarkdownPresenceLocation as OfficeMarkdownPresenceLocation,
  WorkOfficePdfPresenceLocation as OfficePdfPresenceLocation,
  WorkOfficePresentationPresenceLocation as OfficePresentationPresenceLocation,
  WorkOfficeSpreadsheetPresenceCell as OfficeSpreadsheetPresenceCell,
  WorkOfficeSpreadsheetPresenceLocation as OfficeSpreadsheetPresenceLocation,
  WorkOfficeSpreadsheetPresenceRange as OfficeSpreadsheetPresenceRange,
  WorkOfficeTextPresenceLocation as OfficeTextPresenceLocation,
} from './internal/collaboration/office-collaboration-presence';
export {
  createWorkOfficeCollaborationPresence as createOfficeCollaborationPresence,
  WORK_OFFICE_COLLABORATION_PRESENCE_FIELD as OFFICE_COLLABORATION_PRESENCE_FIELD,
} from './internal/collaboration/office-collaboration-presence';
export type {
  WorkOfficeCollaborationTransport as OfficeCollaborationTransport,
  WorkOfficeCollaborationTransportBinding as OfficeCollaborationTransportBinding,
  WorkOfficeCollaborationTransportBindingOptions as OfficeCollaborationTransportBindingOptions,
  WorkOfficeCollaborationTransportMessage as OfficeCollaborationTransportMessage,
  WorkOfficeCollaborationTransportMessageType as OfficeCollaborationTransportMessageType,
} from './internal/collaboration/office-collaboration-transport';
export {
  createWorkOfficeCollaborationTransportBinding as createOfficeCollaborationTransportBinding,
  WORK_OFFICE_COLLABORATION_DEFAULT_MAX_TRANSPORT_PAYLOAD_BYTES as OFFICE_COLLABORATION_DEFAULT_MAX_TRANSPORT_PAYLOAD_BYTES,
  WORK_OFFICE_COLLABORATION_MAX_TRANSPORT_PAYLOAD_BYTES as OFFICE_COLLABORATION_MAX_TRANSPORT_PAYLOAD_BYTES,
} from './internal/collaboration/office-collaboration-transport';
export type {
  WorkOfficeDocumentCollaborationBinding as OfficeDocumentCollaborationBinding,
  WorkOfficeDocumentCollaborationBindingOptions as OfficeDocumentCollaborationBindingOptions,
  WorkOfficeDocumentCollaborationChange as OfficeDocumentCollaborationChange,
} from './internal/collaboration/office-document-collaboration';
export {
  createWorkOfficeDocumentCollaborationBinding as createOfficeDocumentCollaborationBinding,
  initializeWorkOfficeDocumentCollaboration as initializeOfficeDocumentCollaboration,
  readWorkOfficeDocumentCollaboration as readOfficeDocumentCollaboration,
  workOfficeDocumentCollaborationFragment as officeDocumentCollaborationFragment,
} from './internal/collaboration/office-document-collaboration';
export type {
  WorkOfficeMarkdownCollaborationBinding as OfficeMarkdownCollaborationBinding,
  WorkOfficeMarkdownCollaborationBindingOptions as OfficeMarkdownCollaborationBindingOptions,
  WorkOfficeMarkdownCollaborationChange as OfficeMarkdownCollaborationChange,
} from './internal/collaboration/office-markdown-collaboration';
export {
  createWorkOfficeMarkdownCollaborationBinding as createOfficeMarkdownCollaborationBinding,
  initializeWorkOfficeMarkdownCollaboration as initializeOfficeMarkdownCollaboration,
  readWorkOfficeMarkdownCollaboration as readOfficeMarkdownCollaboration,
  replaceWorkOfficeMarkdownCollaboration as replaceOfficeMarkdownCollaboration,
} from './internal/collaboration/office-markdown-collaboration';
export type {
  WorkOfficePdfCollaborationBinding as OfficePdfCollaborationBinding,
  WorkOfficePdfCollaborationBindingOptions as OfficePdfCollaborationBindingOptions,
  WorkOfficePdfCollaborationChange as OfficePdfCollaborationChange,
  WorkPdfCollaborationAnnotation as PdfCollaborationAnnotation,
  WorkPdfCollaborationContent as PdfCollaborationContent,
  WorkPdfCollaborationDeletePagesOperation as PdfCollaborationDeletePagesOperation,
  WorkPdfCollaborationFormValue as PdfCollaborationFormValue,
  WorkPdfCollaborationPageOperation as PdfCollaborationPageOperation,
  WorkPdfCollaborationRect as PdfCollaborationRect,
  WorkPdfCollaborationRedactionProposal as PdfCollaborationRedactionProposal,
  WorkPdfCollaborationReorderPagesOperation as PdfCollaborationReorderPagesOperation,
  WorkPdfCollaborationReviewDecision as PdfCollaborationReviewDecision,
  WorkPdfCollaborationRotatePagesOperation as PdfCollaborationRotatePagesOperation,
  WorkPdfCollaborationSignatureAppearance as PdfCollaborationSignatureAppearance,
  WorkPdfCollaborationSignaturePlacement as PdfCollaborationSignaturePlacement,
  WorkPdfCollaborationSource as PdfCollaborationSource,
} from './internal/collaboration/office-pdf-collaboration';
export {
  assertWorkOfficePdfCollaborationSource as assertPdfCollaborationSource,
  createWorkOfficePdfCollaborationBinding as createOfficePdfCollaborationBinding,
  initializeWorkOfficePdfCollaboration as initializeOfficePdfCollaboration,
  readWorkOfficePdfCollaboration as readOfficePdfCollaboration,
  readWorkOfficePdfCollaborationSource as readOfficePdfCollaborationSource,
  replaceWorkOfficePdfCollaboration as replaceOfficePdfCollaboration,
} from './internal/collaboration/office-pdf-collaboration';
export { createWorkPdfCollaborationContent as createPdfCollaborationContent } from './internal/collaboration/office-pdf-collaboration-types';
export type {
  WorkOfficePresentationCollaborationBinding as OfficePresentationCollaborationBinding,
  WorkOfficePresentationCollaborationBindingOptions as OfficePresentationCollaborationBindingOptions,
  WorkOfficePresentationCollaborationChange as OfficePresentationCollaborationChange,
} from './internal/collaboration/office-presentation-collaboration';
export {
  createWorkOfficePresentationCollaborationBinding as createOfficePresentationCollaborationBinding,
  initializeWorkOfficePresentationCollaboration as initializeOfficePresentationCollaboration,
  readWorkOfficePresentationCollaboration as readOfficePresentationCollaboration,
  replaceWorkOfficePresentationCollaboration as replaceOfficePresentationCollaboration,
} from './internal/collaboration/office-presentation-collaboration';
export type {
  WorkOfficeSpreadsheetCollaborationBinding as OfficeSpreadsheetCollaborationBinding,
  WorkOfficeSpreadsheetCollaborationBindingOptions as OfficeSpreadsheetCollaborationBindingOptions,
  WorkOfficeSpreadsheetCollaborationChange as OfficeSpreadsheetCollaborationChange,
} from './internal/collaboration/office-spreadsheet-collaboration';
export {
  createWorkOfficeSpreadsheetCollaborationBinding as createOfficeSpreadsheetCollaborationBinding,
  initializeWorkOfficeSpreadsheetCollaboration as initializeOfficeSpreadsheetCollaboration,
  readWorkOfficeSpreadsheetCollaboration as readOfficeSpreadsheetCollaboration,
  replaceWorkOfficeSpreadsheetCollaboration as replaceOfficeSpreadsheetCollaboration,
} from './internal/collaboration/office-spreadsheet-collaboration';
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
export {
  createWorkAgentProposalRequest as createAgentProposalRequest,
  WORK_AGENT_PROPOSAL_PROTOCOL as AGENT_PROPOSAL_PROTOCOL,
  workAgentProposalInstruction as createAgentProposalInstruction,
  workAgentProposalStatus as getAgentProposalStatus,
} from './internal/features/work/work-agent-proposal';
export type { WorkEditorAgentRequest as EditorAgentRequest } from './internal/features/work/work-agent-request';
export type {
  WorkDocumentReviewConflict as DocumentReviewConflict,
  WorkDocumentReviewConflictEvent as DocumentReviewConflictEvent,
  WorkDocumentReviewConflictReason as DocumentReviewConflictReason,
  WorkDocumentReviewKind as DocumentReviewKind,
} from './internal/features/work/work-document-review-conflicts';
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
export type { WorkArtifactExportOptions as ArtifactExportOptions } from './internal/features/work/work-file-io';
export {
  createWorkArtifactBlob as createArtifactBlob,
  exportWorkArtifact as downloadArtifact,
  importWorkFile as importOfficeFile,
  WORK_IMPORT_ACCEPT as OFFICE_FILE_ACCEPT,
  workKindForFile as officeKindForFile,
} from './internal/features/work/work-file-io';
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
export {
  exportWorkArtifactPdf as downloadArtifactPdf,
  type WorkPdfExportOptions as PdfExportOptions,
  workPdfPagesForExport as pdfPagesForExport,
} from './internal/features/work/work-pdf-export';
export { defaultPptxRuntimeUrl } from './internal/features/work/work-presentation-file-io';
export {
  forgetWorkSourceBlob as forgetSourceBlob,
  readWorkSourceBlob as readSourceBlob,
  rememberWorkSourceBlob as registerSourceBlob,
} from './internal/features/work/work-repository';
export {
  createWorkArtifact as createArtifact,
  createWorkId as createOfficeId,
  WORK_TEMPLATES as officeTemplates,
} from './internal/features/work/work-templates';
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
  workArtifactExtension as artifactExtension,
  workArtifactKindLabel as artifactKindLabel,
} from './internal/features/work/work-types';
export {
  OFFICE_NOTIFICATION_EVENT,
  type OfficeNotification,
  type OfficeNotificationTone,
  subscribeOfficeNotifications,
} from './internal/state/app-state';
