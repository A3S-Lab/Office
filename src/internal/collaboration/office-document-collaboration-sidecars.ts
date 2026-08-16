import type * as Y from 'yjs';
import type {
  WorkDocumentBibliography,
  WorkDocumentChangeDecision,
  WorkDocumentComment,
  WorkDocumentContent,
} from '../features/work/work-types';
import {
  assertWorkOfficeCollaborationEditable,
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationOrigin,
  type WorkOfficeCollaborationSession,
} from './office-collaboration';
import { assertWorkOfficeDocumentCommentMutationAllowed } from './office-document-collaboration-comment-permissions';
import {
  assertWorkOfficeDocumentChangeDecisionConflicts,
  initializeWorkOfficeDocumentChangeDecisions,
  patchWorkOfficeDocumentChangeDecisions,
  readWorkOfficeDocumentChangeDecisions,
  validatedWorkOfficeDocumentChangeDecisions,
} from './office-document-collaboration-change-decisions';
import {
  assertWorkOfficeDocumentBibliographyConflicts,
  initializeWorkOfficeDocumentBibliography,
  patchWorkOfficeDocumentBibliography,
  readWorkOfficeDocumentBibliography,
  validatedWorkOfficeDocumentBibliography,
} from './office-document-collaboration-bibliography';
import {
  assertWorkOfficeDocumentCommentConflicts,
  initializeWorkOfficeDocumentComments,
  patchWorkOfficeDocumentComments,
  readWorkOfficeDocumentComments,
  validatedWorkOfficeDocumentComments,
} from './office-document-collaboration-comments';
import {
  appendWorkOfficeDocumentRecordClaims,
  assertWorkOfficeDocumentRecordClaims,
} from './office-document-collaboration-claims';
import {
  invalidInputSidecars,
  invalidSharedSidecars,
  jsonEqual,
  patchOptionalScalar,
} from './office-document-collaboration-sidecar-utils';

const DOCUMENT_OPTIONS_ROOT = 'document.options';
const DOCUMENT_COMMENTS_ROOT = 'document.comments';
const DOCUMENT_COMMENT_ORDER_ROOT = 'document.comment-order';
const DOCUMENT_CHANGE_DECISIONS_ROOT = 'document.change-decisions';
const DOCUMENT_CHANGE_DECISION_ORDER_ROOT = 'document.change-decision-order';
const DOCUMENT_BIBLIOGRAPHY_ROOT = 'document.bibliography';
const DOCUMENT_BIBLIOGRAPHY_SOURCES_ROOT = 'document.bibliography.sources';
const DOCUMENT_BIBLIOGRAPHY_SOURCE_ORDER_ROOT =
  'document.bibliography.source-order';
const DOCUMENT_RECORD_CLAIMS_ROOT = 'document.record-claims';

export interface WorkOfficeDocumentSidecars {
  pageColor?: string;
  trackChanges?: boolean;
  changeDecisions?: WorkDocumentChangeDecision[];
  comments?: WorkDocumentComment[];
  bibliography?: WorkDocumentBibliography;
}

interface DocumentSidecarRoots {
  options: Y.Map<unknown>;
  comments: Y.Map<unknown>;
  commentOrder: Y.Array<string>;
  changeDecisions: Y.Map<unknown>;
  changeDecisionOrder: Y.Array<string>;
  bibliography: Y.Map<unknown>;
  bibliographySources: Y.Map<unknown>;
  bibliographySourceOrder: Y.Array<string>;
  recordClaims: Y.Array<string>;
}

export function assertWorkOfficeDocumentSidecarsEmpty(
  session: WorkOfficeCollaborationSession,
): void {
  const roots = documentSidecarRoots(session);
  if (
    roots.options.size > 0 ||
    roots.comments.size > 0 ||
    roots.commentOrder.length > 0 ||
    roots.changeDecisions.size > 0 ||
    roots.changeDecisionOrder.length > 0 ||
    roots.bibliography.size > 0 ||
    roots.bibliographySources.size > 0 ||
    roots.bibliographySourceOrder.length > 0 ||
    roots.recordClaims.length > 0
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The Document collaboration sidecars contain data without initialized metadata.',
    );
  }
}

export function initializeWorkOfficeDocumentSidecars(
  session: WorkOfficeCollaborationSession,
  content: WorkDocumentContent,
): void {
  const roots = documentSidecarRoots(session);
  const sidecars = validatedWorkOfficeDocumentSidecars(content);
  appendWorkOfficeDocumentRecordClaims(roots.recordClaims, {}, sidecars);
  if (sidecars.pageColor !== undefined) {
    roots.options.set('pageColor', sidecars.pageColor);
  }
  if (sidecars.trackChanges !== undefined) {
    roots.options.set('trackChanges', sidecars.trackChanges);
  }
  if (sidecars.changeDecisions !== undefined) {
    initializeWorkOfficeDocumentChangeDecisions(
      roots.changeDecisions,
      roots.changeDecisionOrder,
      sidecars.changeDecisions,
    );
  }
  if (sidecars.comments !== undefined) {
    roots.options.set('commentsPresent', true);
    initializeWorkOfficeDocumentComments(
      roots.comments,
      roots.commentOrder,
      sidecars.comments,
    );
  }
  if (sidecars.bibliography) {
    roots.options.set('bibliographyPresent', true);
    initializeWorkOfficeDocumentBibliography(
      roots.bibliography,
      roots.bibliographySources,
      roots.bibliographySourceOrder,
      sidecars.bibliography,
    );
  }
}

export function readWorkOfficeDocumentSidecars(
  session: WorkOfficeCollaborationSession,
): WorkOfficeDocumentSidecars {
  const roots = documentSidecarRoots(session);
  const result: WorkOfficeDocumentSidecars = {};
  const pageColor = roots.options.get('pageColor');
  if (pageColor !== undefined) {
    if (typeof pageColor !== 'string') invalidSharedSidecars('page color');
    result.pageColor = pageColor as string;
  }
  const trackChanges = roots.options.get('trackChanges');
  if (trackChanges !== undefined) {
    if (typeof trackChanges !== 'boolean') {
      invalidSharedSidecars('track-changes setting');
    }
    result.trackChanges = trackChanges as boolean;
  }
  if (roots.changeDecisionOrder.length > 0) {
    result.changeDecisions = readWorkOfficeDocumentChangeDecisions(
      roots.changeDecisions,
      roots.changeDecisionOrder,
    );
  } else if (roots.changeDecisions.size > 0) {
    invalidSharedSidecars('tracked-change decision order and record set');
  }
  const commentsPresent = optionalPresence(
    roots.options.get('commentsPresent'),
    'comment',
  );
  if (commentsPresent === true || roots.commentOrder.length > 0) {
    result.comments = readWorkOfficeDocumentComments(
      roots.comments,
      roots.commentOrder,
    );
  } else if (roots.comments.size > 0) {
    invalidSharedSidecars('comment order and record set');
  }

  const bibliographyPresent = optionalPresence(
    roots.options.get('bibliographyPresent'),
    'bibliography',
  );
  if (
    bibliographyPresent === true ||
    roots.bibliographySourceOrder.length > 0 ||
    (bibliographyPresent === undefined && roots.bibliography.size > 0)
  ) {
    result.bibliography = readWorkOfficeDocumentBibliography(
      roots.bibliography,
      roots.bibliographySources,
      roots.bibliographySourceOrder,
    );
  } else if (roots.bibliographySources.size > 0) {
    invalidSharedSidecars('bibliography source order and record set');
  }
  assertWorkOfficeDocumentRecordClaims(roots.recordClaims, result);
  return result;
}

export function updateWorkOfficeDocumentSidecars(
  session: WorkOfficeCollaborationSession,
  previous: WorkDocumentContent,
  next: WorkDocumentContent,
  origin: WorkOfficeCollaborationOrigin,
): boolean {
  const before = validatedWorkOfficeDocumentSidecars(previous);
  const after = validatedWorkOfficeDocumentSidecars(next);
  if (jsonEqual(before, after)) return false;
  assertDocumentSidecarMutationAllowed(session, before, after);

  const roots = documentSidecarRoots(session);
  const shared = readWorkOfficeDocumentSidecars(session);
  assertWorkOfficeDocumentCommentConflicts(
    before.comments ?? [],
    after.comments ?? [],
    shared.comments ?? [],
  );
  assertWorkOfficeDocumentChangeDecisionConflicts(
    before.changeDecisions ?? [],
    after.changeDecisions ?? [],
    shared.changeDecisions ?? [],
  );
  assertWorkOfficeDocumentBibliographyConflicts(
    before.bibliography,
    after.bibliography,
    shared.bibliography,
  );
  session.transact(
    () => {
      appendWorkOfficeDocumentRecordClaims(roots.recordClaims, before, after);
      patchOptionalScalar(
        roots.options,
        'pageColor',
        before.pageColor,
        after.pageColor,
      );
      patchOptionalScalar(
        roots.options,
        'trackChanges',
        before.trackChanges,
        after.trackChanges,
      );
      if (!jsonEqual(before.changeDecisions, after.changeDecisions)) {
        patchWorkOfficeDocumentChangeDecisions(
          roots.changeDecisions,
          roots.changeDecisionOrder,
          before.changeDecisions ?? [],
          after.changeDecisions ?? [],
        );
      }
      if (!jsonEqual(before.comments, after.comments)) {
        patchPresence(
          roots.options,
          'commentsPresent',
          before.comments,
          after.comments,
        );
        patchWorkOfficeDocumentComments(
          roots.comments,
          roots.commentOrder,
          before.comments ?? [],
          after.comments ?? [],
        );
      }
      if (!jsonEqual(before.bibliography, after.bibliography)) {
        patchPresence(
          roots.options,
          'bibliographyPresent',
          before.bibliography,
          after.bibliography,
          false,
        );
        patchWorkOfficeDocumentBibliography(
          roots.bibliography,
          roots.bibliographySources,
          roots.bibliographySourceOrder,
          before.bibliography,
          after.bibliography,
        );
      }
    },
    origin,
    session.mode === 'comment' ? 'document-comment' : 'content',
  );
  return true;
}

export function workOfficeDocumentSidecarUndoScope(
  session: WorkOfficeCollaborationSession,
): Array<Y.Map<unknown> | Y.Array<string>> {
  const roots = documentSidecarRoots(session);
  return [
    roots.options,
    roots.comments,
    roots.commentOrder,
    roots.bibliography,
    roots.bibliographySources,
    roots.bibliographySourceOrder,
    roots.recordClaims,
  ];
}

export function workOfficeDocumentDecisionRootsChanged(
  session: WorkOfficeCollaborationSession,
  transaction: Y.Transaction,
): boolean {
  const roots = documentSidecarRoots(session);
  const changed = new Set<unknown>(transaction.changedParentTypes.keys());
  return (
    changed.has(roots.changeDecisions) || changed.has(roots.changeDecisionOrder)
  );
}

export function workOfficeDocumentSidecarsChanged(
  session: WorkOfficeCollaborationSession,
  transaction: Y.Transaction,
): boolean {
  const changed = new Set<unknown>(transaction.changedParentTypes.keys());
  const roots = documentSidecarRoots(session);
  return [
    ...workOfficeDocumentSidecarUndoScope(session),
    roots.changeDecisions,
    roots.changeDecisionOrder,
  ].some((root) => changed.has(root));
}

export function validatedWorkOfficeDocumentSidecars(
  content: WorkDocumentContent,
): WorkOfficeDocumentSidecars {
  if (!content || content.type !== 'document') {
    invalidInputSidecars('a Document content value');
  }
  const result: WorkOfficeDocumentSidecars = {};
  if (content.pageColor !== undefined) {
    if (typeof content.pageColor !== 'string') {
      invalidInputSidecars('a string page color');
    }
    result.pageColor = content.pageColor;
  }
  if (content.trackChanges !== undefined) {
    if (typeof content.trackChanges !== 'boolean') {
      invalidInputSidecars('a boolean track-changes setting');
    }
    result.trackChanges = content.trackChanges;
  }
  if (content.changeDecisions !== undefined) {
    result.changeDecisions = validatedWorkOfficeDocumentChangeDecisions(
      content.changeDecisions,
    );
  }
  if (content.comments !== undefined) {
    result.comments = validatedWorkOfficeDocumentComments(content.comments);
  }
  if (content.bibliography !== undefined) {
    result.bibliography = validatedWorkOfficeDocumentBibliography(
      content.bibliography,
    );
  }
  return result;
}

function assertDocumentSidecarMutationAllowed(
  session: WorkOfficeCollaborationSession,
  previous: WorkOfficeDocumentSidecars,
  next: WorkOfficeDocumentSidecars,
): void {
  if (session.mode === 'edit') return;
  if (session.mode !== 'comment') {
    assertWorkOfficeCollaborationEditable(session);
    return;
  }
  if (
    !jsonEqual(previous.pageColor, next.pageColor) ||
    !jsonEqual(previous.trackChanges, next.trackChanges) ||
    !jsonEqual(previous.changeDecisions, next.changeDecisions) ||
    !jsonEqual(previous.bibliography, next.bibliography)
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.permission_denied',
      'The comment collaboration mode can modify only Document review records.',
    );
  }
  assertWorkOfficeDocumentCommentMutationAllowed(
    session,
    previous.comments ?? [],
    next.comments ?? [],
  );
}

function documentSidecarRoots(
  session: WorkOfficeCollaborationSession,
): DocumentSidecarRoots {
  return {
    options: session.document.getMap(session.rootName(DOCUMENT_OPTIONS_ROOT)),
    comments: session.document.getMap(session.rootName(DOCUMENT_COMMENTS_ROOT)),
    commentOrder: session.document.getArray(
      session.rootName(DOCUMENT_COMMENT_ORDER_ROOT),
    ),
    changeDecisions: session.document.getMap(
      session.rootName(DOCUMENT_CHANGE_DECISIONS_ROOT),
    ),
    changeDecisionOrder: session.document.getArray(
      session.rootName(DOCUMENT_CHANGE_DECISION_ORDER_ROOT),
    ),
    bibliography: session.document.getMap(
      session.rootName(DOCUMENT_BIBLIOGRAPHY_ROOT),
    ),
    bibliographySources: session.document.getMap(
      session.rootName(DOCUMENT_BIBLIOGRAPHY_SOURCES_ROOT),
    ),
    bibliographySourceOrder: session.document.getArray(
      session.rootName(DOCUMENT_BIBLIOGRAPHY_SOURCE_ORDER_ROOT),
    ),
    recordClaims: session.document.getArray(
      session.rootName(DOCUMENT_RECORD_CLAIMS_ROOT),
    ),
  };
}

function optionalPresence(value: unknown, label: string): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'boolean') {
    invalidSharedSidecars(`${label} presence setting`);
  }
  return value as boolean;
}

function patchPresence(
  target: Y.Map<unknown>,
  key: string,
  previous: unknown,
  next: unknown,
  absentValue: false | undefined = undefined,
): void {
  if (next !== undefined) target.set(key, true);
  else if (previous !== undefined) {
    if (absentValue === false) target.set(key, false);
    else target.delete(key);
  }
}
