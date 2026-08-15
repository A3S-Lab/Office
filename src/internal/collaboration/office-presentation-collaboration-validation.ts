import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideComment,
  WorkSlideElement,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';

export function validateWorkOfficePresentationContent(
  content: WorkPresentationContent,
): WorkPresentationContent {
  if (!content || content.type !== 'presentation') {
    invalidWorkOfficePresentationInput('a Presentation content value');
  }
  const result: WorkPresentationContent = {
    type: 'presentation',
    slides: validateRecords(content.slides, 'slide', validateSlide),
  };
  if (content.width !== undefined) {
    result.width = requiredFiniteNumber(content.width, 'presentation width');
  }
  if (content.height !== undefined) {
    result.height = requiredFiniteNumber(content.height, 'presentation height');
  }
  if (content.masters !== undefined) {
    result.masters = validateRecords(
      content.masters,
      'presentation master',
      validateMaster,
    );
  }
  if (content.layouts !== undefined) {
    result.layouts = validateRecords(
      content.layouts,
      'presentation layout',
      validateLayout,
    );
  }
  assertPresentationReferences(result);
  return result;
}

export function validateSharedWorkOfficePresentationContent(
  content: WorkPresentationContent,
): WorkPresentationContent {
  try {
    return validateWorkOfficePresentationContent(content);
  } catch (error) {
    if (error instanceof WorkOfficeCollaborationError) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The shared Presentation collaboration content is invalid: ${error.message}`,
      );
    }
    throw error;
  }
}

export function invalidWorkOfficePresentationShared(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The shared Presentation collaboration ${label} is invalid.`,
  );
}

function validateRecords<T extends { id: string }>(
  value: unknown,
  label: string,
  validate: (value: unknown) => T,
): T[] {
  if (!Array.isArray(value)) {
    invalidWorkOfficePresentationInput(`an array of ${label}s`);
  }
  const ids = new Set<string>();
  return value.map((item) => {
    const validated = validate(item);
    if (ids.has(validated.id)) {
      invalidWorkOfficePresentationInput(
        `a unique ${label} ID; '${validated.id}' is repeated`,
      );
    }
    ids.add(validated.id);
    return validated;
  });
}

function validateSlide(value: unknown): WorkSlide {
  const record = requiredInputRecord(value, 'slide');
  const slide = validateJsonRecord(record, 'slide') as unknown as WorkSlide;
  slide.id = requiredIdentifier(record.id, 'slide');
  slide.name = requiredString(record.name, 'slide name');
  slide.background = requiredString(record.background, 'slide background');
  slide.elements = validateRecords(
    record.elements,
    `element in slide '${slide.id}'`,
    validateElement,
  );
  if (record.layoutId !== undefined) {
    slide.layoutId = requiredIdentifier(record.layoutId, 'slide layout');
  }
  if (record.comments !== undefined) {
    slide.comments = validateRecords(
      record.comments,
      `comment in slide '${slide.id}'`,
      validateComment,
    );
  }
  return slide;
}

function validateMaster(value: unknown): WorkPresentationMaster {
  const record = requiredInputRecord(value, 'presentation master');
  const master = validateJsonRecord(
    record,
    'presentation master',
  ) as unknown as WorkPresentationMaster;
  master.id = requiredIdentifier(record.id, 'presentation master');
  master.name = requiredString(record.name, 'presentation master name');
  master.background = requiredString(
    record.background,
    'presentation master background',
  );
  master.elements = validateRecords(
    record.elements,
    `element in presentation master '${master.id}'`,
    validateElement,
  );
  return master;
}

function validateLayout(value: unknown): WorkPresentationLayout {
  const record = requiredInputRecord(value, 'presentation layout');
  const layout = validateJsonRecord(
    record,
    'presentation layout',
  ) as unknown as WorkPresentationLayout;
  layout.id = requiredIdentifier(record.id, 'presentation layout');
  layout.name = requiredString(record.name, 'presentation layout name');
  layout.masterId = requiredIdentifier(
    record.masterId,
    'presentation layout master',
  );
  layout.elements = validateRecords(
    record.elements,
    `element in presentation layout '${layout.id}'`,
    validateElement,
  );
  return layout;
}

function validateElement(value: unknown): WorkSlideElement {
  const record = requiredInputRecord(value, 'presentation element');
  if (Object.hasOwn(record, 'tombstone')) {
    invalidWorkOfficePresentationInput(
      "a presentation element without the reserved 'tombstone' field",
    );
  }
  const element = validateJsonRecord(
    record,
    'presentation element',
  ) as unknown as WorkSlideElement;
  element.id = requiredIdentifier(record.id, 'presentation element');
  element.type = requiredString(
    record.type,
    'presentation element type',
  ) as WorkSlideElement['type'];
  if (
    element.type !== 'text' &&
    element.type !== 'shape' &&
    element.type !== 'image' &&
    element.type !== 'table' &&
    element.type !== 'chart' &&
    element.type !== 'line'
  ) {
    invalidWorkOfficePresentationInput('a supported presentation element type');
  }
  for (const key of ['x', 'y', 'width', 'height', 'fontSize'] as const) {
    element[key] = requiredFiniteNumber(
      record[key],
      `presentation element ${key}`,
    );
  }
  element.text = requiredString(record.text, 'presentation element text');
  element.color = requiredString(record.color, 'presentation element color');
  element.fill = requiredString(record.fill, 'presentation element fill');
  if (typeof record.bold !== 'boolean') {
    invalidWorkOfficePresentationInput(
      'a boolean presentation element bold value',
    );
  }
  element.bold = record.bold;
  if (
    record.align !== 'left' &&
    record.align !== 'center' &&
    record.align !== 'right'
  ) {
    invalidWorkOfficePresentationInput(
      'a supported presentation text alignment',
    );
  }
  element.align = record.align;
  return element;
}

function validateComment(value: unknown): WorkSlideComment {
  const record = requiredInputRecord(value, 'slide comment');
  const comment = validateJsonRecord(
    record,
    'slide comment',
  ) as unknown as WorkSlideComment;
  comment.id = requiredIdentifier(record.id, 'slide comment');
  comment.author = requiredString(record.author, 'slide comment author');
  comment.date = requiredString(record.date, 'slide comment date');
  comment.text = requiredString(record.text, 'slide comment text');
  comment.x = requiredFiniteNumber(record.x, 'slide comment x');
  comment.y = requiredFiniteNumber(record.y, 'slide comment y');
  if (record.initials !== undefined) {
    comment.initials = requiredString(
      record.initials,
      'slide comment initials',
    );
  }
  return comment;
}

function validateJsonRecord(
  value: Record<string, unknown>,
  label: string,
): Record<string, unknown> {
  try {
    return cloneJsonValue(value) as Record<string, unknown>;
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `Presentation collaboration requires a JSON-compatible ${label}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

function assertPresentationReferences(content: WorkPresentationContent): void {
  const masterIds = new Set((content.masters ?? []).map(({ id }) => id));
  for (const layout of content.layouts ?? []) {
    if (!masterIds.has(layout.masterId)) {
      invalidWorkOfficePresentationInput(
        `presentation layout '${layout.id}' to reference an existing master`,
      );
    }
  }
  const layoutIds = new Set((content.layouts ?? []).map(({ id }) => id));
  for (const slide of content.slides) {
    if (slide.layoutId !== undefined && !layoutIds.has(slide.layoutId)) {
      invalidWorkOfficePresentationInput(
        `slide '${slide.id}' to reference an existing presentation layout`,
      );
    }
  }
}

function requiredInputRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  if (!isRecord(value)) {
    invalidWorkOfficePresentationInput(`a valid ${label} record`);
  }
  return value as Record<string, unknown>;
}

function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredString(value, `${label} ID`);
  if (!result || result !== result.trim() || result.length > 256) {
    invalidWorkOfficePresentationInput(
      `a ${label} ID containing 1 to 256 characters`,
    );
  }
  return result;
}

function requiredString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    invalidWorkOfficePresentationInput(`a string ${label}`);
  }
  return value as string;
}

function requiredFiniteNumber(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    invalidWorkOfficePresentationInput(`a finite number for ${label}`);
  }
  return value as number;
}

function invalidWorkOfficePresentationInput(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `Presentation collaboration requires ${expected}.`,
  );
}
