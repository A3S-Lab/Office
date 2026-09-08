import type { ParagraphChild } from 'docx';
import {
  docxDocumentFieldKind,
  supportedDocxDocumentFieldInstruction,
} from './work-document-fields';

export function docxDocumentFieldRun(
  element: HTMLElement,
  docx: typeof import('docx'),
): ParagraphChild {
  const instruction = element.dataset.fieldInstruction?.trim() ?? '';
  const display =
    element.dataset.fieldDisplay?.trim() || element.textContent?.trim() || '';
  if (
    !docxDocumentFieldKind(instruction) ||
    !supportedDocxDocumentFieldInstruction(instruction)
  )
    return new docx.TextRun(display);
  if (element.dataset.fieldLocked === 'true') {
    // Preserve WPS/Word fldLock on simple fields (Ctrl+F11).
    return new docx.BuilderElement({
      name: 'w:fldSimple',
      attributes: {
        instr: { key: 'w:instr', value: instruction },
        fldLock: { key: 'w:fldLock', value: '1' },
      },
      children: [new docx.TextRun(display)],
    }) as ParagraphChild;
  }
  return new docx.SimpleField(instruction, display);
}
