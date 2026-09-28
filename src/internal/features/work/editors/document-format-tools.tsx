import type { Editor } from '@tiptap/core';
import { ClipboardCopy, ClipboardPaste, Paintbrush } from 'lucide-react';
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import {
  copyDocumentFormatting,
  hasDocumentFormatClipboard,
  pasteDocumentFormatting,
  subscribeDocumentFormatClipboard,
} from './document-format-clipboard';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

interface DocumentFormatPainterSelection {
  from: number;
  to: number;
}

export function DocumentFormatTools({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const hasFormat = useSyncExternalStore(
    subscribeDocumentFormatClipboard,
    hasDocumentFormatClipboard,
    hasDocumentFormatClipboard,
  );
  const [formatPainterActive, setFormatPainterActive] = useState(false);
  const painterSourceRef = useRef<DocumentFormatPainterSelection | null>(null);
  const copyFormatCommand = getDocumentCommandDefinition('copyFormat');
  const pasteFormatCommand = getDocumentCommandDefinition('pasteFormat');

  useEffect(() => {
    if (!formatPainterActive || editor.isDestroyed) return;
    const editorDom = editor.view.dom;
    editorDom.dataset.documentFormatPainter = 'true';
    const applyFormatPainter = () => {
      if (editor.isDestroyed) return;
      const selection = editor.state.selection;
      const source = painterSourceRef.current;
      if (
        source &&
        source.from === selection.from &&
        source.to === selection.to
      ) {
        return;
      }
      if (pasteDocumentFormatting(editor)) setFormatPainterActive(false);
    };
    const onPointerUp = (event: PointerEvent) => {
      if (event.button === 0) applyFormatPainter();
    };
    const onKeyUp = (event: KeyboardEvent) => {
      if (
        event.shiftKey ||
        event.key.startsWith('Arrow') ||
        event.key === 'Home' ||
        event.key === 'End'
      ) {
        applyFormatPainter();
      }
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      setFormatPainterActive(false);
    };
    editorDom.addEventListener('pointerup', onPointerUp);
    editorDom.addEventListener('keyup', onKeyUp);
    editorDom.addEventListener('keydown', onKeyDown, true);
    return () => {
      delete editorDom.dataset.documentFormatPainter;
      editorDom.removeEventListener('pointerup', onPointerUp);
      editorDom.removeEventListener('keyup', onKeyUp);
      editorDom.removeEventListener('keydown', onKeyDown, true);
    };
  }, [editor, formatPainterActive]);

  const copyFormat = () => copyDocumentFormatting(editor);
  const toggleFormatPainter = () => {
    if (formatPainterActive) {
      setFormatPainterActive(false);
      return;
    }
    if (!copyDocumentFormatting(editor)) return;
    painterSourceRef.current = {
      from: editor.state.selection.from,
      to: editor.state.selection.to,
    };
    setFormatPainterActive(true);
  };

  return (
    <WorkOfficeRibbonGroup
      label={officeMessage(messages, 'document.format.clipboardGroup')}
      priority="normal"
    >
      <WorkOfficeRibbonButton
        label={documentCommandLabel('copyFormat', messages)}
        title={officeMessage(messages, 'document.format.copyTitle', {
          shortcut: copyFormatCommand.shortcut?.label ?? '',
        })}
        aria-keyshortcuts={copyFormatCommand.shortcut?.aria}
        onClick={copyFormat}
      >
        <ClipboardCopy size={17} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={documentCommandLabel('pasteFormat', messages)}
        title={officeMessage(messages, 'document.format.pasteTitle', {
          shortcut: pasteFormatCommand.shortcut?.label ?? '',
        })}
        aria-keyshortcuts={pasteFormatCommand.shortcut?.aria}
        disabled={!hasFormat}
        onClick={() => pasteDocumentFormatting(editor)}
      >
        <ClipboardPaste size={17} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={documentCommandLabel('formatPainter', messages)}
        title={officeMessage(messages, 'document.format.painterTitle')}
        active={formatPainterActive}
        onClick={toggleFormatPainter}
      >
        <Paintbrush size={17} />
      </WorkOfficeRibbonButton>
    </WorkOfficeRibbonGroup>
  );
}
