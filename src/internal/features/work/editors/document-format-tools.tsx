import type { Editor } from '@tiptap/core';
import { ClipboardCopy, ClipboardPaste, Paintbrush } from 'lucide-react';
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import {
  copyDocumentFormatting,
  hasDocumentFormatClipboard,
  pasteDocumentFormatting,
  subscribeDocumentFormatClipboard,
} from './document-format-clipboard';
import { getDocumentCommandDefinition } from './document-command-catalog';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

interface DocumentFormatPainterSelection {
  from: number;
  to: number;
}

export function DocumentFormatTools({ editor }: { editor: Editor }) {
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
    <WorkOfficeRibbonGroup label="剪贴板" priority="normal">
      <WorkOfficeRibbonButton
        label="复制格式"
        title={`复制格式（${copyFormatCommand.shortcut?.label}）`}
        aria-keyshortcuts={copyFormatCommand.shortcut?.aria}
        onClick={copyFormat}
      >
        <ClipboardCopy size={17} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="粘贴格式"
        title={`粘贴格式（${pasteFormatCommand.shortcut?.label}）`}
        aria-keyshortcuts={pasteFormatCommand.shortcut?.aria}
        disabled={!hasFormat}
        onClick={() => pasteDocumentFormatting(editor)}
      >
        <ClipboardPaste size={17} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="格式刷"
        title="格式刷（选择源格式后应用到下一处选择）"
        active={formatPainterActive}
        onClick={toggleFormatPainter}
      >
        <Paintbrush size={17} />
      </WorkOfficeRibbonButton>
    </WorkOfficeRibbonGroup>
  );
}
