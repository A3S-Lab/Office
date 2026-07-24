import { CheckCheck, Cloud, FileText, Globe2 } from 'lucide-react';
import {
  clampDocumentZoom,
  MAX_DOCUMENT_ZOOM,
  MIN_DOCUMENT_ZOOM,
} from './document-editor-support';
import type { DocumentViewMode } from './document-toolbar';
import {
  WorkOfficeStatusBar,
  WorkOfficeZoomControls,
} from './work-office-chrome';

interface DocumentStatusBarProps {
  bibliographyCount: number;
  citationCount: number;
  currentPage: number;
  pageCount: number;
  saveStatus: string;
  sectionCount: number;
  sectionIndex: number;
  spellcheckEnabled: boolean;
  viewMode: DocumentViewMode;
  wordCount: number;
  zoom: number;
  onSpellcheckChange: (enabled: boolean) => void;
  onViewModeChange: (mode: DocumentViewMode) => void;
  onZoomChange: (zoom: number) => void;
}

export function DocumentStatusBar({
  bibliographyCount,
  citationCount,
  currentPage,
  pageCount,
  saveStatus,
  sectionCount,
  sectionIndex,
  spellcheckEnabled,
  viewMode,
  wordCount,
  zoom,
  onSpellcheckChange,
  onViewModeChange,
  onZoomChange,
}: DocumentStatusBarProps) {
  return (
    <WorkOfficeStatusBar
      className="work-document-footer"
      controls={
        <>
          <button
            type="button"
            aria-label="页面视图"
            title="页面视图"
            aria-pressed={viewMode === 'page'}
            onClick={() => onViewModeChange('page')}
          >
            <FileText size={13} />
          </button>
          <button
            type="button"
            aria-label="网页视图"
            title="网页视图"
            aria-pressed={viewMode === 'web'}
            onClick={() => onViewModeChange('web')}
          >
            <Globe2 size={13} />
          </button>
          <span className="work-office-status-divider" />
          <WorkOfficeZoomControls
            zoom={zoom}
            minimum={MIN_DOCUMENT_ZOOM}
            maximum={MAX_DOCUMENT_ZOOM}
            step={5}
            decreaseLabel="缩小文档"
            increaseLabel="放大文档"
            outputLabel="文档缩放比例"
            sliderLabel="文档缩放"
            onChange={(value) => onZoomChange(clampDocumentZoom(value))}
          />
        </>
      }
    >
      <output aria-label="页码状态">
        第 {currentPage} 页，共 {pageCount} 页
      </output>
      <output aria-label="分节状态">
        第 {sectionIndex + 1} 节，共 {sectionCount} 节
      </output>
      <output aria-label="字数统计">字数：{wordCount}</output>
      <button
        type="button"
        aria-label={`校对：${spellcheckEnabled ? '已开启' : '已关闭'}`}
        title={`校对：${spellcheckEnabled ? '已开启' : '已关闭'}`}
        aria-pressed={spellcheckEnabled}
        onClick={() => onSpellcheckChange(!spellcheckEnabled)}
      >
        <CheckCheck size={12} />
      </button>
      <output aria-label="引用状态">
        {bibliographyCount} 条文献 · {citationCount} 处引文
      </output>
      <output aria-label="文档保存状态" className="work-office-save-status">
        <Cloud size={12} />
        {saveStatus}
      </output>
    </WorkOfficeStatusBar>
  );
}
