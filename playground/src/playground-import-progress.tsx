import type {
  OfficeFileImportProgress,
  OfficeFileImportStage,
} from '@a3s-lab/office/core';
import { FileUp, X } from 'lucide-react';

export interface PlaygroundImportState {
  fileName: string;
  id: number;
  progress: OfficeFileImportProgress;
}

const stageLabels: Record<OfficeFileImportStage, string> = {
  reading: '正在读取文件',
  parsing: '正在解析内容',
  analyzing: '正在检查兼容性',
  finalizing: '正在完成导入',
};

export function PlaygroundImportProgress({
  state,
  onCancel,
}: {
  state: PlaygroundImportState;
  onCancel: () => void;
}) {
  const percentage = Math.round(state.progress.progress * 100);
  return (
    <section
      className="playground-import-progress"
      aria-label={`正在导入 ${state.fileName}`}
    >
      <FileUp size={18} aria-hidden="true" />
      <div className="playground-import-progress-copy">
        <strong title={state.fileName}>{state.fileName}</strong>
        <output aria-live="polite">
          {stageLabels[state.progress.stage]} · {percentage}%
        </output>
      </div>
      <button
        type="button"
        aria-label={`取消导入 ${state.fileName}`}
        title="取消导入"
        onClick={onCancel}
      >
        <X size={16} />
      </button>
      <progress
        aria-label={`导入进度 ${percentage}%`}
        max={1}
        value={state.progress.progress}
      />
    </section>
  );
}
