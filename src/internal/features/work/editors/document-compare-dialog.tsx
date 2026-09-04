import { FileCheck2, FileUp, GitCompareArrows } from 'lucide-react';
import { type FormEvent, useId, useRef, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type {
  DocumentComparisonApplyResult,
  DocumentComparisonDiagnostic,
  DocumentComparisonMode,
} from '../work-document-compare';
import { OfficeFileInput, OfficeTextField } from './office-controls';

export interface DocumentCompareDialogRequest {
  author: string;
  file: File;
  mode: DocumentComparisonMode;
}

export interface DocumentCompareDialogProps {
  initialMode: DocumentComparisonMode;
  restoreFocusTarget: () => HTMLElement | null;
  onApplied: (result: DocumentComparisonApplyResult) => void;
  onClose: () => void;
  onSubmit: (
    request: DocumentCompareDialogRequest,
  ) => Promise<DocumentComparisonApplyResult>;
}

export function DocumentCompareDialog({
  initialMode,
  restoreFocusTarget,
  onApplied,
  onClose,
  onSubmit,
}: DocumentCompareDialogProps) {
  const [mode, setMode] = useState(initialMode);
  const [file, setFile] = useState<File | null>(null);
  const [author, setAuthor] = useState('审阅者');
  const [authorTouched, setAuthorTouched] = useState(false);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<DocumentComparisonApplyResult | null>(
    null,
  );
  const fileInputRef = useRef<HTMLInputElement>(null);
  const formId = useId();
  const authorId = useId();
  const submit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!file || busy || (mode === 'compare' && !author.trim())) return;
    setBusy(true);
    setResult(null);
    try {
      const next = await onSubmit({ author: author.trim(), file, mode });
      setResult(next);
      if (next.status === 'applied') onApplied(next);
    } catch {
      setResult({
        status: 'unsupported',
        summary: {
          deletions: 0,
          formatting: 0,
          insertions: 0,
          paragraphFormatting: 0,
        },
        diagnostics: [
          {
            code: 'invalid-revised-content',
            message: 'The selected file could not be imported.',
          },
        ],
      });
    } finally {
      setBusy(false);
    }
  };
  const selectFile = (next: File) => {
    setFile(next);
    setResult(null);
    if (!authorTouched) setAuthor(fileStem(next.name) || '审阅者');
  };

  return (
    <Dialog
      title="比较与合并文档"
      description="把另一版本转换为可逐项接受或拒绝的 Writer 修订。"
      className="work-document-compare-dialog"
      closeDisabled={busy}
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" disabled={busy} onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!file || busy || (mode === 'compare' && !author.trim())}
          >
            {busy
              ? '正在处理…'
              : mode === 'compare'
                ? '生成比较结果'
                : '合并修订'}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-document-compare-mode">
          <legend>处理方式</legend>
          <label>
            <input
              type="radio"
              name="document-comparison-mode"
              value="compare"
              checked={mode === 'compare'}
              onChange={() => {
                setMode('compare');
                setResult(null);
              }}
            />
            <span>
              <GitCompareArrows size={18} aria-hidden="true" />
              <strong>比较</strong>
              <small>当前文档作为原稿，导入文件作为修订稿。</small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="document-comparison-mode"
              value="combine"
              checked={mode === 'combine'}
              onChange={() => {
                setMode('combine');
                setResult(null);
              }}
            />
            <span>
              <FileCheck2 size={18} aria-hidden="true" />
              <strong>合并</strong>
              <small>导入带修订的审阅副本，并核验其原始基线。</small>
            </span>
          </label>
        </fieldset>

        <OfficeFileInput
          ref={fileInputRef}
          accept=".docx,.html,.htm,.txt,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/html,text/plain"
          aria-label={
            mode === 'compare' ? '选择修订版本文件' : '选择审阅副本文件'
          }
          disabled={busy}
          onFileSelect={selectFile}
        />
        <button
          type="button"
          className="work-document-compare-file"
          data-autofocus="true"
          disabled={busy}
          onClick={() => fileInputRef.current?.click()}
        >
          <FileUp size={20} aria-hidden="true" />
          <span>
            <strong>
              {file?.name ??
                (mode === 'compare' ? '选择修订版本' : '选择带修订的审阅副本')}
            </strong>
            <small>
              {file
                ? `${fileTypeLabel(file.name)} · ${formatFileSize(file.size)}`
                : '支持 DOCX、HTML 和 TXT；当前文档不会在预检失败时改变。'}
            </small>
          </span>
        </button>

        {mode === 'compare' && (
          <label className="work-document-compare-author" htmlFor={authorId}>
            <span>修订者名称</span>
            <OfficeTextField
              id={authorId}
              aria-label="比较结果修订者名称"
              value={author}
              disabled={busy}
              maxLength={256}
              onChange={(event) => {
                setAuthorTouched(true);
                setAuthor(event.currentTarget.value);
              }}
            />
            <small>该名称会显示在生成的插入、删除、移动与格式修订中。</small>
          </label>
        )}

        <ComparisonBoundary mode={mode} />
        {result?.status === 'unchanged' && (
          <p className="work-document-compare-status" role="status">
            两份文档在支持的比较范围内完全一致，没有生成修订。
          </p>
        )}
        {result?.status === 'unsupported' && (
          <div className="work-document-compare-errors" role="alert">
            <strong>无法安全处理这份文档</strong>
            <ul>
              {result.diagnostics.map((diagnostic, index) => (
                <li key={`${diagnostic.code}-${index}`}>
                  {comparisonDiagnosticText(diagnostic)}
                </li>
              ))}
            </ul>
          </div>
        )}
      </form>
    </Dialog>
  );
}

function ComparisonBoundary({ mode }: { mode: DocumentComparisonMode }) {
  return (
    <div className="work-document-compare-boundary">
      <strong>{mode === 'compare' ? '确定性比较边界' : '安全合并边界'}</strong>
      <p>
        {mode === 'compare'
          ? '支持同一分节布局中的段落、标题、文字、格式差异，以及同一段内可安全识别的文本移动；复杂对象或节布局变化会明确停止。'
          : '审阅副本必须包含修订，且拒绝全部修订后与当前文档一致；现有修订须先处理。'}
      </p>
    </div>
  );
}

function comparisonDiagnosticText(
  diagnostic: DocumentComparisonDiagnostic,
): string {
  const location =
    diagnostic.section === undefined
      ? ''
      : `（第 ${diagnostic.section + 1} 节${
          diagnostic.block === undefined
            ? ''
            : `，差异块 ${diagnostic.block + 1}`
        }）`;
  const messages: Record<DocumentComparisonDiagnostic['code'], string> = {
    'changed-complex-structure':
      '检测到表格、图片、列表或其他复杂结构变化，未降级为纯文字。',
    'combine-baseline-mismatch': '审阅副本的原始基线与当前文档不一致。',
    'combine-resolution-invalid': '审阅副本包含损坏或无法拒绝的格式修订。',
    'combine-structural-revisions':
      '审阅副本改变了段落树；当前合并路径只接受行内和格式修订。',
    'combine-without-revisions': '审阅副本中没有可合并的修订。',
    'comparison-limit-exceeded': '文档超过本地有界比较限制。',
    'current-revisions-present': '当前文档仍有未处理修订，请先接受或拒绝。',
    'empty-structural-change': '空段落的结构变化无法承载可审阅文字修订。',
    'invalid-revised-content': '文件无法转换为当前 Writer 文档模型。',
    'revised-revisions-present': '修订稿已经包含修订，请改用“合并”。',
    'section-layout-mismatch': '分节数量或页面布局不一致。',
    'unsupported-inline-review-state':
      '导入内容包含无法安全迁移的批注或审阅状态。',
  };
  return `${messages[diagnostic.code]}${location}`;
}

function fileStem(name: string): string {
  return name.replace(/\.[^.]+$/, '').trim();
}

function fileTypeLabel(name: string): string {
  return name.split('.').at(-1)?.toLocaleUpperCase() || '文件';
}

function formatFileSize(bytes: number): string {
  if (bytes < 1_024) return `${bytes} B`;
  if (bytes < 1_048_576) return `${Math.max(0.1, bytes / 1_024).toFixed(1)} KB`;
  return `${Math.max(0.1, bytes / 1_048_576).toFixed(1)} MB`;
}
