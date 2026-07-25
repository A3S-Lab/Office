import { Quote, Trash2 } from 'lucide-react';
import type { Ref } from 'react';
import { Button, InlineNotice } from '../../../design-system/primitives';
import {
  OfficeSelect,
  OfficeTextArea,
  OfficeTextField,
} from './office-controls';

export interface CitationSourceDraft {
  id?: string;
  tag: string;
  sourceType: string;
  title: string;
  year: string;
  authors: string;
  corporateAuthor: string;
  publisher: string;
  city: string;
  journalName: string;
  volume: string;
  issue: string;
  pages: string;
  url: string;
  standardNumber: string;
  conferenceName: string;
  institution: string;
}

const SOURCE_TYPES = [
  ['Book', '书籍'],
  ['BookSection', '书籍章节'],
  ['JournalArticle', '期刊文章'],
  ['ArticleInAPeriodical', '报刊文章'],
  ['ConferenceProceedings', '会议论文'],
  ['Report', '报告'],
  ['InternetSite', '网站'],
  ['DocumentFromInternetSite', '网页文档'],
  ['ElectronicSource', '电子资源'],
  ['Misc', '其他'],
] as const;

export function DocumentCitationSourceForm({
  draft,
  dirty,
  error,
  tagInputRef,
  onDraftChange,
  onSave,
  onInsert,
  onDelete,
}: {
  draft: CitationSourceDraft;
  dirty: boolean;
  error: string;
  tagInputRef?: Ref<HTMLInputElement>;
  onDraftChange: (draft: CitationSourceDraft) => void;
  onSave: () => void;
  onInsert: () => void;
  onDelete: () => void;
}) {
  const saved = Boolean(draft.id);
  const knownSourceType = SOURCE_TYPES.some(
    ([value]) => value === draft.sourceType,
  );
  const update = <Key extends keyof CitationSourceDraft>(
    key: Key,
    value: CitationSourceDraft[Key],
  ) => onDraftChange({ ...draft, [key]: value });

  return (
    <form
      aria-label={saved ? '编辑文献' : '新建文献'}
      onSubmit={(event) => {
        event.preventDefault();
        onSave();
      }}
    >
      <div className="work-document-citation-form-heading wide">
        <strong>{saved ? '编辑文献' : '新建文献'}</strong>
        <span>标题和简称为必填</span>
      </div>
      <div className="work-office-field">
        <span>简称</span>
        <OfficeTextField
          ref={tagInputRef}
          aria-label="文献简称"
          value={draft.tag}
          maxLength={80}
          placeholder="例如 Smith2026"
          onChange={(event) => update('tag', event.target.value)}
        />
      </div>
      <div className="work-office-field">
        <span>类型</span>
        <OfficeSelect
          ariaLabel="文献类型"
          value={draft.sourceType}
          options={[
            ...(!knownSourceType && draft.sourceType
              ? [
                  {
                    value: draft.sourceType,
                    label: `${draft.sourceType}（原始类型）`,
                  },
                ]
              : []),
            ...SOURCE_TYPES.map(([value, label]) => ({ value, label })),
          ]}
          onValueChange={(sourceType) => update('sourceType', sourceType)}
        />
      </div>
      <div className="work-office-field wide">
        <span>标题</span>
        <OfficeTextField
          aria-label="文献标题"
          value={draft.title}
          onChange={(event) => update('title', event.target.value)}
        />
      </div>
      <div className="work-office-field">
        <span>年份</span>
        <OfficeTextField
          aria-label="文献年份"
          value={draft.year}
          inputMode="numeric"
          placeholder="2026"
          onChange={(event) => update('year', event.target.value)}
        />
      </div>
      <div className="work-office-field">
        <span>机构作者</span>
        <OfficeTextField
          aria-label="机构作者"
          value={draft.corporateAuthor}
          placeholder="与个人作者二选一"
          onChange={(event) => update('corporateAuthor', event.target.value)}
        />
      </div>
      <div className="work-office-field wide">
        <span>个人作者</span>
        <OfficeTextArea
          aria-label="个人作者"
          value={draft.authors}
          placeholder={'每行一位，例如：\nSmith, Jane\nLi, Ming'}
          onChange={(event) => update('authors', event.target.value)}
        />
      </div>
      <details className="work-document-citation-more-fields wide">
        <summary>更多出版信息</summary>
        <div>
          <div className="work-office-field">
            <span>出版者</span>
            <OfficeTextField
              aria-label="出版者"
              value={draft.publisher}
              onChange={(event) => update('publisher', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>出版城市</span>
            <OfficeTextField
              aria-label="出版城市"
              value={draft.city}
              onChange={(event) => update('city', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>期刊名</span>
            <OfficeTextField
              aria-label="期刊名"
              value={draft.journalName}
              onChange={(event) => update('journalName', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>卷 / 期</span>
            <span className="paired">
              <OfficeTextField
                aria-label="卷"
                value={draft.volume}
                placeholder="卷"
                onChange={(event) => update('volume', event.target.value)}
              />
              <OfficeTextField
                aria-label="期"
                value={draft.issue}
                placeholder="期"
                onChange={(event) => update('issue', event.target.value)}
              />
            </span>
          </div>
          <div className="work-office-field">
            <span>页码</span>
            <OfficeTextField
              aria-label="文献页码"
              value={draft.pages}
              placeholder="12–28"
              onChange={(event) => update('pages', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>ISBN / DOI</span>
            <OfficeTextField
              aria-label="标准编号"
              value={draft.standardNumber}
              onChange={(event) => update('standardNumber', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>会议名称</span>
            <OfficeTextField
              aria-label="会议名称"
              value={draft.conferenceName}
              onChange={(event) => update('conferenceName', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>报告机构</span>
            <OfficeTextField
              aria-label="报告机构"
              value={draft.institution}
              onChange={(event) => update('institution', event.target.value)}
            />
          </div>
          <div className="work-office-field wide">
            <span>网址</span>
            <OfficeTextField
              aria-label="文献网址"
              value={draft.url}
              inputMode="url"
              placeholder="https://"
              onChange={(event) => update('url', event.target.value)}
            />
          </div>
        </div>
      </details>
      <div className="actions wide">
        {error && (
          <InlineNotice
            className="work-office-form-error"
            tone="danger"
            role="alert"
          >
            {error}
          </InlineNotice>
        )}
        <div className="work-document-citation-form-buttons">
          {saved && (
            <Button tone="danger" aria-label="删除文献" onClick={onDelete}>
              <Trash2 size={13} />
              删除
            </Button>
          )}
          <span aria-hidden="true" />
          {saved && (
            <Button tone="secondary" disabled={dirty} onClick={onInsert}>
              <Quote size={13} />
              插入引文
            </Button>
          )}
          <Button
            type="submit"
            tone="primary"
            aria-label="保存文献"
            disabled={!dirty}
          >
            保存
          </Button>
        </div>
      </div>
    </form>
  );
}
