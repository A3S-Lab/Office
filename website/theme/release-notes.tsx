import { useLang, useVersion } from '@rspress/core/runtime';
import type { ReactNode } from 'react';
import {
  officeReleaseNotesThroughVersion,
  type OfficeLocalizedReleaseText,
  type OfficeReleaseNote,
  type OfficeReleaseNoteKind,
  type OfficeReleaseSurface,
} from './release-notes-data';

type ReleaseNotesLanguage = 'en' | 'zh';

const COMPLETE_CHANGELOG_URL =
  'https://github.com/A3S-Lab/Office/blob/main/CHANGELOG.md';
const NPM_PACKAGE_URL = 'https://www.npmjs.com/package/@a3s-lab/office';

const KIND_LABELS: Record<OfficeReleaseNoteKind, OfficeLocalizedReleaseText> = {
  new: { en: 'New', zh: '新增' },
  improved: { en: 'Improved', zh: '改进' },
  fixed: { en: 'Fixed', zh: '修复' },
};

const SURFACE_LABELS: Record<OfficeReleaseSurface, OfficeLocalizedReleaseText> =
  {
    documentation: { en: 'Documentation', zh: '文档' },
    shared: { en: 'Shared UX', zh: '共享体验' },
    writer: { en: 'Writer', zh: '文字' },
    markdown: { en: 'Markdown', zh: 'Markdown' },
    spreadsheet: { en: 'Spreadsheet', zh: '表格' },
    presentation: { en: 'Presentation', zh: '演示文稿' },
    pdf: { en: 'PDF', zh: 'PDF' },
    playground: { en: 'Playground', zh: 'Playground' },
  };

export function ReleaseNotes() {
  const language: ReleaseNotesLanguage = useLang() === 'en' ? 'en' : 'zh';
  const activeVersion = useVersion() || 'latest';
  const releases = officeReleaseNotesThroughVersion(activeVersion);
  const isLatest = activeVersion === 'latest';

  return (
    <div className="office-release-notes">
      <section
        className="office-release-notes__scope"
        aria-label={language === 'en' ? 'Release notes scope' : '更新日志范围'}
      >
        <div className="office-release-notes__scope-copy">
          <span className="office-release-notes__live" aria-hidden="true">
            <i />
            {isLatest ? 'latest' : `v${activeVersion}`}
          </span>
          <div>
            <strong>
              {isLatest
                ? language === 'en'
                  ? 'Current release history'
                  : '当前版本历史'
                : language === 'en'
                  ? `History available in ${activeVersion}`
                  : `${activeVersion} 当时可用的历史`}
            </strong>
            <p>
              {isLatest
                ? language === 'en'
                  ? 'Newest first, with product outcomes before implementation details.'
                  : '按新到旧排列，先说明产品收益，再进入实现细节。'
                : language === 'en'
                  ? 'Later releases are intentionally hidden in this frozen documentation view.'
                  : '此冻结文档会有意隐藏后来才发布的版本。'}
            </p>
          </div>
        </div>
        <div className="office-release-notes__scope-links">
          <a href={NPM_PACKAGE_URL}>npm</a>
          <a href={COMPLETE_CHANGELOG_URL}>
            {language === 'en' ? 'Complete archive' : '完整档案'}
          </a>
        </div>
      </section>

      {releases.length > 0 ? (
        <ol
          className="office-release-notes__timeline"
          aria-label={language === 'en' ? 'Recent releases' : '近期版本'}
        >
          {releases.map((release, index) => (
            <li key={release.version}>
              <ReleaseCard
                language={language}
                release={release}
                featured={index === 0}
              />
            </li>
          ))}
        </ol>
      ) : (
        <section className="office-release-notes__empty">
          <SurfaceIcon surface="documentation" />
          <div>
            <h2>
              {language === 'en'
                ? 'The visual history starts at 0.30.0'
                : '可视化版本历史从 0.30.0 开始'}
            </h2>
            <p>
              {language === 'en'
                ? 'This older documentation remains frozen. Use the complete repository changelog for its exhaustive engineering record.'
                : '当前旧版文档继续保持冻结；其完整工程记录可在仓库更新日志中查看。'}
            </p>
            <a href={COMPLETE_CHANGELOG_URL}>
              {language === 'en'
                ? 'Browse the complete changelog'
                : '浏览完整更新日志'}
            </a>
          </div>
        </section>
      )}

      <aside className="office-release-notes__archive">
        <div className="office-release-notes__archive-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path d="M6.75 3.75h10.5a2 2 0 0 1 2 2v12.5a2 2 0 0 1-2 2H6.75a2 2 0 0 1-2-2V5.75a2 2 0 0 1 2-2Z" />
            <path d="M8 8h8M8 12h8M8 16h5" />
          </svg>
        </div>
        <div>
          <strong>
            {language === 'en'
              ? 'Need the engineering-level record?'
              : '需要工程级完整记录？'}
          </strong>
          <p>
            {language === 'en'
              ? 'The repository changelog retains every compatibility boundary, test gate, and implementation detail.'
              : '仓库更新日志保留全部兼容性边界、测试门禁和实现细节。'}
          </p>
        </div>
        <a href={COMPLETE_CHANGELOG_URL}>
          {language === 'en' ? 'Open CHANGELOG.md' : '打开 CHANGELOG.md'}
          <span aria-hidden="true">↗</span>
        </a>
      </aside>
    </div>
  );
}

function ReleaseCard({
  featured,
  language,
  release,
}: {
  featured: boolean;
  language: ReleaseNotesLanguage;
  release: OfficeReleaseNote;
}) {
  const primarySurface = release.surfaces[0] ?? 'shared';
  const resourcesLabel =
    language === 'en'
      ? `${release.version} release resources`
      : `${release.version} 版本资源`;

  return (
    <article
      className={`office-release-card${
        featured ? ' office-release-card--featured' : ''
      }`}
      data-kind={release.kind}
      data-surface={primarySurface}
      data-version={release.version}
    >
      <header className="office-release-card__header">
        <SurfaceIcon surface={primarySurface} />
        <div className="office-release-card__identity">
          <div className="office-release-card__version-row">
            <span className="office-release-card__version">
              v{release.version}
            </span>
            <time dateTime={release.date}>
              {formatReleaseDate(release.date, language)}
            </time>
          </div>
          <ul
            className="office-release-card__surfaces"
            aria-label={language === 'en' ? 'Editor surfaces' : '编辑器范围'}
          >
            <li data-kind={release.kind}>
              {localized(KIND_LABELS[release.kind], language)}
            </li>
            {release.surfaces.map((surface) => (
              <li key={surface} data-surface={surface}>
                {localized(SURFACE_LABELS[surface], language)}
              </li>
            ))}
          </ul>
        </div>
      </header>

      <div className="office-release-card__copy">
        <h2>{localized(release.title, language)}</h2>
        <p>{localized(release.summary, language)}</p>
      </div>

      <ul className="office-release-card__highlights">
        {release.highlights.map((highlight) => (
          <li key={highlight.title.en}>
            <strong>{localized(highlight.title, language)}</strong>
            <span>{localized(highlight.detail, language)}</span>
          </li>
        ))}
      </ul>

      <nav className="office-release-card__links" aria-label={resourcesLabel}>
        {release.links.map((link) => {
          const href = localized(link.href, language);
          return (
            <a key={href} href={href}>
              {localized(link.label, language)}
              <span aria-hidden="true">
                {href.startsWith('http') ? '↗' : '→'}
              </span>
            </a>
          );
        })}
      </nav>
    </article>
  );
}

function SurfaceIcon({ surface }: { surface: OfficeReleaseSurface }) {
  let paths: ReactNode;
  switch (surface) {
    case 'writer':
      paths = (
        <>
          <path d="M7 3.75h7l3 3v13.5H7z" />
          <path d="M14 3.75v3h3M9.5 11h5M9.5 14h5M9.5 17h3.5" />
        </>
      );
      break;
    case 'spreadsheet':
      paths = (
        <>
          <rect x="3.75" y="4.25" width="16.5" height="15.5" rx="2" />
          <path d="M3.75 9h16.5M9.25 9v10.75M14.75 9v10.75" />
        </>
      );
      break;
    case 'presentation':
      paths = (
        <>
          <rect x="3.75" y="4.25" width="16.5" height="11.5" rx="2" />
          <path d="m9 19.75 3-4 3 4M8 8h8M8 11h5" />
        </>
      );
      break;
    case 'pdf':
      paths = (
        <>
          <path d="M7 3.75h7l3 3v13.5H7z" />
          <path d="M14 3.75v3h3M9 11h2a1.5 1.5 0 0 1 0 3H9v-3Zm5 3v-3h1.25a1.5 1.5 0 0 1 0 3H14Zm0 0v2" />
        </>
      );
      break;
    case 'playground':
      paths = (
        <>
          <rect x="3.75" y="4.25" width="16.5" height="15.5" rx="2" />
          <path d="M3.75 8.5h16.5M7 6.35h.01M9.5 6.35h.01M8 12.25l2.25 2L8 16.25M12.75 16.25H16" />
        </>
      );
      break;
    case 'markdown':
      paths = (
        <>
          <rect x="3.25" y="5" width="17.5" height="14" rx="2" />
          <path d="M6.5 15v-6l2.75 3 2.75-3v6M15 12h2.5M16.25 9.75V15" />
        </>
      );
      break;
    case 'documentation':
      paths = (
        <>
          <path d="M4.5 5.5A2.5 2.5 0 0 1 7 3h4.5v16H7a2.5 2.5 0 0 0-2.5 2.5zM19.5 5.5A2.5 2.5 0 0 0 17 3h-4.5v16H17a2.5 2.5 0 0 1 2.5 2.5z" />
          <path d="M7.5 7h2M7.5 10h2M14.5 7h2M14.5 10h2" />
        </>
      );
      break;
    default:
      paths = (
        <>
          <circle cx="7" cy="12" r="3.25" />
          <circle cx="17" cy="7" r="3.25" />
          <circle cx="17" cy="17" r="3.25" />
          <path d="m9.75 10.25 4.5-2M9.75 13.75l4.5 2" />
        </>
      );
  }

  return (
    <span className="office-release-card__icon" aria-hidden="true">
      <svg viewBox="0 0 24 24" aria-hidden="true">
        {paths}
      </svg>
    </span>
  );
}

function localized(
  value: OfficeLocalizedReleaseText,
  language: ReleaseNotesLanguage,
): string {
  return value[language];
}

function formatReleaseDate(
  value: string,
  language: ReleaseNotesLanguage,
): string {
  const [year, month, day] = value.split('-').map(Number);
  if (!year || !month || !day) return value;
  if (language === 'zh') return `${year} 年 ${month} 月 ${day} 日`;
  const monthName = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ][month - 1];
  return monthName ? `${monthName} ${day}, ${year}` : value;
}
