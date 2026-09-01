export type OfficeReleaseNoteKind = 'new' | 'improved' | 'fixed';

export type OfficeReleaseSurface =
  | 'documentation'
  | 'shared'
  | 'writer'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf'
  | 'playground';

export interface OfficeLocalizedReleaseText {
  en: string;
  zh: string;
}

export interface OfficeReleaseNoteHighlight {
  title: OfficeLocalizedReleaseText;
  detail: OfficeLocalizedReleaseText;
}

export interface OfficeReleaseNoteLink {
  href: OfficeLocalizedReleaseText;
  label: OfficeLocalizedReleaseText;
}

export interface OfficeReleaseNote {
  version: string;
  date: string;
  kind: OfficeReleaseNoteKind;
  surfaces: OfficeReleaseSurface[];
  title: OfficeLocalizedReleaseText;
  summary: OfficeLocalizedReleaseText;
  highlights: [
    OfficeReleaseNoteHighlight,
    OfficeReleaseNoteHighlight,
    OfficeReleaseNoteHighlight,
  ];
  links: OfficeReleaseNoteLink[];
}

export const OFFICE_RELEASE_NOTES: readonly OfficeReleaseNote[] = [
  {
    version: '0.38.1',
    date: '2026-09-01',
    kind: 'improved',
    surfaces: ['documentation', 'shared'],
    title: {
      en: 'Release notes that explain the product change',
      zh: '真正说明产品变化的更新日志',
    },
    summary: {
      en: 'The documentation now turns recent releases into a version-aware, bilingual story instead of a flat list of implementation links.',
      zh: '文档现在把近期版本组织成支持双语和版本冻结的产品故事，不再只提供扁平的实现链接列表。',
    },
    highlights: [
      {
        title: { en: 'Scannable hierarchy', zh: '可快速扫读' },
        detail: {
          en: 'Version, date, change type, editor surface, outcome, and evidence are visible before opening a deep reference.',
          zh: '无需先打开深层参考，即可看到版本、日期、变更类型、编辑器、用户收益与验证入口。',
        },
      },
      {
        title: { en: 'Frozen history', zh: '冻结历史' },
        detail: {
          en: 'Selecting an older documentation version hides every release that did not exist at that point.',
          zh: '切换到旧版文档时，会自动隐藏当时尚未发布的版本，避免历史页面漂移。',
        },
      },
      {
        title: { en: 'One complete archive', zh: '一份完整档案' },
        detail: {
          en: 'Curated release stories link back to the exhaustive repository changelog and immutable GitHub Releases.',
          zh: '面向用户的发布摘要继续链接到完整仓库日志和不可变的 GitHub Release。',
        },
      },
    ],
    links: [
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/blob/main/CHANGELOG.md',
          zh: 'https://github.com/A3S-Lab/Office/blob/main/CHANGELOG.md',
        },
        label: { en: 'Open the complete changelog', zh: '查看完整更新日志' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.38.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.38.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.38.0',
    date: '2026-09-01',
    kind: 'new',
    surfaces: ['writer'],
    title: {
      en: 'Native OpenType typography in Writer',
      zh: 'Writer 原生 OpenType 排版',
    },
    summary: {
      en: 'Ordinary text and structured equations now share one bounded Office 2010 typography model with exact DOCX reopen behavior.',
      zh: '普通文字与结构化公式现在共用一套有界的 Office 2010 排版模型，并支持精确 DOCX 重开。',
    },
    highlights: [
      {
        title: { en: 'Complete controls', zh: '完整控制项' },
        detail: {
          en: 'All 16 ligature combinations, numeral forms and spacing, stylistic sets 1–20, and contextual alternates are editable.',
          zh: '可编辑全部 16 种连字组合、数字形式与间距、样式集 1–20 和上下文替代。',
        },
      },
      {
        title: { en: 'Mixed-selection safe', zh: '混合选区安全' },
        detail: {
          en: 'The advanced Font dialog changes only touched properties, restores the selection, and creates one Undo record.',
          zh: '高级字体弹窗只修改用户触碰的属性，恢复选区，并只生成一条撤销记录。',
        },
      },
      {
        title: { en: 'Native fidelity', zh: '原生保真' },
        detail: {
          en: 'Body, headers, footers, notes, comments, styles, Format Painter, and formatting revisions retain the same model.',
          zh: '正文、页眉页脚、脚注尾注、批注、样式、格式刷与格式修订都保留同一模型。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#native-opentype-typography',
          zh: './components/document.html#原生-opentype-排版',
        },
        label: { en: 'Read the Writer guide', zh: '阅读 Writer 指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.38.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.38.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.5',
    date: '2026-09-01',
    kind: 'improved',
    surfaces: ['spreadsheet'],
    title: {
      en: 'XLSX 1900 and 1904 dates stay exact',
      zh: 'XLSX 1900 与 1904 日期保持精确',
    },
    summary: {
      en: 'Spreadsheet now owns the workbook date epoch, preserving exact native serials through filtering, collaboration, export, and reopen.',
      zh: 'Spreadsheet 现在显式持有工作簿日期纪元，并在筛选、协作、导出与重开中保留精确原生序列。',
    },
    highlights: [
      {
        title: { en: 'No timezone drift', zh: '无时区漂移' },
        detail: {
          en: 'Date-typed scalar values and formula caches remain numeric instead of passing through JavaScript Date conversion.',
          zh: '日期类型的标量值和公式缓存保持数值形式，不再经过 JavaScript Date 转换。',
        },
      },
      {
        title: { en: 'One workbook epoch', zh: '统一工作簿纪元' },
        detail: {
          en: 'Dynamic filters, menu profiling, sort reconciliation, and current-date authoring consume the same 1900 or 1904 setting.',
          zh: '动态筛选、菜单分析、排序后协调和当前日期输入共用同一 1900 或 1904 设置。',
        },
      },
      {
        title: { en: 'Real-package proof', zh: '真实文件验证' },
        detail: {
          en: 'Import, export, reopen, 1904 serial zero, collaboration, and browser workflows are covered together.',
          zh: '导入、导出、重开、1904 序列 0、协作和浏览器工作流得到联合验证。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#xlsx-1904-date-system-retention',
          zh: './components/spreadsheet.html#xlsx-1904-日期系统保留',
        },
        label: {
          en: 'Read the Spreadsheet guide',
          zh: '阅读 Spreadsheet 指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.5',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.5',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.4',
    date: '2026-08-31',
    kind: 'improved',
    surfaces: ['spreadsheet'],
    title: {
      en: 'Custom Sort respects Tables and AutoFilter',
      zh: '自定义排序理解表格与 AutoFilter 所属区域',
    },
    summary: {
      en: 'Sorting from inside a native table or worksheet filter now resolves the exact structural range and reconciles visibility after rows move.',
      zh: '从原生表格或工作表筛选区域内部排序时，会解析精确结构范围，并在行移动后重新协调可见性。',
    },
    highlights: [
      {
        title: { en: 'Structural boundaries', zh: '结构边界' },
        detail: {
          en: 'Headers stay fixed, enabled totals rows stay outside the sort, and left-to-right movement is disabled for owned ranges.',
          zh: '表头保持固定，启用的汇总行不参与排序，所属区域禁止从左到右移动。',
        },
      },
      {
        title: { en: 'Filter-safe movement', zh: '筛选安全移动' },
        detail: {
          en: 'Typed filters are reevaluated while opaque hidden-row ownership follows the stable sort permutation.',
          zh: '类型化筛选会重新计算，不透明隐藏行的所有权则跟随稳定排序置换。',
        },
      },
      {
        title: { en: 'Bounded execution', zh: '有界执行' },
        detail: {
          en: 'Dense and sparse reads share explicit budgets, fail-closed fingerprints, one controlled update, and one-step Undo.',
          zh: '稠密与稀疏读取共享明确预算、失败关闭指纹、一次受控更新和单步撤销。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#multi-key-custom-sort',
          zh: './components/spreadsheet.html#多关键字自定义排序',
        },
        label: { en: 'Read the Custom Sort guide', zh: '阅读自定义排序指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.4',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.4',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.3',
    date: '2026-08-31',
    kind: 'new',
    surfaces: ['spreadsheet'],
    title: {
      en: 'Manage reusable Custom Lists',
      zh: '管理可复用的自定义序列',
    },
    summary: {
      en: 'Custom Sort gains a responsive preference manager for built-in and user-authored month, weekday, and domain-specific sequences.',
      zh: '自定义排序新增响应式偏好管理器，可管理内置月份、星期及用户定义的业务序列。',
    },
    highlights: [
      {
        title: { en: 'Staged editing', zh: '暂存式编辑' },
        detail: {
          en: 'Create, edit, delete, and reorder up to 32 user lists before one atomic confirmation.',
          zh: '可先创建、编辑、删除和重排最多 32 个用户序列，再一次性确认。',
        },
      },
      {
        title: { en: 'Coherent sort keys', zh: '排序键保持一致' },
        detail: {
          en: 'Active sort levels follow edited lists and fall back safely when a referenced list is removed.',
          zh: '活动排序层级会跟随序列编辑，并在引用序列被删除时安全回退。',
        },
      },
      {
        title: { en: 'Host-owned preferences', zh: '宿主持有偏好' },
        detail: {
          en: 'Optional typed persistence never leaks personal list preferences into controlled workbook content.',
          zh: '可选的类型化持久化不会把个人序列偏好写入受控工作簿内容。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#multi-key-custom-sort',
          zh: './components/spreadsheet.html#多关键字自定义排序',
        },
        label: { en: 'Read the Custom Sort guide', zh: '阅读自定义排序指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.3',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.3',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.2',
    date: '2026-08-31',
    kind: 'fixed',
    surfaces: ['writer', 'markdown', 'shared'],
    title: {
      en: 'Controlled Chinese IME input commits once',
      zh: '受控中文输入法只提交一次最终文本',
    },
    summary: {
      en: 'Pinyin and other pre-edit values remain local until composition settles, while authoritative host replacements wait for the same boundary.',
      zh: '拼音等预编辑值在组合完成前保持本地，权威宿主替换也等待同一组合边界后再协调。',
    },
    highlights: [
      {
        title: { en: 'No phonetic leakage', zh: '不再泄漏拼音' },
        detail: {
          en: 'Document and visual Markdown publish only the committed Chinese value instead of intermediate composition text.',
          zh: 'Document 与可视化 Markdown 只发布最终中文，不再发布中间组合文本。',
        },
      },
      {
        title: { en: 'Controlled reconciliation', zh: '受控状态协调' },
        detail: {
          en: 'A host update received mid-composition is applied only after the local input method finishes.',
          zh: '组合期间收到的宿主更新会在本地输入法结束后再应用。',
        },
      },
      {
        title: { en: 'WebKit release gate', zh: 'WebKit 发布门禁' },
        detail: {
          en: 'A dedicated browser contract verifies zero pre-edit publications and exactly one final change.',
          zh: '专用浏览器契约验证预编辑阶段零次发布，并且最终只产生一次变化。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#ime-and-controlled-updates',
          zh: './components/document.html#输入法与受控更新',
        },
        label: {
          en: 'Read the Document IME contract',
          zh: '阅读 Document 输入法约定',
        },
      },
      {
        href: {
          en: './components/markdown.html#visual-editor-ime-behavior',
          zh: './components/markdown.html#可视化编辑器的输入法行为',
        },
        label: {
          en: 'Read the Markdown IME contract',
          zh: '阅读 Markdown 输入法约定',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.2',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.2',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.1',
    date: '2026-08-31',
    kind: 'improved',
    surfaces: ['shared', 'presentation', 'spreadsheet'],
    title: {
      en: 'Sharper File actions, Presentation IME, and local filters',
      zh: '文件操作、演示文稿输入法与本地筛选体验升级',
    },
    summary: {
      en: 'The shared shell gains clearer action semantics while Presentation composition and Spreadsheet filtering become dependable offline workflows.',
      zh: '共享外壳获得更清晰的操作语义，同时演示文稿组合输入与表格筛选成为可靠的离线工作流。',
    },
    highlights: [
      {
        title: { en: 'Readable File menu', zh: '清晰的文件菜单' },
        detail: {
          en: 'Explicit icons, danger treatment, bounded scrolling, keyboard focus, and legible disabled states clarify host-owned actions.',
          zh: '明确图标、危险操作样式、有界滚动、键盘焦点和清晰禁用态让宿主操作更易理解。',
        },
      },
      {
        title: { en: 'Presentation composition', zh: '演示文稿组合输入' },
        detail: {
          en: 'Slide text keeps pre-edit input local and editor shortcuts yield while an input method is active.',
          zh: '幻灯片文字把预编辑输入保留在本地，输入法活动时编辑器快捷键主动让行。',
        },
      },
      {
        title: { en: 'Native local filters', zh: '原生本地筛选' },
        detail: {
          en: 'Wildcard, Top/Bottom, compound custom, and typed AutoFilter criteria author and round-trip without a cloud dependency.',
          zh: '通配符、前后若干项、复合自定义及类型化 AutoFilter 条件均可离线创作并往返。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/react.html#host-owned-file-actions',
          zh: './components/react.html#宿主持有的文件操作',
        },
        label: { en: 'Read the File action contract', zh: '阅读文件操作约定' },
      },
      {
        href: {
          en: './components/spreadsheet.html#worksheet-autofilter-conditions',
          zh: './components/spreadsheet.html#工作表-autofilter-条件',
        },
        label: { en: 'Read the AutoFilter guide', zh: '阅读 AutoFilter 指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.37.0',
    date: '2026-08-30',
    kind: 'new',
    surfaces: ['spreadsheet'],
    title: {
      en: 'Native Spreadsheet totals rows',
      zh: '原生 Spreadsheet 汇总行',
    },
    summary: {
      en: 'Table creation and Table Design can author filter-aware native totals functions, labels, or bounded custom formulas.',
      zh: '表格创建与表格设计现在可创作支持筛选的原生汇总函数、标签或有界自定义公式。',
    },
    highlights: [
      {
        title: { en: 'Common functions', zh: '常用函数' },
        detail: {
          en: 'Sum, average, count, extrema, variance, standard deviation, labels, and custom formulas remain typed.',
          zh: '求和、平均值、计数、极值、方差、标准差、标签和自定义公式都保持类型化。',
        },
      },
      {
        title: { en: 'Filter-aware calculation', zh: '理解筛选的计算' },
        detail: {
          en: 'Generated native totals use SUBTOTAL semantics shared by Rust, WebAssembly, Worker, and JavaScript paths.',
          zh: '生成的原生汇总使用由 Rust、WebAssembly、Worker 与 JavaScript 共用的 SUBTOTAL 语义。',
        },
      },
      {
        title: { en: 'Exact XLSX round trip', zh: '精确 XLSX 往返' },
        detail: {
          en: 'Totals functions, labels, formulas, dense/sparse ownership, and collaboration state survive export and reopen.',
          zh: '汇总函数、标签、公式、稠密/稀疏所有权和协作状态都能经受导出与重开。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#native-totals-row-authoring',
          zh: './components/spreadsheet.html#原生汇总行创作',
        },
        label: { en: 'Read the Tables guide', zh: '阅读表格指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.37.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.36.0',
    date: '2026-08-30',
    kind: 'new',
    surfaces: ['spreadsheet'],
    title: {
      en: 'Structured references calculate and fill',
      zh: '结构化引用参与计算并自动填充',
    },
    summary: {
      en: 'Native table formulas now calculate through bounded Rust, WebAssembly, Worker, and JavaScript paths while inserted rows inherit safe calculated-column rules.',
      zh: '原生表格公式现在可通过有界的 Rust、WebAssembly、Worker 与 JavaScript 路径计算，插入行也能安全继承计算列规则。',
    },
    highlights: [
      {
        title: { en: 'Native table syntax', zh: '原生表格语法' },
        detail: {
          en: 'Table columns, current-row formulas, worksheet qualifiers, and common row selections share one parser and dependency graph.',
          zh: '表格列、当前行公式、工作表限定符与常见行选择器共用同一解析器和依赖图。',
        },
      },
      {
        title: { en: 'Safe calculated columns', zh: '安全计算列' },
        detail: {
          en: 'Only empty cells in newly inserted body rows receive an inferred formula; manual exceptions and existing values stay authoritative.',
          zh: '只有新插入正文行中的空单元格会接收推断公式，手动例外与已有值始终保持权威。',
        },
      },
      {
        title: { en: 'Explicit bounds', zh: '明确边界' },
        detail: {
          en: 'Table count, materialization, unsupported references, collaboration validation, and XLSX metadata all fail closed.',
          zh: '表格数量、物化范围、不支持的引用、协作验证与 XLSX 元数据都采用失败关闭策略。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#structured-reference-calculation',
          zh: './components/spreadsheet.html#结构化引用计算',
        },
        label: {
          en: 'Read the structured-reference guide',
          zh: '阅读结构化引用指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.36.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.36.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.35.0',
    date: '2026-08-29',
    kind: 'improved',
    surfaces: ['playground', 'shared'],
    title: {
      en: 'An immersive Playground with release context',
      zh: '带版本上下文的沉浸式 Playground',
    },
    summary: {
      en: 'The public Playground becomes one full-viewport workspace with a typed, filterable release gallery and more reliable first-open editor focus.',
      zh: '公开 Playground 现在是一体化全视口工作区，提供类型化、可筛选的版本能力画廊，并提升编辑器首次打开时的焦点可靠性。',
    },
    highlights: [
      {
        title: { en: 'One workspace', zh: '一体化工作区' },
        detail: {
          en: 'Duplicated global chrome is removed so the editor, recent files, and creation flows share the available viewport.',
          zh: '移除重复的全局外壳，让编辑器、最近文件和创建流程共享完整可用视口。',
        },
      },
      {
        title: { en: 'Typed release gallery', zh: '类型化版本画廊' },
        detail: {
          en: 'Every featured workflow owns one launch target, editor type, and release label in a responsive filterable grid.',
          zh: '每个精选工作流都在响应式筛选网格中持有唯一启动目标、编辑器类型和版本标签。',
        },
      },
      {
        title: { en: 'Stable first interaction', zh: '稳定首次交互' },
        detail: {
          en: 'Blank Presentation titles and live Spreadsheet selections retain intentional focus instead of dropping the first action.',
          zh: '空白演示文稿标题与表格实时选区会保留预期焦点，不再丢失第一次操作。',
        },
      },
    ],
    links: [
      {
        href: {
          en: 'https://a3s-lab.github.io/Office/playground/',
          zh: 'https://a3s-lab.github.io/Office/playground/',
        },
        label: { en: 'Open the Playground', zh: '打开 Playground' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.35.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.35.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.34.0',
    date: '2026-08-25',
    kind: 'new',
    surfaces: ['presentation'],
    title: {
      en: 'Presentation entrance animations',
      zh: '演示文稿入场动画',
    },
    summary: {
      en: 'Presentation can author, order, preview, play, collaborate on, and round-trip a bounded native entrance-animation subset.',
      zh: '演示文稿现在可创作、排序、预览、播放、协作并原生往返一组有界的入场动画。',
    },
    highlights: [
      {
        title: { en: 'Four effects', zh: '四种效果' },
        detail: {
          en: 'Appear, fade, fly-in, and zoom expose editable direction and timing where the effect supports them.',
          zh: '出现、淡入、飞入与缩放会在效果支持时提供可编辑方向和时序。',
        },
      },
      {
        title: { en: 'Three triggers', zh: '三种触发方式' },
        detail: {
          en: 'On-click, with-previous, and after-previous cues keep stable identities and deterministic playback order.',
          zh: '单击时、与上一动画同时、上一动画之后三种提示保持稳定身份和确定播放顺序。',
        },
      },
      {
        title: { en: 'Native timing trees', zh: '原生时序树' },
        detail: {
          en: 'Supported cues survive collaboration, copy/delete remapping, slideshow playback, PPTX export, and reopen.',
          zh: '支持的提示可经受协作、复制/删除重映射、放映、PPTX 导出与重开。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/presentation.html#entrance-animations',
          zh: './components/presentation.html#入场动画',
        },
        label: {
          en: 'Read the animation guide',
          zh: '阅读入场动画指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.34.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.34.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.33.0',
    date: '2026-08-25',
    kind: 'new',
    surfaces: ['pdf'],
    title: {
      en: 'PDF page organization',
      zh: 'PDF 页面组织',
    },
    summary: {
      en: 'PDF gains a responsive page organizer for structural edits while preserving Blob ownership, native history priority, and explicit safety limits.',
      zh: 'PDF 新增响应式页面组织器，可执行结构化编辑，同时保留 Blob 所有权、原生历史优先级和明确安全限制。',
    },
    highlights: [
      {
        title: { en: 'Complete page plan', zh: '完整页面计划' },
        detail: {
          en: 'Insert, delete, rotate, reorder, extract, merge, and split share one previewable operation model.',
          zh: '插入、删除、旋转、重排、抽取、合并与拆分共用同一套可预览操作模型。',
        },
      },
      {
        title: { en: 'Worker-owned mutation', zh: 'Worker 持有变更' },
        detail: {
          en: 'A lazy dedicated Worker returns one complete Blob and one page-history record per source mutation.',
          zh: '按需加载的独立 Worker 为每次源文件变更返回一个完整 Blob 和一条页面历史记录。',
        },
      },
      {
        title: { en: 'Fail-closed safety', zh: '失败关闭安全' },
        detail: {
          en: 'Source size, merge size, page count, encryption, signatures, catalog risk, and invalid plans are bounded before mutation.',
          zh: '源文件大小、合并大小、页数、加密、签名、目录风险与无效计划都会在变更前受限。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/pdf.html#page-organization',
          zh: './components/pdf.html#页面组织',
        },
        label: { en: 'Read the PDF guide', zh: '阅读 PDF 指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.33.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.33.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.32.0',
    date: '2026-08-25',
    kind: 'new',
    surfaces: ['writer'],
    title: {
      en: 'Document compare and combine',
      zh: '文档比较与合并',
    },
    summary: {
      en: 'Writer turns DOCX, HTML, or TXT differences into deterministic reviewable changes without replacing the current document.',
      zh: 'Writer 可把 DOCX、HTML 或 TXT 差异转成确定、可审阅的修订，同时不会替换当前文档。',
    },
    highlights: [
      {
        title: { en: 'Reviewable differences', zh: '可审阅差异' },
        detail: {
          en: 'Insertions, deletions, character formatting, and paragraph formatting arrive in one transaction and one Undo record.',
          zh: '插入、删除、字符格式与段落格式通过一次事务和一条撤销记录进入文档。',
        },
      },
      {
        title: { en: 'Deterministic identity', zh: '确定身份' },
        detail: {
          en: 'Stable identities, author attribution, and bounded diff matrices make review decisions reproducible.',
          zh: '稳定身份、作者归属和有界差异矩阵让审阅决策可以复现。',
        },
      },
      {
        title: { en: 'Safe combine', zh: '安全合并' },
        detail: {
          en: 'Combining requires an exact reject-all baseline and refuses unsupported structures or unresolved current revisions.',
          zh: '合并要求精确的全部拒绝基线，并拒绝不支持的结构或尚未解决的当前修订。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#document-compare-and-combine',
          zh: './components/document.html#文档比较与合并',
        },
        label: {
          en: 'Read the compare guide',
          zh: '阅读文档比较指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.32.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.32.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.31.0',
    date: '2026-08-25',
    kind: 'new',
    surfaces: ['writer'],
    title: {
      en: 'Native Writer indexes',
      zh: 'Writer 原生文档索引',
    },
    summary: {
      en: 'Writer can mark, edit, generate, customize, and refresh a bounded native document index with exact DOCX reopen behavior.',
      zh: 'Writer 现在可标记、编辑、生成、自定义并刷新有界的原生文档索引，并支持精确 DOCX 重开。',
    },
    highlights: [
      {
        title: { en: 'Typed index entries', zh: '类型化索引项' },
        detail: {
          en: 'Primary and secondary entries, cross-references, page emphasis, and stable marker identities stay selectable.',
          zh: '主次索引项、交叉引用、页码强调和稳定标记身份均保持可选择。',
        },
      },
      {
        title: { en: 'Responsive authoring', zh: '响应式创作' },
        detail: {
          en: 'Marking, insertion, customization, and refresh restore focus and create one transaction plus one Undo record.',
          zh: '标记、插入、自定义和刷新都会恢复焦点，并只创建一次事务和一条撤销记录。',
        },
      },
      {
        title: { en: 'Native field fidelity', zh: '原生域保真' },
        detail: {
          en: 'XE entries, INDEX fields, cached rows, columns, leaders, and page styles survive export, reopen, and second export.',
          zh: 'XE 索引项、INDEX 域、缓存行、分栏、前导符和页码样式可经受导出、重开与再次导出。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#native-document-index',
          zh: './components/document.html#原生文档索引',
        },
        label: { en: 'Read the index guide', zh: '阅读文档索引指南' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.31.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.31.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.30.0',
    date: '2026-08-24',
    kind: 'new',
    surfaces: ['writer', 'spreadsheet'],
    title: {
      en: 'Writer references and Spreadsheet validation',
      zh: 'Writer 引用工具与 Spreadsheet 数据验证',
    },
    summary: {
      en: 'Table of contents, character fidelity, and complete data-validation settings become first-class editable workflows with native round trips.',
      zh: '目录、字符保真和完整数据验证设置成为一等可编辑工作流，并支持原生往返。',
    },
    highlights: [
      {
        title: { en: 'Updatable table of contents', zh: '可更新目录' },
        detail: {
          en: 'Heading levels, links, page numbers, alignment, leaders, cached entries, and explicit refresh share one typed block.',
          zh: '标题级别、链接、页码、对齐、前导符、缓存条目和显式刷新共用同一类型化块。',
        },
      },
      {
        title: { en: 'Character fidelity', zh: '字符保真' },
        detail: {
          en: 'Native shading patterns and Latin, East Asian, or bidirectional proofing languages author and round-trip explicitly.',
          zh: '原生底纹图案以及拉丁、东亚、双向校对语言都可显式创作并往返。',
        },
      },
      {
        title: { en: 'Complete validation UX', zh: '完整验证体验' },
        detail: {
          en: 'Spreadsheet data validation exposes dropdowns, input prompts, error alerts, formulas, ranges, and native XLSX semantics.',
          zh: 'Spreadsheet 数据验证覆盖下拉列表、输入提示、错误警告、公式、区域与原生 XLSX 语义。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#native-table-of-contents',
          zh: './components/document.html#原生可更新目录',
        },
        label: {
          en: 'Read the Writer reference guide',
          zh: '阅读 Writer 引用指南',
        },
      },
      {
        href: {
          en: './components/spreadsheet.html#data-validation',
          zh: './components/spreadsheet.html#数据验证',
        },
        label: {
          en: 'Read the validation guide',
          zh: '阅读数据验证指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.30.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.30.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
] as const;

export function officeReleaseNotesThroughVersion(
  version: string,
): readonly OfficeReleaseNote[] {
  if (version === 'latest') return OFFICE_RELEASE_NOTES;
  const ceiling = parseSemanticVersion(version);
  if (!ceiling) return [];
  return OFFICE_RELEASE_NOTES.filter((release) => {
    const candidate = parseSemanticVersion(release.version);
    return (
      candidate !== null && compareSemanticVersions(candidate, ceiling) <= 0
    );
  });
}

type SemanticVersion = readonly [major: number, minor: number, patch: number];

function parseSemanticVersion(value: string): SemanticVersion | null {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(value);
  if (!match) return null;
  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

function compareSemanticVersions(
  left: SemanticVersion,
  right: SemanticVersion,
): number {
  for (let index = 0; index < left.length; index += 1) {
    const difference = (left[index] ?? 0) - (right[index] ?? 0);
    if (difference !== 0) return difference;
  }
  return 0;
}
