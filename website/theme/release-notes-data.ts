export const OFFICE_RELEASE_NOTES: readonly OfficeReleaseNote[] = [
  {
    version: '0.86.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes row wAfter revisions reviewable',
      zh: 'Writer 让行 wAfter 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free trPrChange priors with wAfter import as row-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 wAfter 的无关系 trPrChange 先验导入为 row-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Row width after',
          zh: '行后置宽度',
        },
        detail: {
          en: 'Relationship-free wAfter priors round-trip as reviewable row-formatting, including current-row export.',
          zh: '无关系的 wAfter 先验作为可审阅 row-formatting 往返，并覆盖当前行导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-row-width-after under track-changes creates a pending row-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-row-width-after 会生成待审阅的 row-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque row fixtures moved off wAfter to cnfStyle.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明行夹具已从 wAfter 迁至 cnfStyle。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.86.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.86.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.85.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes row wBefore revisions reviewable',
      zh: 'Writer 让行 wBefore 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free trPrChange priors with wBefore import as row-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 wBefore 的无关系 trPrChange 先验导入为 row-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Row width before',
          zh: '行前置宽度',
        },
        detail: {
          en: 'Relationship-free wBefore priors round-trip as reviewable row-formatting, including current-row export.',
          zh: '无关系的 wBefore 先验作为可审阅 row-formatting 往返，并覆盖当前行导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-row-width-before under track-changes creates a pending row-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-row-width-before 会生成待审阅的 row-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque row fixtures moved off wBefore to wAfter.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明行夹具已从 wBefore 迁至 wAfter。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.85.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.85.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.84.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes row gridAfter revisions reviewable',
      zh: 'Writer 让行 gridAfter 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free trPrChange priors with gridAfter import as row-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 gridAfter 的无关系 trPrChange 先验导入为 row-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Row grid after',
          zh: '行后置网格',
        },
        detail: {
          en: 'Relationship-free gridAfter priors round-trip as reviewable row-formatting, including current-row export.',
          zh: '无关系的 gridAfter 先验作为可审阅 row-formatting 往返，并覆盖当前行导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-row-grid-after under track-changes creates a pending row-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-row-grid-after 会生成待审阅的 row-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque row fixtures moved off gridAfter to wBefore.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明行夹具已从 gridAfter 迁至 wBefore。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.84.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.84.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.83.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes row gridBefore revisions reviewable',
      zh: 'Writer 让行 gridBefore 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free trPrChange priors with gridBefore import as row-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 gridBefore 的无关系 trPrChange 先验导入为 row-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Row grid before',
          zh: '行前置网格',
        },
        detail: {
          en: 'Relationship-free gridBefore priors round-trip as reviewable row-formatting, including current-row export.',
          zh: '无关系的 gridBefore 先验作为可审阅 row-formatting 往返，并覆盖当前行导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-row-grid-before under track-changes creates a pending row-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-row-grid-before 会生成待审阅的 row-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque row fixtures moved off gridBefore to gridAfter.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明行夹具已从 gridBefore 迁至 gridAfter。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.83.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.83.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.82.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes cell hideMark revisions reviewable',
      zh: 'Writer 让单元格 hideMark 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free tcPrChange priors with hideMark import as cell-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 hideMark 的无关系 tcPrChange 先验导入为 cell-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Cell hide mark',
          zh: '单元格隐藏标记',
        },
        detail: {
          en: 'Relationship-free hideMark priors round-trip as reviewable cell-formatting, including current-cell export.',
          zh: '无关系的 hideMark 先验作为可审阅 cell-formatting 往返，并覆盖当前单元格导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-cell-hide-mark under track-changes creates a pending cell-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-cell-hide-mark 会生成待审阅的 cell-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque cell fixtures moved off hideMark to cnfStyle.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明单元格夹具已从 hideMark 迁至 cnfStyle。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.82.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.82.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.81.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes cell tcFitText revisions reviewable',
      zh: 'Writer 让单元格 tcFitText 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free tcPrChange priors with tcFitText import as cell-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 tcFitText 的无关系 tcPrChange 先验导入为 cell-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Cell fit text',
          zh: '单元格适应文字',
        },
        detail: {
          en: 'Relationship-free tcFitText priors round-trip as reviewable cell-formatting, including current-cell export.',
          zh: '无关系的 tcFitText 先验作为可审阅 cell-formatting 往返，并覆盖当前单元格导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-cell-fit-text under track-changes creates a pending cell-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-cell-fit-text 会生成待审阅的 cell-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque cell fixtures moved off tcFitText to hideMark.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明单元格夹具已从 tcFitText 迁至 hideMark。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.81.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.81.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.80.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer makes cell textDirection revisions reviewable',
      zh: 'Writer 让单元格 textDirection 修订可审阅',
    },
    summary: {
      en: 'Phase 0 fidelity: relationship-free tcPrChange priors with textDirection import as cell-formatting with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：含 textDirection 的无关系 tcPrChange 先验导入为 cell-formatting，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Cell text direction',
          zh: '单元格文字方向',
        },
        detail: {
          en: 'Supported ST_TextDirection values (lrTb, tbRl, btLr, lrTbV, tbRlV, tbLrV) round-trip as reviewable cell-formatting, including current-cell export.',
          zh: '支持的 ST_TextDirection 值（lrTb、tbRl、btLr、lrTbV、tbRlV、tbLrV）作为可审阅 cell-formatting 往返，并覆盖当前单元格导出。',
        },
      },
      {
        title: {
          en: 'Live track-changes',
          zh: '实时修订跟踪',
        },
        detail: {
          en: 'Editing data-office-cell-text-direction under track-changes creates a pending cell-formatting revision that accept/reject can restore.',
          zh: '在修订跟踪开启时编辑 data-office-cell-text-direction 会生成待审阅的 cell-formatting 修订，接受/拒绝可还原。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; opaque cell fixtures moved off textDirection to tcFitText.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；不透明单元格夹具已从 textDirection 迁至 tcFitText。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.80.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.80.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './native-office-engine.html', zh: './native-office-engine.html' },
        label: { en: 'Native engine notes', zh: '原生引擎说明' },
      },
    ],
  },
  {
    version: '0.79.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer deepens reviewable cell and row property revisions',
      zh: 'Writer 深化可审阅的单元格与行属性修订',
    },
    summary: {
      en: 'Phase 0 fidelity: reviewable tcW and noWrap cell-formatting plus hidden and jc row-formatting, with accept/reject, live track-changes, and native DOCX export.',
      zh: 'Phase 0 保真：可审阅的 tcW / noWrap 单元格格式修订，以及 hidden / jc 行格式修订，支持接受/拒绝、实时修订跟踪与原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Cell width and wrap',
          zh: '单元格宽度与换行',
        },
        detail: {
          en: 'Relationship-free tcPrChange priors with tcW (auto/dxa/pct) or noWrap import as cell-formatting; live colwidth and noWrap edits track and export natively.',
          zh: '含 tcW（auto/dxa/pct）或 noWrap 的无关系 tcPrChange 先验导入为 cell-formatting；实时 colwidth 与 noWrap 编辑可跟踪并原生导出。',
        },
      },
      {
        title: {
          en: 'Row hide and justify',
          zh: '行隐藏与对齐',
        },
        detail: {
          en: 'trPrChange priors with hidden or jc (left/center/right) are reviewable as row-formatting, including current-row export of those properties.',
          zh: '含 hidden 或 jc（left/center/right）的 trPrChange 先验可作为 row-formatting 审阅，并覆盖当前行属性的导出。',
        },
      },
      {
        title: {
          en: 'Opaque keepers stay fail-closed',
          zh: '不透明路径保持失败闭合',
        },
        detail: {
          en: 'Broader cell/row property sets remain opaque metadata or fail-closed; fixtures moved off admitted priors to textDirection and gridBefore.',
          zh: '更广的单元格/行属性集仍为不透明元数据或失败闭合；夹具已从已接纳先验迁至 textDirection 与 gridBefore。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.79.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.79.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './native-office-engine.html', zh: './native-office-engine.html' },
        label: { en: 'Native engine notes', zh: '原生引擎说明' },
      },
    ],
  },
  {
    version: '0.78.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: ['writer', 'pdf', 'playground', 'documentation'],
    title: {
      en: 'Writer no-clobber revisions deepen with searchable PDF vectors',
      zh: 'Writer 无破坏修订深化，PDF 导出补齐可搜索矢量文本',
    },
    summary: {
      en: 'Phase 0 fidelity: reviewable table/row/cell/section property revisions, richer paragraph-mark and move companions, opaque float/OMML preservation, and Latin (optional CJK) searchable PDF vector text with a tagged structure bootstrap.',
      zh: 'Phase 0 保真：可审阅的表格/行/单元格/节属性修订、更丰富的段落标记与移动伴随标记、不透明浮动/公式保留，以及 Latin（可选 CJK）可搜索 PDF 矢量文本与结构引导。',
    },
    highlights: [
      {
        title: {
          en: 'Reviewable property revisions',
          zh: '可审阅的属性修订',
        },
        detail: {
          en: 'tblPrChange, trPrChange, tcPrChange, and sectPrChange subsets (including layout, cell margins, paper source, and equal-width columns) accept/reject with live track-changes.',
          zh: 'tblPrChange、trPrChange、tcPrChange 与 sectPrChange 子集（含布局、单元格边距、纸张来源与等宽分栏）支持接受/拒绝，并覆盖实时修订跟踪。',
        },
      },
      {
        title: {
          en: 'Paragraph-mark and move fidelity',
          zh: '段落标记与移动保真',
        },
        detail: {
          en: 'Multi-wrapper mark bodies, soft breaks, internal hyperlinks, bookmarks, empty rPr siblings, move-range companions, and eligible paragraph-break merge/split stay reviewable.',
          zh: '多包装标记正文、软换行、内部超链接、书签、空 rPr 兄弟、移动范围伴随标记，以及合格的段落断裂合并/拆分保持可审阅。',
        },
      },
      {
        title: {
          en: 'Searchable PDF vectors',
          zh: '可搜索 PDF 矢量文本',
        },
        detail: {
          en: 'Measured Latin clears to Helvetica vectors; hosts may register a CJK TrueType face; title/lang/heading outline bootstrap ships without claiming full PDF/UA.',
          zh: '已测量 Latin 清除为 Helvetica 矢量；宿主可注册 CJK TrueType；标题/语言/标题大纲引导已交付，但不宣称完整 PDF/UA。',
        },
      },
    ],
    links: [
      {
        href: { en: './components/document.html', zh: './components/document.html' },
        label: { en: 'Document editor', zh: '文档编辑器' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.78.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.78.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './native-office-engine.html', zh: './native-office-engine.html' },
        label: { en: 'Native engine notes', zh: '原生引擎说明' },
      },
    ],
  },
  {
    version: '0.77.0',
    date: '2026-09-08',
    kind: 'improved',
    surfaces: [
      'writer',
      'spreadsheet',
      'presentation',
      'markdown',
      'pdf',
      'playground',
      'documentation',
    ],
    title: {
      en: 'Five editors share one WPS shortcut matrix (103 ACL / 78 visual)',
      zh: '五编辑器共享同一 WPS 快捷键矩阵（103 ACL / 78 视觉）',
    },
    summary: {
      en: 'One release story across Writer, Spreadsheet, Presentation, Markdown, and PDF: WPS-aligned chords with ribbon ads plus dedicated ACL and responsive visual evidence on the shared operator matrix.',
      zh: 'Writer、Spreadsheet、Presentation、Markdown 与 PDF 共用同一发布叙事：对齐 WPS 的快捷键、功能区广告，以及共享操作器矩阵上的专用 ACL 与响应式视觉证据。',
    },
    highlights: [
      {
        title: {
          en: 'Writer paragraph and field chords',
          zh: 'Writer 段落与域快捷键',
        },
        detail: {
          en: 'Paragraph move, distribute align, bullets, indent families, double strikethrough, Normal style, case cycle, clear formatting, and field lock/unlink join the matrix.',
          zh: '段落移动、分散对齐、项目符号、缩进族、双删除线、正文样式、大小写循环、清除格式以及域锁定/取消链接进入矩阵。',
        },
      },
      {
        title: {
          en: 'Presentation and Spreadsheet chrome',
          zh: 'Presentation 与 Spreadsheet 界面',
        },
        detail: {
          en: 'Slide new/duplicate/delete, notes toggle, group ads, ribbon collapse, workbook View toggles, and sheet PageUp/PageDown navigation are covered.',
          zh: '覆盖新建/复制/删除幻灯片、备注开关、组合广告、功能区折叠、工作簿视图开关以及工作表 PageUp/PageDown 导航。',
        },
      },
      {
        title: {
          en: 'Markdown and PDF keyboard proof',
          zh: 'Markdown 与 PDF 键盘证据',
        },
        detail: {
          en: 'Markdown bold/italic/link and PDF zoom, page navigation, and Ctrl+F search focus ship with ACL plus desktop/compact visuals.',
          zh: 'Markdown 加粗/斜体/链接以及 PDF 缩放、翻页和 Ctrl+F 搜索焦点均配有 ACL 与桌面/紧凑视觉证据。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.77.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.77.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.76.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer WPS numeric fields gain responsive visual evidence',
      zh: 'Writer WPS 数字字段补齐响应式视觉证据',
    },
    summary: {
      en: 'The shared operator now inventories 61 ACL and 56 visual contracts, with focused Writer field results, instructions, labels, and focus evidence.',
      zh: '共享操作器现在公开 61 个 ACL 和 56 个视觉契约，深入覆盖 Writer 字段结果、指令、标签与焦点证据。',
    },
    highlights: [
      {
        title: { en: 'Live WPS field results', zh: '实时 WPS 字段结果' },
        detail: {
          en: 'PAGE, NUMPAGES, SECTION, and PAGEREF resolve to the expected Roman, alphabetic, and ordinal displays on the measured page.',
          zh: 'PAGE、NUMPAGES、SECTION 和 PAGEREF 在测量页面上解析为预期的罗马、字母和序数显示。',
        },
      },
      {
        title: {
          en: 'Native instructions and labels',
          zh: '原生指令与可访问标签',
        },
        detail: {
          en: 'MERGEFORMAT instructions, bookmark targets, and localized ARIA labels remain inspectable in the Writer DOM projection.',
          zh: 'MERGEFORMAT 指令、书签目标和本地化 ARIA 标签在 Writer DOM 投影中保持可检查。',
        },
      },
      {
        title: { en: 'Responsive focus proof', zh: '响应式焦点证据' },
        detail: {
          en: 'Desktop and compact layouts keep the one-page projection stable, and F9 refresh returns focus to the document editor without browser errors.',
          zh: '桌面与紧凑布局保持单页投影稳定，F9 刷新后焦点回到文档编辑器，且无浏览器错误。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.76.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.76.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.75.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer WPS layout and font-grid parity gain visual evidence',
      zh: 'Writer WPS 版式与字体网格对齐补齐视觉证据',
    },
    summary: {
      en: 'The shared operator now inventories 61 ACL and 55 visual contracts, with focused Writer pagination, font, grid, and mixed-script evidence.',
      zh: '共享操作器现在公开 61 个 ACL 和 55 个视觉契约，深入覆盖 Writer 分页、字体、网格与混合脚本证据。',
    },
    highlights: [
      {
        title: { en: 'WPS page layout', zh: 'WPS 页面版式' },
        detail: {
          en: 'Imported A4 portrait metadata, automatic line metrics, table flow, and one-page pagination are checked in desktop and compact layouts.',
          zh: '桌面与紧凑布局检查导入的 A4 纵向元数据、自动行距指标、表格流和单页分页。',
        },
      },
      {
        title: { en: 'Font and document grid', zh: '字体与文档网格' },
        detail: {
          en: 'Latin/CJK line-height factors and the WPS document-grid line pitch remain inspectable through the Writer DOM projection.',
          zh: '拉丁/CJK 行高因子和 WPS 文档网格行距可通过 Writer DOM 投影检查。',
        },
      },
      {
        title: { en: 'Mixed-script fidelity', zh: '混合脚本保真度' },
        detail: {
          en: 'RTL paragraphs, Arabic/Hebrew/CJK/Latin font slots, and bold/italic mixed runs keep their WPS-aligned metrics without browser errors.',
          zh: 'RTL 段落、阿拉伯语/希伯来语/CJK/拉丁字体槽以及粗体/斜体混合文本保留 WPS 对齐指标，且无浏览器错误。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.75.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.75.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.74.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer WPS shortcuts gain desktop and compact visual evidence',
      zh: 'Writer WPS 快捷键补齐桌面与紧凑视觉证据',
    },
    summary: {
      en: 'The shared operator now inventories 61 ACL and 54 visual contracts, with focused Writer formatting, review, menu, and focus evidence.',
      zh: '共享操作器现在公开 61 个 ACL 和 54 个视觉契约，深入覆盖 Writer 格式、审阅、菜单与焦点证据。',
    },
    highlights: [
      {
        title: { en: 'WPS formatting shortcuts', zh: 'WPS 格式快捷键' },
        detail: {
          en: 'Format copy/paste, mutually exclusive text-case effects, one-step undo, paragraph alignment/spacing, and heading styles are covered in both layouts.',
          zh: '两种布局均覆盖格式复制/粘贴、互斥大小写效果、单步撤销、段落对齐/行距和标题样式。',
        },
      },
      {
        title: { en: 'Accessible command discovery', zh: '可访问的命令发现' },
        detail: {
          en: 'Text-case menu items expose their shortcuts, Escape restores the menu trigger, and spelling/review shortcuts preserve editor focus.',
          zh: '大小写菜单项公开快捷键，Escape 恢复菜单触发器焦点，拼写/审阅快捷键保留正文焦点。',
        },
      },
      {
        title: {
          en: 'Local WPS and browser proof',
          zh: '本机 WPS 与浏览器证据',
        },
        detail: {
          en: 'The WPS Writer COM probe ran against build 12.1.0.21541, and A3S Test 1.0.1 passed the shortcut ACL with empty console and page-error evidence.',
          zh: 'WPS Writer COM 探针在 12.1.0.21541 版本运行；A3S Test 1.0.1 的快捷键 ACL 通过，控制台和页面错误证据均为空。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.74.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.74.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.73.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet copy-from-above shortcuts gain WPS evidence',
      zh: 'Spreadsheet 从上方复制快捷键补齐 WPS 证据',
    },
    summary: {
      en: 'The shared operator now inventories 61 ACL and 53 visual contracts, with focused copy-from-above, style, undo, and focus evidence.',
      zh: '共享操作器现在公开 61 个 ACL 和 53 个视觉契约，深入覆盖从上方复制、样式、撤销与焦点证据。',
    },
    highlights: [
      {
        title: { en: 'WPS copy shortcuts', zh: 'WPS 复制快捷键' },
        detail: {
          en: "Ctrl+' copies the source formula, while Ctrl+Shift+' copies its calculated value from the active grid.",
          zh: "Ctrl+' 从活动网格复制源公式，Ctrl+Shift+' 复制其计算值。",
        },
      },
      {
        title: { en: 'Style and undo parity', zh: '样式与撤销对齐' },
        detail: {
          en: 'Copy-from-above preserves the target style, supports one-step undo, and restores grid focus after each action.',
          zh: '从上方复制保留目标样式，支持单步撤销，并在每次操作后恢复网格焦点。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'The Commander-based visual gate passed desktop and compact layouts, and A3S Test 1.0.1 passed the ACL with empty console and page-error evidence.',
          zh: '基于 Commander 的视觉门禁通过桌面与紧凑布局；A3S Test 1.0.1 的 ACL 场景通过，控制台和页面错误证据均为空。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.73.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.73.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.72.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet WPS font aliases and color resets gain visual evidence',
      zh: 'Spreadsheet WPS 字体快捷键与颜色复位补齐视觉证据',
    },
    summary: {
      en: 'The shared operator now inventories 60 ACL and 52 visual contracts, with focused font-alias, color-reset, and focus-restoration evidence.',
      zh: '共享操作器现在公开 60 个 ACL 和 52 个视觉契约，深入覆盖字体快捷键、颜色复位与焦点恢复证据。',
    },
    highlights: [
      {
        title: { en: 'WPS font aliases', zh: 'WPS 字体快捷键' },
        detail: {
          en: 'Ctrl+2, Ctrl+3, and Ctrl+4 remain discoverable through shared aria-keyshortcuts metadata and apply from the active grid.',
          zh: 'Ctrl+2、Ctrl+3 和 Ctrl+4 通过统一的 aria-keyshortcuts 元数据保持可发现，并从活动网格执行。',
        },
      },
      {
        title: { en: 'Direct color resets', zh: '直接颜色复位' },
        detail: {
          en: 'Automatic Color and No Fill are explicit actions that clear direct styles, preserve selection, and return focus to the grid.',
          zh: '自动颜色与无填充是显式操作，可清除直接样式、保留选区并将焦点返回网格。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'The Commander-based visual gate passed desktop and compact layouts, and A3S Test 1.0.1 passed both ACL scenarios with empty console and page-error evidence.',
          zh: '基于 Commander 的视觉门禁通过桌面与紧凑布局；A3S Test 1.0.1 的两个 ACL 场景全部通过，控制台和页面错误证据均为空。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.72.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.72.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.71.0',
    date: '2026-09-07',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet hyperlink safety and underline parity become explicit',
      zh: 'Spreadsheet 超链接安全与下划线对齐成为显式证据',
    },
    summary: {
      en: 'The shared operator now inventories 60 ACL and 51 visual contracts, with focused hyperlink, underline, and shell-asset evidence.',
      zh: '共享操作器现在公开 60 个 ACL 和 51 个视觉契约，深入覆盖超链接、下划线与外壳资源证据。',
    },
    highlights: [
      {
        title: { en: 'Hyperlink safety', zh: '超链接安全' },
        detail: {
          en: 'The dialog rejects javascript URLs, protects hidden worksheets, normalizes HTTPS addresses, and returns focus to the grid after removal.',
          zh: '对话框拒绝 javascript URL、保护隐藏工作表、规范化 HTTPS 地址，并在移除后将焦点返回网格。',
        },
      },
      {
        title: { en: 'Underline parity', zh: '下划线对齐' },
        detail: {
          en: 'Advanced underline styles stay aligned across the ribbon menu, cell-format dialog, keyboard shortcuts, and undo history.',
          zh: '高级下划线样式在功能区菜单、单元格格式对话框、键盘快捷键和撤销历史之间保持一致。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Four Spreadsheet visual cases passed at desktop and compact layouts, plus a live A3S Test 1.0.1 hyperlink session with no page errors or console messages; shared shell assets remove the logo 404.',
          zh: '4 个 Spreadsheet 视觉用例在桌面与紧凑布局通过，并完成无页面错误或控制台消息的 A3S Test 1.0.1 实时超链接会话；共享外壳资源消除了 logo 404。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.71.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.71.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.70.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet phone Find and worksheet naming join the WPS evidence matrix',
      zh: 'Spreadsheet 手机端查找与工作表命名进入 WPS 证据矩阵',
    },
    summary: {
      en: 'The shared operator now inventories 58 ACL and 49 visual contracts, with focused phone Find and worksheet-naming evidence.',
      zh: '共享操作器现在公开 58 个 ACL 和 49 个视觉契约，深入覆盖手机端查找与工作表命名证据。',
    },
    highlights: [
      {
        title: { en: 'Phone Find', zh: '手机端查找' },
        detail: {
          en: 'Touch-sized Find controls report the match count, select the matching cell, and restore grid focus on Escape.',
          zh: '触控尺寸的查找控件显示匹配数量、选中匹配单元格，并在 Escape 后恢复网格焦点。',
        },
      },
      {
        title: { en: 'Worksheet naming', zh: '工作表命名' },
        detail: {
          en: 'Invalid worksheet names provide inline feedback; cancellation is safe and the options-menu invoker regains focus after dismissal.',
          zh: '无效工作表名称提供行内反馈；取消操作安全，关闭菜单后选项触发器恢复焦点。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Four Spreadsheet phone visual cases passed at desktop and compact layouts, plus a live A3S Test 1.0.1 Find session with no page errors or console messages.',
          zh: '4 个 Spreadsheet 手机端视觉用例在桌面与紧凑布局通过，并完成无页面错误或控制台消息的 A3S Test 1.0.1 实时查找会话。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.70.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.70.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.69.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet phone workflows join the WPS evidence matrix',
      zh: 'Spreadsheet 手机端流程进入 WPS 证据矩阵',
    },
    summary: {
      en: 'The shared operator now inventories 56 ACL and 47 visual contracts, with focused phone task-pane and context-menu evidence.',
      zh: '共享操作器现在公开 56 个 ACL 和 47 个视觉契约，深入覆盖手机端任务面板与上下文菜单证据。',
    },
    highlights: [
      {
        title: { en: 'Phone task pane', zh: '手机端任务面板' },
        detail: {
          en: 'The data-pivot pane keeps focus inside the modal, supports forward and reverse keyboard traversal, and restores the ribbon invoker on Escape.',
          zh: '数据透视表面板保持模态焦点隔离，支持正向与反向键盘遍历，并在 Escape 后恢复功能区触发器焦点。',
        },
      },
      {
        title: { en: 'Phone context menu', zh: '手机端上下文菜单' },
        detail: {
          en: 'The touch-sized selection menu exposes a focused first action and returns focus to the Spreadsheet grid after dismissal.',
          zh: '触控尺寸的选区菜单聚焦首个操作，并在关闭后将焦点恢复到 Spreadsheet 网格。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Four Spreadsheet phone visual cases passed at desktop and compact layouts, plus a live A3S Test 1.0.1 task-pane session with no page errors or console messages.',
          zh: '4 个 Spreadsheet 手机端视觉用例在桌面与紧凑布局通过，并完成无页面错误或控制台消息的 A3S Test 1.0.1 实时任务面板会话。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.69.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.69.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.68.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['presentation', 'playground', 'documentation'],
    title: {
      en: 'Presentation phone review joins the WPS evidence matrix',
      zh: 'Presentation 手机端审阅进入 WPS 证据矩阵',
    },
    summary: {
      en: 'The shared operator now inventories 54 ACL and 45 visual contracts, with focused phone chart and comments evidence.',
      zh: '共享操作器现在公开 54 个 ACL 和 45 个视觉契约，深入覆盖手机端图表与批注证据。',
    },
    highlights: [
      {
        title: { en: 'Phone chart pane', zh: '手机端图表面板' },
        detail: {
          en: 'Chart editing keeps modal focus contained, rolls back invalid drafts, and restores the selected chart trigger on close.',
          zh: '图表编辑保持模态焦点隔离，回滚无效草稿，并在关闭后恢复选中图表触发器焦点。',
        },
      },
      {
        title: { en: 'Phone comments review', zh: '手机端批注审阅' },
        detail: {
          en: 'Comments review keeps the phone editor inert, cancels drafts cleanly, and restores the original invoker after closing.',
          zh: '批注审阅让手机编辑器保持模态隔离，干净取消草稿，并在关闭后恢复原始触发器焦点。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Four Presentation phone visual cases passed at desktop and compact layouts, plus a live A3S Test 1.0.1 review with no page errors.',
          zh: '4 个 Presentation 手机端视觉用例在桌面与紧凑布局通过，并完成无页面错误的 A3S Test 1.0.1 实时审阅。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.68.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.68.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.67.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet appearance editing joins the WPS evidence matrix',
      zh: 'Spreadsheet 外观编辑进入 WPS 证据矩阵',
    },
    summary: {
      en: 'The shared operator now inventories 52 ACL and 43 visual contracts, with deep border and fill evidence.',
      zh: '共享操作器现在公开 52 个 ACL 和 43 个视觉契约，深入覆盖边框与填充证据。',
    },
    highlights: [
      {
        title: { en: 'Native border controls', zh: '原生边框控件' },
        detail: {
          en: 'Diagonal-up and diagonal-down borders remain editable, reversible, and responsive through Format Cells.',
          zh: '斜上与斜下边框可通过“设置单元格格式”编辑、撤销并适配紧凑布局。',
        },
      },
      {
        title: { en: 'Gradient and pattern fills', zh: '渐变与图案填充' },
        detail: {
          en: 'Linear/path gradients and seventeen native pattern styles are explicit contracts with editable dialogs and Undo.',
          zh: '线性/路径渐变与 17 种原生图案样式成为显式契约，支持可编辑对话框与撤销。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Ten Spreadsheet visual cases passed at desktop and compact layouts, plus a bounded live path-gradient edit with no page errors.',
          zh: '10 个 Spreadsheet 视觉用例在桌面与紧凑布局通过，并完成无页面错误的有界路径渐变实时编辑。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.67.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.67.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.66.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['spreadsheet', 'playground', 'documentation'],
    title: {
      en: 'Spreadsheet sorting becomes a first-class WPS matrix slice',
      zh: 'Spreadsheet 排序进入 WPS 一等证据矩阵',
    },
    summary: {
      en: 'The shared operator now inventories 49 ACL and 40 visual contracts, with deep sort and AutoFilter evidence.',
      zh: '共享操作器现在公开 49 个 ACL 和 40 个视觉契约，深入覆盖排序与自动筛选证据。',
    },
    highlights: [
      {
        title: { en: 'Sort semantics', zh: '排序语义' },
        detail: {
          en: 'Appearance/color, custom-list, table-owned, left-to-right, partial-range, and Simplified Chinese text sorting are now explicit contracts.',
          zh: '外观/颜色、自定义序列、表格所有者、左右、部分范围和简体中文文本排序现在都是显式契约。',
        },
      },
      {
        title: { en: 'Filter-safe ranges', zh: '筛选安全范围' },
        detail: {
          en: 'Owned-range sorting re-applies AutoFilter criteria and preserves Undo/focus behavior across responsive layouts.',
          zh: '所有者范围排序会重新应用自动筛选条件，并在响应式布局中保持撤销/焦点行为。',
        },
      },
      {
        title: { en: 'Browser proof', zh: '浏览器证据' },
        detail: {
          en: 'Fourteen Spreadsheet visual cases passed at desktop and compact layouts, plus a bounded live A3S Test 1.0.1 sort session.',
          zh: '14 个 Spreadsheet 视觉用例在桌面与紧凑布局通过，并完成有界 A3S Test 1.0.1 实时排序会话。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.66.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.66.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.65.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer typography and review paths join the WPS matrix',
      zh: 'Writer 排版与审阅路径进入 WPS 矩阵',
    },
    summary: {
      en: 'The shared operator now exposes 43 ACL and 34 visual contracts, with deep Writer font, hidden-text, content-control, and review evidence.',
      zh: '共享操作器现在公开 43 个 ACL 和 34 个视觉契约，深入覆盖 Writer 字体、隐藏文字、内容控件与审阅证据。',
    },
    highlights: [
      {
        title: { en: 'WPS typography paths', zh: 'WPS 排版路径' },
        detail: {
          en: 'Character position, scale, spacing, emphasis, hidden text, and OpenType controls are now first-class Writer matrix contracts.',
          zh: '字符位置、缩放、间距、着重号、隐藏文字和 OpenType 控件现在都是 Writer 矩阵的一等契约。',
        },
      },
      {
        title: { en: 'Review-safe authoring', zh: '安全审阅编辑' },
        detail: {
          en: 'Controlled review conflicts, content controls, paired move revisions, and phone track-changes paths are checked beside the existing WPS suites.',
          zh: '受控审阅冲突、内容控件、成对移动修订和手机端修订路径与既有 WPS 套件一起接受检查。',
        },
      },
      {
        title: { en: 'Live focus evidence', zh: '实时焦点证据' },
        detail: {
          en: 'Forty-one Writer visual cases passed across desktop and compact layouts, with one platform-conditional case skipped; a live typed font-dialog edit passed without page errors.',
          zh: 'Writer 桌面与紧凑布局共 41 个视觉用例通过，另有 1 个平台条件用例跳过；实时字体对话框类型化编辑通过且无页面错误。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.65.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.65.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.64.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: [
      'spreadsheet',
      'presentation',
      'pdf',
      'playground',
      'documentation',
    ],
    title: {
      en: 'The WPS evidence matrix gets deep Spreadsheet coverage',
      zh: 'WPS 证据矩阵深入覆盖 Spreadsheet 等编辑器',
    },
    summary: {
      en: 'A3S Test now inventories 33 ACL and 24 visual contracts across the five editors, with live Spreadsheet focus evidence.',
      zh: 'A3S Test 现在为五种编辑器清晰列出 33 个 ACL 和 24 个视觉契约，并提供实时 Spreadsheet 焦点证据。',
    },
    highlights: [
      {
        title: {
          en: 'Deeper Spreadsheet paths',
          zh: '更深入的 Spreadsheet 路径',
        },
        detail: {
          en: 'AutoSum, Paste Special, conditional formatting, date/time, cell styles, rich text, table totals, orientation/visibility, and direct-color shortcuts are now matrix contracts.',
          zh: '自动求和、选择性粘贴、条件格式、日期时间、单元格样式、富文本、表格汇总、方向/可见性和直接颜色快捷键现在都进入矩阵契约。',
        },
      },
      {
        title: { en: 'Cross-surface regression', zh: '跨编辑器回归' },
        detail: {
          en: 'Presentation IME/large-window and PDF large-file ACLs join the Writer, Markdown, and existing Spreadsheet workflows.',
          zh: 'Presentation IME/大窗口与 PDF 大文件 ACL 加入 Writer、Markdown 以及既有 Spreadsheet 流程。',
        },
      },
      {
        title: { en: 'Real browser evidence', zh: '真实浏览器证据' },
        detail: {
          en: 'Sixteen Spreadsheet visual tests passed at desktop and compact layouts; a bounded A3S Test 1.0.1 session restored grid selection after a typed edit.',
          zh: '16 个 Spreadsheet 视觉测试在桌面和紧凑布局通过；有界 A3S Test 1.0.1 会话在类型化编辑后恢复了网格选区。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.64.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.64.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.63.1',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: [
      'writer',
      'spreadsheet',
      'presentation',
      'pdf',
      'playground',
      'documentation',
    ],
    title: {
      en: 'Windows A3S Test dispatch stays native and typed',
      zh: 'Windows A3S Test 派发保持原生且类型化',
    },
    summary: {
      en: 'The WPS-informed operator now runs the five-surface evidence matrix with a native Windows CDP adapter and a fail-closed capability check.',
      zh: '基于 WPS 参考的操作器现在使用原生 Windows CDP 适配器和 fail-closed 能力检查运行五种编辑器证据矩阵。',
    },
    highlights: [
      {
        title: { en: 'Cross-surface coverage', zh: '跨编辑器覆盖' },
        detail: {
          en: 'Supplying `--cdp-port` compiles the native `.exe` adapter under `.a3s-test/office-ops`, keeping selectors and action JSON out of `.cmd` parsing.',
          zh: '提供 `--cdp-port` 后会编译 `.a3s-test/office-ops` 下的原生 `.exe` 适配器，避免选择器和 action JSON 进入 `.cmd` 解析。',
        },
      },
      {
        title: { en: 'Typed Windows dispatch', zh: 'Windows 类型化派发' },
        detail: {
          en: '`doctor --json` only reports CUA certification after a compatible A3S Test capability probe succeeds; stale 0.x binaries remain fail-closed.',
          zh: '`doctor --json` 只有在兼容的 A3S Test 能力探测通过后才报告 CUA certification；过时的 0.x 二进制仍会 fail-closed。',
        },
      },
      {
        title: {
          en: 'Fail-closed browser gates',
          zh: '浏览器门禁 fail-closed',
        },
        detail: {
          en: 'A live Writer session completed the bounded `start → observe → one action → observe → finish` lifecycle through the Office CLI.',
          zh: '通过 Office CLI 完成了有界的 `start → observe → 一次 action → observe → finish` Writer 实时会话。',
        },
      },
    ],
    links: [
      {
        href: { en: './automation/index.html', zh: './automation/index.html' },
        label: { en: 'Run the five-surface matrix', zh: '运行五种编辑器矩阵' },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.63.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.63.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: { en: './cli-reference.html', zh: './cli-reference.html' },
        label: { en: 'CLI reference', zh: 'CLI 参考' },
      },
    ],
  },
  {
    version: '0.62.0',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer WPS evidence becomes one complete matrix',
      zh: 'Writer WPS 证据统一进入完整矩阵',
    },
    summary: {
      en: 'Existing WPS shortcut, layout-parity, and font/grid suites now flow through the same declarative operator as every other editor surface.',
      zh: '已有的 WPS 快捷键、版式对齐和字体/网格套件现在与其他编辑器共享同一个声明式操作矩阵。',
    },
    highlights: [
      {
        title: {
          en: 'No hidden Writer contracts',
          zh: 'Writer 契约不再隐藏',
        },
        detail: {
          en: '`check all`, `gate writer`, and `plan writer --json` now include shortcut, layout, and font/grid ACLs that were previously only reachable through standalone scripts.',
          zh: '`check all`、`gate writer` 和 `plan writer --json` 现在包含此前只能通过独立脚本运行的快捷键、版式和字体/网格 ACL。',
        },
      },
      {
        title: {
          en: 'Fixture inventory is explicit',
          zh: '夹具清单显式化',
        },
        detail: {
          en: 'The Writer plan reports every WPS layout and font-matrix DOCX fixture before an agent chooses an interaction path.',
          zh: 'Writer 清单会在智能体选择交互路径前报告全部 WPS 版式和字体矩阵 DOCX 夹具。',
        },
      },
      {
        title: {
          en: 'Metadata travels with the plan',
          zh: '计划携带完整元数据',
        },
        detail: {
          en: 'Machine-readable plans now carry visual, ACL, fixture, and agent metadata while retaining the observe → one action → observe boundary.',
          zh: '机器可读计划现在携带视觉、ACL、夹具和智能体元数据，同时保留 observe → 一次动作 → observe 边界。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the Writer matrix',
          zh: '运行 Writer 矩阵',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.62.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.62.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: {
          en: './cli-reference.html',
          zh: './cli-reference.html',
        },
        label: {
          en: 'CLI reference',
          zh: 'CLI 参考',
        },
      },
    ],
  },
  {
    version: '0.61.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: [
      'documentation',
      'playground',
      'writer',
      'spreadsheet',
      'presentation',
      'markdown',
      'pdf',
    ],
    title: {
      en: 'One typed plan for every editor surface',
      zh: '五种编辑器共享一份类型化工作流清单',
    },
    summary: {
      en: 'The Commander operator now emits a deterministic JSON workflow manifest so Codex can run UI/UX evidence without ad-hoc per-surface shell logic.',
      zh: 'Commander 操作 CLI 现在可以输出确定性的 JSON 工作流清单，让 Codex 无需临时拼接各编辑器的 shell 条件即可执行 UI/UX 证据流程。',
    },
    highlights: [
      {
        title: {
          en: 'Matrix-driven commands',
          zh: '矩阵驱动命令',
        },
        detail: {
          en: '`plan all --json` expands fixtures, ACL checks, A3S Test gates, visual projects, and bounded agent sessions for Writer, Spreadsheet, Presentation, Markdown, and PDF.',
          zh: '`plan all --json` 为 Writer、Spreadsheet、Presentation、Markdown 和 PDF 展开夹具、ACL、A3S Test 门禁、视觉项目和有界智能体会话。',
        },
      },
      {
        title: {
          en: 'Desktop and compact evidence',
          zh: '桌面与紧凑视口证据',
        },
        detail: {
          en: 'Each visual command carries both desktop-1280 and compact-768 projects, while the primary interaction authority remains A3S Test.',
          zh: '每个视觉命令同时携带 desktop-1280 与 compact-768 项目，主交互权威仍然是 A3S Test。',
        },
      },
      {
        title: {
          en: 'WPS reference stays bounded',
          zh: 'WPS 参考保持有界',
        },
        detail: {
          en: 'Writer plans include the isolated WPS COM UI and field probes; Windows CUA remains fail-closed when the locked driver is unsupported.',
          zh: 'Writer 清单包含隔离的 WPS COM UI 与字段探针；锁定驱动不支持时，Windows CUA 仍保持 fail-closed。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the editor workflow plan',
          zh: '运行编辑器工作流清单',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.61.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.61.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
      {
        href: {
          en: './cli-reference.html',
          zh: './cli-reference.html',
        },
        label: {
          en: 'CLI reference',
          zh: 'CLI 参考',
        },
      },
    ],
  },
  {
    version: '0.60.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'WPS UI references become typed evidence',
      zh: 'WPS UI 参考升级为类型化证据',
    },
    summary: {
      en: 'A typed Commander probe now captures an isolated WPS Writer shell and field command inventory so UI decisions can be grounded in the installed build.',
      zh: '类型化 Commander 探针现在可以捕获隔离的 WPS Writer 外壳与字段命令清单，让 UI 决策基于本机实际版本。',
    },
    highlights: [
      {
        title: {
          en: 'Three explicit profiles',
          zh: '三个明确 profile',
        },
        detail: {
          en: '`shell`, `fields`, and `all` keep the receipt bounded for everyday UX review while preserving a complete inventory when needed.',
          zh: '`shell`、`fields` 和 `all` 让日常 UX 复核保持有界，并在需要时保留完整命令清单。',
        },
      },
      {
        title: {
          en: 'Owned COM lifecycle',
          zh: 'COM 生命周期可控',
        },
        detail: {
          en: 'Each probe starts one hidden WPS Writer instance, records version, window bounds, Ribbon/status bars, and native command IDs, then closes only that instance.',
          zh: '每次探针启动一个隐藏的 WPS Writer 实例，记录版本、窗口几何、Ribbon/状态栏和原生命令 ID，然后只关闭该实例。',
        },
      },
      {
        title: {
          en: 'Evidence stays honest',
          zh: '证据边界保持诚实',
        },
        detail: {
          en: 'A3S Test Web/CDP remains the browser UI contract; the locked Windows CUA profile remains unsupported and the COM receipt is never a CI dependency.',
          zh: 'A3S Test Web/CDP 仍是浏览器 UI 契约；锁定的 Windows CUA profile 仍不支持，COM 结果也不是 CI 依赖。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the WPS UI probe',
          zh: '运行 WPS UI 探针',
        },
      },
      {
        href: {
          en: './components/document.html#common-live-fields',
          zh: './components/document.html#常用实时字段',
        },
        label: {
          en: 'Review field command evidence',
          zh: '查看字段命令证据',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.60.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.60.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.59.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer gains typed field settings',
      zh: 'Writer 新增类型化字段设置',
    },
    summary: {
      en: 'A bounded field-settings dialog now authors and edits WPS-compatible page, section, and bookmark page-reference formats with responsive A3S Test evidence.',
      zh: '有界字段设置弹窗现在可以插入和编辑兼容 WPS 的页码、节号与书签目标页码格式，并配套响应式 A3S Test 证据。',
    },
    highlights: [
      {
        title: {
          en: 'One typed insert/edit path',
          zh: '一条类型化插入/编辑路径',
        },
        detail: {
          en: 'Numeric page and section formats, date/time presets, bookmark targets, hyperlinks, and MERGEFORMAT retention share one controlled dialog and one undoable document update.',
          zh: '页码与节号数字格式、日期/时间预设、书签目标、超链接和 MERGEFORMAT 保留共享同一个受控弹窗与一次可撤销文档更新。',
        },
      },
      {
        title: {
          en: 'Native WPS formats stay honest',
          zh: '原生 WPS 格式保持诚实',
        },
        detail: {
          en: 'Unknown imported clock formats remain source-preserved until changed, while unsupported switches remain cached and diagnosed instead of being approximated.',
          zh: '导入的未知时间格式在用户修改前保持源格式；不支持的开关继续缓存并诊断，不被近似伪装。',
        },
      },
      {
        title: {
          en: 'A3S Test covers desktop and phone',
          zh: 'A3S Test 覆盖桌面与手机',
        },
        detail: {
          en: 'Web/CDP scenarios cover the two-step compact ribbon discovery, preview, F9 refresh, focus recovery, screenshots, accessibility, and empty console/page-error evidence at 1280px and 390px.',
          zh: 'Web/CDP 场景覆盖 1280px 与 390px 下的两步紧凑功能区发现、预览、F9 刷新、焦点恢复、截图、可访问性及空控制台/页面错误证据。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#common-live-fields',
          zh: './components/document.html#常用实时字段',
        },
        label: {
          en: 'Read the field settings guide',
          zh: '阅读字段设置指南',
        },
      },
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the common WPS field probe',
          zh: '运行常用 WPS 字段探针',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.59.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.59.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.58.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer fields gain bounded WPS numeric-switch parity',
      zh: 'Writer 字段补齐有界 WPS 数字开关对齐',
    },
    summary: {
      en: 'Common PAGE, NUMPAGES, SECTION, and PAGEREF fields now preserve the WPS numeric switches that are safe to evaluate, with a typed COM probe and browser-editor evidence.',
      zh: '常用 PAGE、NUMPAGES、SECTION 与 PAGEREF 字段现在保留可安全求值的 WPS 数字开关，并配套类型化 COM 探针与浏览器编辑器证据。',
    },
    highlights: [
      {
        title: {
          en: 'Switches stay typed and bounded',
          zh: '开关保持类型化且有界',
        },
        detail: {
          en: 'ROMAN, ALPHABETIC, Ordinal, hyperlink, and MERGEFORMAT combinations are parsed explicitly; unknown or malformed switches remain cached and diagnosed.',
          zh: '显式解析 ROMAN、ALPHABETIC、Ordinal、超链接与 MERGEFORMAT 组合；未知或格式错误的开关保持缓存并进入诊断。',
        },
      },
      {
        title: {
          en: 'WPS COM output is inspectable',
          zh: 'WPS COM 输出可检查',
        },
        detail: {
          en: 'The Commander CLI captures the installed WPS field instructions and version into an exact DOCX artifact without making COM a runtime or CI dependency.',
          zh: 'Commander CLI 将本机 WPS 字段指令与版本写入精确 DOCX 夹具，不把 COM 变成运行时或 CI 依赖。',
        },
      },
      {
        title: {
          en: 'A3S Test proves the editor path',
          zh: 'A3S Test 证明编辑器路径',
        },
        detail: {
          en: 'The numeric-field ACL checks visible results, focus, screenshots, accessibility, and empty console/page-error diagnostics; Windows CUA remains fail-closed.',
          zh: '数字字段 ACL 检查可见结果、焦点、截图、可访问性及空控制台/页面错误诊断；Windows CUA 继续失败即关闭。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#common-live-fields',
          zh: './components/document.html#常用实时字段',
        },
        label: {
          en: 'Read the field guide',
          zh: '阅读字段指南',
        },
      },
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the WPS field probe',
          zh: '运行 WPS 字段探针',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.58.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.58.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.57.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer connectors gain typed WPS shape parity',
      zh: 'Writer 连接符补齐 WPS 类型化形状对齐',
    },
    summary: {
      en: 'WPS connector shape types 32, 33, and 37 now share a bounded straight, elbow, and curved model across the Writer ribbon, live SVG, and DOCX round trip.',
      zh: 'WPS 连接符形状类型 32、33、37 现在在 Writer 功能区、实时 SVG 与 DOCX 往返之间共享有界的直线、肘形和曲线模型。',
    },
    highlights: [
      {
        title: {
          en: 'One connector-kind control',
          zh: '一个连接符类型控件',
        },
        detail: {
          en: 'Straight, elbow, and curved values are edited through the contextual ribbon and retained by Undo/Redo and compact layouts.',
          zh: '直线、肘形和曲线值通过上下文功能区编辑，并由撤销/重做与紧凑布局保持一致。',
        },
      },
      {
        title: {
          en: 'Native geometry stays typed',
          zh: '原生几何保持类型化',
        },
        detail: {
          en: 'WPS VML imports by o:spt value; DOCX export uses line, routed custom geometry, or quadratic geometry without claiming arbitrary route editing.',
          zh: 'WPS VML 按 o:spt 值导入；DOCX 导出使用直线、路由自定义几何或二次曲线几何，不虚构任意路由点编辑。',
        },
      },
      {
        title: {
          en: 'A3S Test proves the boundary',
          zh: 'A3S Test 证明边界',
        },
        detail: {
          en: 'The WPS kind matrix retains screenshots, accessibility snapshots, empty console/page-error diagnostics, and distinct SVG checks; Windows CUA remains fail-closed.',
          zh: 'WPS 类型矩阵保留截图、可访问性快照、空控制台/页面错误诊断及不同 SVG 检查；Windows CUA 继续失败即关闭。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the connector kind guide',
          zh: '阅读连接符类型指南',
        },
      },
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the WPS connector A3S Test',
          zh: '运行 WPS 连接符 A3S Test',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.57.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.57.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.56.1',
    date: '2026-09-06',
    kind: 'improved',
    surfaces: [
      'documentation',
      'spreadsheet',
      'presentation',
      'markdown',
      'pdf',
      'playground',
    ],
    title: {
      en: 'A3S Test gates become the daily editor contract',
      zh: 'A3S Test 门禁成为日常编辑器契约',
    },
    summary: {
      en: 'The five-surface local matrix now proves discoverable workflows at desktop and phone widths, with the browser/CDP evidence kept separate from the fail-closed Windows CUA capability.',
      zh: '五大编辑器本地矩阵现在覆盖桌面与手机宽度下的可发现工作流，并将浏览器/CDP 证据与 Windows CUA 失败即关闭能力明确分离。',
    },
    highlights: [
      {
        title: {
          en: 'Stable keyboard-first discovery',
          zh: '稳定的键盘优先发现路径',
        },
        detail: {
          en: 'Template cards are focused before activation, so below-fold Spreadsheet, Presentation, Markdown, and phone workflows use the same visibility behavior as keyboard users.',
          zh: '模板卡片在激活前先获得焦点，折叠线下的表格、演示、Markdown 与手机工作流因此与键盘用户使用相同的可见性行为。',
        },
      },
      {
        title: {
          en: 'PDF export survives slower observers',
          zh: 'PDF 导出兼容较慢观察器',
        },
        detail: {
          en: 'Blob URLs remain alive long enough for Web/CDP download observers, while the PDF ACL proves every page operation stays enabled after extraction.',
          zh: 'Blob URL 会保持足够长时间供 Web/CDP 下载观察器读取，同时 PDF ACL 证明抽取后所有页面操作仍保持可用。',
        },
      },
      {
        title: {
          en: 'Evidence is typed and auditable',
          zh: '证据具备类型与可审计性',
        },
        detail: {
          en: 'A3S Test retains screenshots, accessibility snapshots, empty console/page-error diagnostics, and explicit standalone-driver classifications; exact PDF bytes stay in the supplemental Playwright contract.',
          zh: 'A3S Test 保留截图、可访问性快照、空控制台/页面错误诊断及明确的 standalone 驱动分类；精确 PDF 字节内容继续由补充 Playwright 契约覆盖。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Run the A3S Test editor matrix',
          zh: '运行 A3S Test 编辑器矩阵',
        },
      },
      {
        href: {
          en: './components/pdf.html#page-organization',
          zh: './components/pdf.html#页面组织',
        },
        label: {
          en: 'Review PDF page organization',
          zh: '查看 PDF 页面组织',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.56.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.56.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.56.0',
    date: '2026-09-06',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer connectors gain WPS arrow-style parity',
      zh: 'Writer 连接符补齐 WPS 箭头样式对齐',
    },
    summary: {
      en: 'Straight connectors now preserve typed none, triangle, stealth, diamond, oval, and open endpoint-arrow intent through the responsive Writer ribbon and DOCX round trip.',
      zh: '直线连接符现在可以把无箭头、三角、隐形、菱形、圆形和开放端点箭头意图穿过响应式 Writer 功能区与 DOCX 往返。',
    },
    highlights: [
      {
        title: {
          en: 'One typed arrow-style model',
          zh: '一个类型化箭头样式模型',
        },
        detail: {
          en: 'The Connector ribbon, live SVG markers, compact controls, Undo/Redo, VML import, and DrawingML headEnd/tailEnd export share six bounded values.',
          zh: '连接符功能区、实时 SVG 标记、紧凑控件、撤销/重做、VML 导入和 DrawingML headEnd/tailEnd 导出共享六种有界值。',
        },
      },
      {
        title: {
          en: 'Windows COM 3/4 reference',
          zh: 'Windows COM 3/4 参考',
        },
        detail: {
          en: 'The installed WPS 12.0 probe emits VML startarrow=open and endarrow=classic for the bounded COM pair; classic reopens as the typed stealth arrow.',
          zh: '本机 WPS 12.0 探针在有界 COM 配对下生成 VML startarrow=open、endarrow=classic；classic 重新打开为类型化隐形箭头。',
        },
      },
      {
        title: {
          en: 'A3S Test evidence stays primary',
          zh: 'A3S Test 证据保持主导',
        },
        detail: {
          en: 'Real Windows CDP runs cover authoring and WPS import with screenshots, accessibility snapshots, and empty console/page-error diagnostics; the locked CUA matrix remains fail-closed on Windows.',
          zh: '真实 Windows CDP 运行覆盖创建与 WPS 导入并保留截图、可访问性快照及空控制台/页面错误诊断；锁定的 CUA 矩阵在 Windows 上继续失败即关闭。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the connector arrow-style guide',
          zh: '阅读连接符箭头样式指南',
        },
      },
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Review the A3S Test and CUA workflow',
          zh: '查看 A3S Test 与 CUA 流程',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.56.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.56.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.55.0',
    date: '2026-09-05',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer connectors gain WPS line-style parity',
      zh: 'Writer 连接符补齐 WPS 线型对齐',
    },
    summary: {
      en: 'Straight connectors now preserve native solid, dash, dot, and dash-dot intent from WPS through the responsive Writer ribbon and DOCX round trip.',
      zh: '直线连接符现在可以把 WPS 的实线、虚线、点线和点划线意图穿过响应式 Writer 功能区与 DOCX 往返。',
    },
    highlights: [
      {
        title: {
          en: 'One typed line-style model',
          zh: '一个类型化线型模型',
        },
        detail: {
          en: 'The Connector ribbon, live SVG, compact controls, Undo/Redo, VML import, and DrawingML a:prstDash export share solid, dash, dot, and dash-dot values.',
          zh: '连接符功能区、实时 SVG、紧凑控件、撤销/重做、VML 导入和 DrawingML a:prstDash 导出共享实线、虚线、点线和点划线值。',
        },
      },
      {
        title: {
          en: 'WPS COM evidence is recorded',
          zh: '记录 WPS COM 证据',
        },
        detail: {
          en: 'The installed WPS 12.0 probe records DashStyle=4 for a native dashed Shapes.AddConnector result; unsupported long-tail styles normalize fail-closed.',
          zh: '本机 WPS 12.0 探针记录原生虚线 Shapes.AddConnector 的 DashStyle=4；不支持的长尾线型会失败即归一化。',
        },
      },
      {
        title: {
          en: 'UI/UX evidence stays deep',
          zh: 'UI/UX 证据保持深入',
        },
        detail: {
          en: 'A3S Test ACLs cover authoring and WPS import with screenshots, accessibility snapshots, and empty console/page-error diagnostics; Playwright covers desktop and compact widths.',
          zh: 'A3S Test ACL 覆盖创建与 WPS 导入并保留截图、可访问性快照及空控制台/页面错误诊断；Playwright 覆盖桌面与紧凑视口。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the connector line-style guide',
          zh: '阅读连接符线型指南',
        },
      },
      {
        href: {
          en: './automation/index.html',
          zh: './automation/index.html',
        },
        label: {
          en: 'Review the A3S Test workflow',
          zh: '查看 A3S Test 流程',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.55.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.55.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.54.1',
    date: '2026-09-05',
    kind: 'fixed',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Windows editor runs stay deterministic',
      zh: 'Windows 编辑器运行保持确定性',
    },
    summary: {
      en: 'The local A3S Test loop now completes pinned CDP runs reliably and records responsive Writer discovery as part of the user-visible contract.',
      zh: '本地 A3S Test 流程现在可以可靠完成锁定版本的 CDP 运行，并把响应式 Writer 操作发现纳入用户可见契约。',
    },
    highlights: [
      {
        title: {
          en: 'Direct CDP lifecycle',
          zh: '直接 CDP 生命周期',
        },
        detail: {
          en: 'The Windows adapter invokes the standalone browser with the requested CDP port and resolves on process exit, avoiding detached-session port polling and inherited stdio hangs.',
          zh: 'Windows 适配器以请求的 CDP 端口直接调用独立浏览器，并在进程退出时完成，避免分离 session 端口轮询和继承 stdio 导致的挂起。',
        },
      },
      {
        title: {
          en: 'Responsive discovery is tested',
          zh: '测试响应式操作发现',
        },
        detail: {
          en: 'The Writer text-box ACL observes the desktop toolbar overflow and phone template scrolling before semantic clicks, matching controls a user can actually reach.',
          zh: 'Writer 文本框 ACL 会在语义点击前观察桌面工具栏溢出并滚动手机模板页面，对齐用户实际能够触达的控件。',
        },
      },
      {
        title: {
          en: 'Evidence stays bounded',
          zh: '证据保持有界',
        },
        detail: {
          en: 'Four desktop/phone scenarios pass with screenshots, accessibility, console, and page-error evidence; Windows CUA remains honestly marked unsupported by the locked 0.10.0 matrix.',
          zh: '四个桌面/手机场景均通过并保留截图、可访问性、控制台和页面错误证据；锁定的 0.10.0 矩阵仍如实将 Windows CUA 标记为 unsupported。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation.html',
          zh: './automation.html',
        },
        label: {
          en: 'Read the editor automation guide',
          zh: '阅读编辑器自动化指南',
        },
      },
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Review the responsive Writer contract',
          zh: '查看响应式 Writer 契约',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.54.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.54.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.54.0',
    date: '2026-09-05',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'A3S Test becomes the primary editor interaction contract',
      zh: 'A3S Test 成为编辑器交互主契约',
    },
    summary: {
      en: 'A bounded Commander operator now drives deterministic editor ACLs, exploratory agent sessions, WPS reference capture, and supplemental visual evidence from one matrix.',
      zh: '有界的 Commander 操作器现在通过同一份矩阵驱动确定性的编辑器 ACL、探索式智能体会话、WPS 参考采集和补充视觉证据。',
    },
    highlights: [
      {
        title: {
          en: 'A declarative five-surface CLI',
          zh: '声明式五编辑器 CLI',
        },
        detail: {
          en: 'Commander subcommands route fixtures, ACL checks and runs, visual baselines, WPS COM probes, and bounded agent lifecycle actions without a switch-based parser.',
          zh: 'Commander 子命令统一路由夹具、ACL 检查与执行、视觉基线、WPS COM 探针和有界智能体生命周期，不再依赖 switch/if-else 解析器。',
        },
      },
      {
        title: {
          en: 'Observed Writer connector editing',
          zh: '已观测的 Writer 连接符编辑',
        },
        detail: {
          en: 'The A3S Test agent and ACL create a straight connector, open its contextual ribbon, change width from 120 mm to 150 mm, and retain screenshot, accessibility, console, and page-error evidence.',
          zh: 'A3S Test 智能体与 ACL 创建直线连接符、打开上下文功能区、将宽度从 120 mm 改为 150 mm，并保留截图、可访问性、控制台和页面错误证据。',
        },
      },
      {
        title: {
          en: 'CUA capability is reported honestly',
          zh: '如实报告 CUA 能力',
        },
        detail: {
          en: 'The locked CUA Driver 0.10.0 matrix reports Windows GUI profiles as unsupported; browser/CDP evidence is kept separate from native WPS GUI claims, while the WPS COM probe remains an explicit reference workflow.',
          zh: '锁定的 CUA Driver 0.10.0 能力矩阵将 Windows GUI profile 如实标记为 unsupported；浏览器/CDP 证据与 WPS 原生 GUI 结论分开，WPS COM 探针仍是明确的参考流程。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './automation.html',
          zh: './automation.html',
        },
        label: {
          en: 'Read the editor automation guide',
          zh: '阅读编辑器自动化指南',
        },
      },
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Review the connector boundary',
          zh: '查看连接符边界',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.54.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.54.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.53.1',
    date: '2026-09-05',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer names the WPS connector boundary honestly',
      zh: 'Writer 明确记录 WPS 连接符边界',
    },
    summary: {
      en: 'Legacy VML connectors are diagnosed as compatibility records instead of being mistaken for editable text boxes or images.',
      zh: '传统 VML 连接符现在作为兼容性记录诊断，不会被误认为可编辑文本框或图片。',
    },
    highlights: [
      {
        title: {
          en: 'COM evidence drives the boundary',
          zh: 'COM 证据决定边界',
        },
        detail: {
          en: 'The installed WPS 12.0 probe records Shapes.AddConnector as a VML v:shape with o:spt="32" and #_x0000_t32, not as a text-bearing wps:wsp.',
          zh: '本机 WPS 12.0 探针记录 Shapes.AddConnector 为带 o:spt="32" 和 #_x0000_t32 的 VML v:shape，而不是带文字主体的 wps:wsp。',
        },
      },
      {
        title: {
          en: 'Fail-closed diagnostics',
          zh: '失败即明确诊断',
        },
        detail: {
          en: 'docx.connectors reports endpoint, routing, arrowhead, and floating-anchor semantics that the editable Writer model does not yet own, while connector-only pict containers avoid the generic image warning.',
          zh: 'docx.connectors 会报告可编辑 Writer 模型尚未负责的端点、路由、箭头和浮动锚点语义；只有连接符的 pict 容器也不会再触发泛化图片警告。',
        },
      },
      {
        title: {
          en: 'A bounded UI contract',
          zh: '有界的 UI 契约',
        },
        detail: {
          en: 'A deterministic DOCX fixture, focused Rstest, and local A3S Test ACL prove that import keeps the Writer editor usable, creates no text-box node, and captures accessibility plus browser diagnostics.',
          zh: '确定性的 DOCX 夹具、聚焦 Rstest 和本地 A3S Test ACL 证明导入后 Writer 仍可用、不创建文本框节点，并捕获可访问性与浏览器诊断。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the connector boundary',
          zh: '阅读连接符边界',
        },
      },
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Review the text-box model',
          zh: '查看文本框模型',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.53.0',
    date: '2026-09-05',
    kind: 'improved',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer text boxes speak the same shape language as WPS',
      zh: 'Writer 文本框与 WPS 使用同一套有界形状语义',
    },
    summary: {
      en: 'Five bounded shape presets now survive WPS import, the structured editor model, responsive controls, Undo/Redo, and native DOCX export.',
      zh: '五种有界形状现在可以穿过 WPS 导入、结构化编辑器模型、响应式控件、撤销/重做和原生 DOCX 导出。',
    },
    highlights: [
      {
        title: {
          en: 'Five shapes, one typed state',
          zh: '五种形状，一个类型化状态',
        },
        detail: {
          en: 'Rectangle, rounded rectangle, ellipse, diamond, and triangle share one bounded model. The contextual ribbon, live page, preview, PDF capture, and DOCX geometry stay aligned.',
          zh: '矩形、圆角矩形、椭圆、菱形和三角形共享一个有界模型。上下文功能区、实时页面、预览、PDF 捕获和 DOCX 几何保持一致。',
        },
      },
      {
        title: {
          en: 'Real WPS import stays editable',
          zh: '真实 WPS 导入后仍可编辑',
        },
        detail: {
          en: 'Isolated WPS mc:AlternateContent text-bearing shapes retain placement, fill, outline, padding, vertical anchor, text, and drawing identity. Mixed paragraphs and connectors remain diagnosed boundaries.',
          zh: '独立的 WPS mc:AlternateContent 文本形状保留位置、填充、轮廓、内边距、垂直锚点、文字和绘图身份。混合段落与连接符继续作为明确诊断边界。',
        },
      },
      {
        title: {
          en: 'Deep UI/UX evidence, not feature count',
          zh: '深度 UI/UX 证据，而不是功能堆砌',
        },
        detail: {
          en: 'Desktop and compact browser flows exercise every preset, contextual discovery, focus, accessibility, viewport containment, clean diagnostics, and WPS import Undo/Redo. The local A3S Test ACL records the same contract.',
          zh: '桌面与紧凑浏览器流程逐一测试每种形状、上下文发现性、焦点、可访问性、视口容纳、干净诊断以及 WPS 导入后的撤销/重做。本地 A3S Test ACL 记录同一契约。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the text-box guide',
          zh: '阅读文本框指南',
        },
      },
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Review the compatibility boundary',
          zh: '查看兼容性边界',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.52.0',
    date: '2026-09-05',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer whole-paragraph changes stay native and atomic',
      zh: 'Writer 整段修订保持原生与原子语义',
    },
    summary: {
      en: 'Bounded text-only paragraph insertions and deletions now review as one complete block and round-trip through native paragraph-mark records.',
      zh: '有界纯文字整段插入与删除现在作为一个完整块审核，并通过原生段落标记记录往返。',
    },
    highlights: [
      {
        title: {
          en: 'One paragraph, one decision',
          zh: '一个段落，一项决定',
        },
        detail: {
          en: 'Matching body and paragraph-mark insertions or deletions import as one review card. Accept, reject, Undo, export, and reopen preserve the complete block, author, date, and native tracked-change state.',
          zh: '正文与段落标记相匹配的插入或删除会导入为一张审核卡。接受、拒绝、撤销、导出和重开都会保留完整块、作者、日期和原生修订状态。',
        },
      },
      {
        title: {
          en: 'Native identities stay truthful',
          zh: '原生身份保持真实',
        },
        detail: {
          en: 'The bounded recognizer validates body and paragraph-mark IDs independently, matching the separate identities observed in the installed WPS reference while requiring the same author, timestamp, kind, and exact text-only structure.',
          zh: '有界识别器会分别校验正文和段落标记 ID，以匹配本机 WPS 参考中观察到的独立身份，同时要求作者、时间、类型和准确纯文字结构一致。',
        },
      },
      {
        title: {
          en: 'Responsive review with explicit limits',
          zh: '边界明确的响应式审核',
        },
        detail: {
          en: 'Isolated paragraph-break merges or splits, mixed content, malformed metadata, and over-limit input remain diagnostics. Desktop and 390 px real-DOCX flows cover touch sizing, keyboard navigation, focus, accessibility, screenshots, and clean browser diagnostics.',
          zh: '孤立段落分隔符合并或拆分、混合内容、格式错误元数据和超限输入仍进入诊断。桌面和 390px 真实 DOCX 流程覆盖触控尺寸、键盘导航、焦点、可访问性、截图与干净浏览器诊断。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#whole-paragraph-mark-revisions',
          zh: './components/document.html#整段段落标记修订',
        },
        label: {
          en: 'Read the paragraph-mark guide',
          zh: '阅读段落标记指南',
        },
      },
      {
        href: {
          en: './components/document.html#document-compare-and-combine',
          zh: './components/document.html#文档比较与合并',
        },
        label: {
          en: 'Read the Compare boundary',
          zh: '阅读比较边界',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.52.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.52.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.51.0',
    date: '2026-09-04',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer Compare pairs same-section text moves',
      zh: 'Writer 比较配对同一分节内文字移动',
    },
    summary: {
      en: 'Unique text ranges can now move between aligned simple paragraphs or headings in one section while remaining one atomic review decision.',
      zh: '唯一文字范围现在可以在同一分节内对齐的简单段落或标题之间移动，并保持为一项原子审核决定。',
    },
    highlights: [
      {
        title: {
          en: 'Paragraph scopes stay intact',
          zh: '段落范围保持完整',
        },
        detail: {
          en: 'Compare retains source and destination paragraph scopes instead of flattening the document tree. A unique lexical range with matching marks and carried separators becomes one paired move card across aligned simple text blocks in the same section.',
          zh: '比较会保留源段落和目标段落范围，不会压平文档树。同一分节内对齐简单文字块中，文字格式一致且分隔空白可携带的唯一词法范围会成为一张成对移动卡。',
        },
      },
      {
        title: {
          en: 'One atomic decision and native reopen',
          zh: '一项原子决定与原生重开',
        },
        detail: {
          en: 'Accept, reject, Undo, DOCX export, and reopen keep both sides of a cross-paragraph move together, preserving exact original and revised text with native w:moveFrom/w:moveTo records.',
          zh: '接受、拒绝、撤销、DOCX 导出和重开会保持跨段移动两侧同步，并通过原生 w:moveFrom/w:moveTo 记录保留精确的原稿和修订稿文字。',
        },
      },
      {
        title: {
          en: 'Boundaries remain explicit and tested',
          zh: '边界明确且经过测试',
        },
        detail: {
          en: 'Duplicates, mark mismatches, section-boundary moves, rich or relationship-bound content, tables, and over-limit inputs stay ordinary revisions or diagnostics. Desktop and 390 px review flows verify containment, accessibility, and clean browser diagnostics; WPS COM/UIA evidence remains documented.',
          zh: '重复候选、格式不一致、跨分节移动、富文本或关系绑定内容、表格及超出上限的输入仍保持普通修订或诊断。桌面和 390px 审阅流程验证容纳性、可访问性与干净浏览器诊断；WPS COM/UIA 证据继续记录在案。',
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
          en: 'Read the Compare guide',
          zh: '阅读比较指南',
        },
      },
      {
        href: {
          en: './components/collaboration.html#synchronize-move-revisions',
          zh: './components/collaboration.html#同步移动修订',
        },
        label: {
          en: 'Read the move contract',
          zh: '阅读移动协作合同',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.51.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.51.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.50.0',
    date: '2026-09-04',
    kind: 'new',
    surfaces: ['writer', 'playground', 'documentation'],
    title: {
      en: 'Writer Compare recognizes bounded text moves',
      zh: 'Writer 比较识别有界文字移动',
    },
    summary: {
      en: 'Same-paragraph lexical moves now appear as one deterministic review card while ambiguous or structurally rich edits remain safely explicit.',
      zh: '同一段落中的词法移动现在显示为一张确定性审核卡，含糊或结构复杂的编辑仍保持明确安全边界。',
    },
    highlights: [
      {
        title: {
          en: 'One aligned move, one decision',
          zh: '一个对齐移动，一项决定',
        },
        detail: {
          en: 'Compare pairs a unique lexical range found once in each aligned delete/insert chunk inside a simple paragraph or heading. Separators travel with the range, so one move card accepts, rejects, undoes, and reopens exactly.',
          zh: '比较会在简单段落或标题的对齐删除/插入块中，为各出现一次的唯一词法范围配对；分隔空白随范围移动，因此一张移动卡的接受、拒绝、撤销和重开都保持精确。',
        },
      },
      {
        title: {
          en: 'Fail closed where identity is unclear',
          zh: '身份不清时安全失败',
        },
        detail: {
          en: 'Duplicate candidates, rich or relationship-bound runs, cross-paragraph ranges, and over-limit text remain ordinary revisions or diagnostics instead of guessed moves.',
          zh: '重复候选、富文本或关系绑定运行、跨段范围和超出上限的文字会保留为普通修订或诊断，不会被猜测成移动。',
        },
      },
      {
        title: {
          en: 'WPS-referenced responsive review',
          zh: '参考 WPS 的响应式审阅',
        },
        detail: {
          en: 'The installed WPS 12.0 COM/UIA probe exposed the tested reorder as delete/insert records; A3S documents that evidence and verifies the local paired enhancement in desktop and 390 px browser flows with clean diagnostics.',
          zh: '本机 WPS 12.0 COM/UIA 探针将测试中的重排暴露为删除/插入记录；A3S 记录该证据，并在桌面和 390px 浏览器流程中验证本地成对增强与干净诊断。',
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
          en: 'Read the Compare guide',
          zh: '阅读比较指南',
        },
      },
      {
        href: {
          en: './components/collaboration.html#synchronize-move-revisions',
          zh: './components/collaboration.html#同步移动修订',
        },
        label: {
          en: 'Read the move contract',
          zh: '阅读移动协作合同',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.50.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.50.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.49.0',
    date: '2026-09-04',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground', 'documentation'],
    title: {
      en: 'Writer move revisions stay paired and native',
      zh: 'Writer 移动修订保持成对并原生往返',
    },
    summary: {
      en: 'Bounded text-only Word move pairs now form one review decision, round-trip as native move records, and keep unsupported structures explicitly diagnosed.',
      zh: '有界纯文字 Word 移动修订现在形成一条审核决定，以原生移动记录往返，并明确诊断不支持的结构。',
    },
    highlights: [
      {
        title: {
          en: 'One move, one review item',
          zh: '一次移动，一张审核卡',
        },
        detail: {
          en: 'Strict and transitional w:moveFrom/w:moveTo pairs with matching identity, author, date, and text import as one atomic move change. Accepting or rejecting it resolves both source and destination sides together.',
          zh: '严格或过渡命名空间中身份、作者、日期和文字一致的 w:moveFrom/w:moveTo 会导入为一条原子移动修订；接受或拒绝会同时处理源位置和目标位置。',
        },
      },
      {
        title: {
          en: 'Native output stays honest',
          zh: '原生输出保持诚实',
        },
        detail: {
          en: 'DOCX export rewrites transient wrappers into native moveFrom and moveTo elements, preserves review attribution, and reopens without private markers or negative IDs.',
          zh: 'DOCX 导出会把临时包装改写为原生 moveFrom 和 moveTo 元素，保留审核归属，重开时不泄露私有标记或负数身份。',
        },
      },
      {
        title: {
          en: 'Unsupported shapes fail closed',
          zh: '不支持的形态失败关闭',
        },
        detail: {
          en: 'Rich or relationship-bound content, range markers, malformed metadata, and unpaired moves remain compatibility diagnostics instead of being presented as editable semantics. The PDFium asset path is also stable for direct desktop and compact imports.',
          zh: '富内容、关系绑定内容、范围标记、格式错误元数据和未配对移动会保留为兼容性诊断，不会伪装成可编辑语义；PDFium 资源路径也已稳定，桌面和紧凑布局的直接导入都能就绪。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#move-revisions',
          zh: './components/document.html#移动修订',
        },
        label: {
          en: 'Read the move-revision guide',
          zh: '阅读移动修订指南',
        },
      },
      {
        href: {
          en: './components/collaboration.html#synchronize-move-revisions',
          zh: './components/collaboration.html#同步移动修订',
        },
        label: {
          en: 'Read the collaboration contract',
          zh: '阅读协作合同',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.49.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.49.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.48.1',
    date: '2026-09-03',
    kind: 'fixed',
    surfaces: ['writer', 'shared', 'playground', 'documentation'],
    title: {
      en: 'Writer selection controls keep their compact visual contract',
      zh: 'Writer 选区控件恢复紧凑视觉契约',
    },
    summary: {
      en: 'The floating selection toolbar no longer exposes browser-native outset borders on its underline and strikethrough split controls.',
      zh: '浮动选区工具栏的下划线和删除线拆分控件不再暴露浏览器原生 outset 边框。',
    },
    highlights: [
      {
        title: {
          en: 'Native defaults are reset',
          zh: '重置浏览器原生默认样式',
        },
        detail: {
          en: 'The two split controls now use the toolbar button baseline, including explicit appearance, border, background, spacing, and icon alignment rules.',
          zh: '两个拆分控件现在使用工具栏按钮基线，显式统一 appearance、边框、背景、间距和图标对齐规则。',
        },
      },
      {
        title: {
          en: 'One group, one divider',
          zh: '一个组合，一条分隔线',
        },
        detail: {
          en: 'The primary action and disclosure remain visually grouped with a single-pixel divider while hover, pressed, and focus states stay readable.',
          zh: '主操作和下拉入口保持视觉分组，只显示一条像素分隔线，同时保留清晰的悬停、按下和焦点状态。',
        },
      },
      {
        title: {
          en: 'Formatting semantics are unchanged',
          zh: '格式语义保持不变',
        },
        detail: {
          en: 'Underline and strikethrough commands, accessible labels, keyboard behavior, and their advanced style menus continue to use the existing typed contract.',
          zh: '下划线和删除线命令、可访问名称、键盘行为及高级样式菜单继续使用既有类型化合同。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#selection-toolbar-controls',
          zh: './components/document.html#选择工具栏控件',
        },
        label: {
          en: 'Read the selection-toolbar guide',
          zh: '阅读选区工具栏说明',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.48.1',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.48.1',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.48.0',
    date: '2026-09-03',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground', 'documentation'],
    title: {
      en: 'Writer content controls stay bounded and native',
      zh: 'Writer 内容控件保持有界并原生往返',
    },
    summary: {
      en: 'Inline plain-text and rich-text content controls now share one typed editor node, one lock boundary, and a strict native DOCX contract without a remote service.',
      zh: '行内纯文本和富文本内容控件现在共享一个类型化编辑器节点、一条锁定边界和严格的原生 DOCX 合同，不引入远程服务。',
    },
    highlights: [
      {
        title: {
          en: 'Author the whole intent in one dialog',
          zh: '在一个弹窗中完成完整意图',
        },
        detail: {
          en: 'The Insert ribbon and responsive dialog create or edit inline controls with aliases, program tags, plain or rich text, multiline behavior, border/tag/hidden appearance, and an optional color. Each accepted intent is one typed update and one Undo step.',
          zh: '“插入”功能区和响应式弹窗可以创建或编辑行内控件：显示名称、程序标签、纯文本或富文本、多行行为、边框/标签/隐藏外观以及可选颜色都在一次意图中完成；每次确定只产生一条类型化更新和一步撤销。',
        },
      },
      {
        title: {
          en: 'Locks are enforced below the UI',
          zh: '锁定在 UI 之下的事务边界生效',
        },
        detail: {
          en: 'Content and shell locks reject accidental typing, paste, replacement, deletion, and generic metadata writes. Explicit typed commands are required for unlock-sensitive operations, while accessible names follow the control alias or tag.',
          zh: '内容锁和控件锁会拒绝意外输入、粘贴、替换、删除和普通元数据写入；需要解锁的操作必须通过显式类型化命令完成，可访问名称会跟随控件显示名称或程序标签。',
        },
      },
      {
        title: {
          en: 'Native DOCX stays honest',
          zh: '原生 DOCX 边界保持诚实',
        },
        detail: {
          en: 'Direct paragraph `w:sdt` controls round-trip through strict and transitional WordprocessingML with collision-free IDs, rich runs, locks, multiline text, and Word 2012 appearance/color. Bindings, placeholders, repeating regions, form controls, nested structures, and relationship-bound content remain diagnosed safe text.',
          zh: '直接位于段落的 `w:sdt` 控件会在严格或过渡 WordprocessingML 中往返无冲突身份、富文本运行、锁定、多行文字及 Word 2012 外观/颜色。绑定、占位符、重复区域、表单控件、嵌套结构和关系绑定内容会被诊断并保留为安全文字。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-content-controls',
          zh: './components/document.html#原生内容控件',
        },
        label: {
          en: 'Read the content-control guide',
          zh: '阅读内容控件指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.48.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.48.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.47.0',
    date: '2026-09-03',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground', 'documentation'],
    title: {
      en: 'Writer common fields stay live and referenceable',
      zh: 'Writer 常用字段保持实时并可引用',
    },
    summary: {
      en: 'Word counts, character counts, and bookmark page references now share one bounded field model, one measured pagination source, and an honest native DOCX boundary.',
      zh: '字数、字符数和书签目标页码现在共享一个有界字段模型、一套实测分页来源，以及明确的原生 DOCX 边界。',
    },
    highlights: [
      {
        title: {
          en: 'Statistics exclude generated results',
          zh: '统计不把生成结果算入正文',
        },
        detail: {
          en: 'Insert exposes Word-compatible NUMWORDS and NUMCHARS fields. They count visible body text, preserve word boundaries around inline atoms, include spaces in character totals, and refresh with PAGE-family fields in one controlled update.',
          zh: '“插入”功能区提供兼容 Word 的 NUMWORDS 与 NUMCHARS 字段：统计可见正文、保留行内原子两侧的字边界、把空格计入字符数，并与 PAGE 系列字段在一次受控更新中刷新。',
        },
      },
      {
        title: {
          en: 'A bookmark can become a live target page',
          zh: '书签可以变成实时目标页码',
        },
        detail: {
          en: 'The Cross-reference dialog offers Insert target page for bookmarks. PAGEREF stores the stable bookmark identity and current name, follows normalization, and changes to an explicit missing state when the target disappears.',
          zh: '“交叉引用”弹窗为书签提供“插入目标页码”。PAGEREF 保存书签稳定身份和当前名称，跟随规范化更新，目标消失时切换为明确的缺失状态。',
        },
      },
      {
        title: {
          en: 'Native round trips fail closed',
          zh: '原生往返坚持安全失败',
        },
        detail: {
          en: 'The importer/exporter atomizes only complete inline common fields with the bounded switch grammar. Unsupported switches, missing targets, nested structures, and malformed fields remain cached text and receive compatibility diagnostics.',
          zh: '导入/导出只把完整行内常用字段和有界开关语法恢复为原子字段；不支持的开关、缺失目标、嵌套结构和损坏字段保留缓存文字并进入兼容性诊断。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#common-live-fields',
          zh: './components/document.html#常用实时字段',
        },
        label: {
          en: 'Read the common-fields guide',
          zh: '阅读常用字段指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.47.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.47.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.46.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground'],
    title: {
      en: 'Writer text boxes are editable and native',
      zh: 'Writer 文本框现在可编辑且原生往返',
    },
    summary: {
      en: 'A bounded text-box node connects familiar Writer controls, deterministic page rendering, and isolated WPS DrawingML round trips without a second document model.',
      zh: '有界文本框节点把熟悉的 Writer 控制、确定性的页面渲染与独立 WPS DrawingML 往返连接起来，无需第二套文档模型。',
    },
    highlights: [
      {
        title: {
          en: 'Insert once, edit in context',
          zh: '一次插入，上下文内编辑',
        },
        detail: {
          en: 'The Insert ribbon creates an isolated text box; its contextual ribbon edits inline or floating layout, millimeter geometry, offsets, fill, outline, padding, and vertical alignment with one Undo step per intent.',
          zh: '“插入”功能区创建隔离文本框；上下文功能区编辑嵌入或浮动布局、毫米几何尺寸、偏移、填充、轮廓、内边距和垂直对齐，每个意图只产生一步撤销。',
        },
      },
      {
        title: {
          en: 'One projection for page, preview, and PDF',
          zh: '页面、预览与 PDF 共用一个投影',
        },
        detail: {
          en: 'The same bounded dimensions and placement state drive live editing, read-only preview, PDF capture, and keep-together pagination behavior.',
          zh: '同一套有界尺寸与位置状态驱动实时编辑、只读预览、PDF 捕获和保持整块分页行为。',
        },
      },
      {
        title: {
          en: 'Native WPS shape with an honest boundary',
          zh: '原生 WPS 形状并明确边界',
        },
        detail: {
          en: 'Isolated `wps:wsp` shapes marked `txBox="1"` retain text, geometry, placement, fill, outline, body padding, vertical anchor, and drawing identity; mixed or malformed branches stay diagnosed on the normal compatibility path.',
          zh: '带 `txBox="1"` 的独立 `wps:wsp` 形状会保留文字、几何尺寸、位置、填充、轮廓、内边距、垂直锚点和绘图身份；混合或损坏分支留在普通兼容路径并明确诊断。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-editable-text-boxes',
          zh: './components/document.html#可编辑文本框',
        },
        label: {
          en: 'Read the text-box guide',
          zh: '阅读文本框指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.46.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.46.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.45.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground'],
    title: {
      en: 'Writer pictures can rotate and reflect locally',
      zh: 'Writer 图片现在可以在本地旋转和翻转',
    },
    summary: {
      en: 'A bounded picture-transform model brings quarter-turn rotation and reflection into the same controlled editor and native DOCX workflow.',
      zh: '有界的图片变换模型把 90° 旋转和翻转纳入同一个受控编辑器与原生 DOCX 工作流。',
    },
    highlights: [
      {
        title: {
          en: 'One familiar Picture ribbon',
          zh: '一个熟悉的“图片”功能区',
        },
        detail: {
          en: 'Accessible rotate-left, rotate-right, horizontal-flip, and vertical-flip actions sit beside existing wrap and alignment commands.',
          zh: '可访问的向左旋转、向右旋转、水平翻转和垂直翻转操作，与现有环绕和对齐命令放在一起。',
        },
      },
      {
        title: {
          en: 'The dialog keeps the full intent together',
          zh: '弹窗保留完整编辑意图',
        },
        detail: {
          en: 'Picture Properties exposes the same quarter-turn and reflection controls in a phone-safe layout; Apply is one controlled update and one Undo step.',
          zh: '“图片属性”在适配手机的布局中提供同样的 90° 旋转与翻转控制；确定只产生一次受控更新和一步撤销。',
        },
      },
      {
        title: {
          en: 'Native round trip with an honest boundary',
          zh: '原生往返并明确边界',
        },
        detail: {
          en: 'DOCX `a:xfrm` rotation and reflection values reopen as editable state; arbitrary-angle or malformed transforms are diagnosed and normalized safely.',
          zh: 'DOCX `a:xfrm` 的旋转与翻转值会重开为可编辑状态；任意角度或格式错误的变换会被诊断并安全归一化。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#built-in-picture-properties',
          zh: './components/document.html#图片属性',
        },
        label: {
          en: 'Read the picture-transform guide',
          zh: '阅读图片变换指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.45.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.45.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.44.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['spreadsheet', 'shared', 'playground'],
    title: {
      en: 'Spreadsheet dropdowns can follow a local driver',
      zh: '表格下拉列表现在可以跟随本地驱动值',
    },
    summary: {
      en: 'Bounded local INDIRECT sources bring dependent dropdowns into the same controlled Data Validation workflow without a remote service.',
      zh: '有界的本地 INDIRECT 来源把依赖下拉纳入同一个受控数据验证流程，不引入远程服务。',
    },
    highlights: [
      {
        title: {
          en: 'Author the dependency in the dialog',
          zh: '在弹窗中直接编写依赖关系',
        },
        detail: {
          en: 'The list source field accepts a bounded =INDIRECT(...) grammar made from quoted text, single-cell references, and concatenation, with a formula note and Sigma affordance.',
          zh: '序列来源字段支持由带引号文本、单元格引用和拼接组成的有界 =INDIRECT(...) 语法，并提供公式说明和 Sigma 图标提示。',
        },
      },
      {
        title: {
          en: 'Each row gets the right options',
          zh: '每一行都得到正确选项',
        },
        detail: {
          en: 'Relative drivers re-evaluate from each selected range anchor; named ranges and one-dimensional local areas resolve per cell, while an empty driver shows an empty list.',
          zh: '相对驱动值会从每个选定区域锚点逐单元格重新求值；工作簿名称和一维本地区域按单元格解析，驱动值为空时显示空列表。',
        },
      },
      {
        title: {
          en: 'Native and fail-closed',
          zh: '原生往返且安全失败',
        },
        detail: {
          en: 'The authored formula stays compact, runtime projection is capped at 1,024 source cells and 10,000 materialized cells, and native XLSX export/reopen keeps the list formula and names.',
          zh: '原始公式保持紧凑，运行时投影限制为最多 1,024 个来源单元格和 10,000 个物化单元格，原生 XLSX 导出/重开保留公式和名称。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#dependent-dropdown-lists',
          zh: './components/spreadsheet.html#依赖下拉列表',
        },
        label: {
          en: 'Read the dependent-dropdown guide',
          zh: '阅读依赖下拉指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.44.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.44.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.43.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['spreadsheet', 'shared', 'playground'],
    title: {
      en: 'Spreadsheet conditional formatting is now formula-editable',
      zh: '表格条件格式现在支持公式编辑',
    },
    summary: {
      en: 'A local, bounded formula rule closes the most common conditional-format gap while keeping precedence, blank ranges, and native XLSX semantics inspectable.',
      zh: '本地有界公式规则补上最常用的条件格式缺口，同时让优先级、空白区域和原生 XLSX 语义保持可检查。',
    },
    highlights: [
      {
        title: {
          en: 'Author in one rule manager',
          zh: '在同一个规则管理器中编辑',
        },
        detail: {
          en: 'Home → Conditional Formatting exposes a first-class Custom formula editor with relative/absolute references, cross-sheet cells, independent text/fill colors, and Stop-if-true ordering.',
          zh: '开始 → 条件格式提供一等自定义公式编辑器，支持相对/绝对引用、跨表单元格、独立文字/填充颜色和匹配后停止的优先级。',
        },
      },
      {
        title: { en: 'Local and bounded', zh: '本地且有界' },
        detail: {
          en: 'Only cached workbook values are read: formulas are capped at 255 Unicode characters and 1,024 referenced cells, with bounded blank-cell scans and fail-closed unsafe references.',
          zh: '只读取工作簿缓存值：公式最多 255 个 Unicode 字符、每次最多读取 1,024 个单元格，有界扫描空白单元格，遇到不安全引用会安全失败。',
        },
      },
      {
        title: {
          en: 'Native round trip and a real template',
          zh: '原生往返与真实模板',
        },
        detail: {
          en: 'XLSX expression records, differential styles, sqref, and priorities survive export and reopen; the public formula conditional-format template shows the workflow without a remote service.',
          zh: 'XLSX expression 条件记录、差异样式、sqref 和优先级可在导出与重开后保留；公开公式条件格式模板展示了无需远程服务的完整流程。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#formula-conditional-formatting',
          zh: './components/spreadsheet.html#公式条件格式',
        },
        label: {
          en: 'Read the formula conditional-format guide',
          zh: '阅读公式条件格式指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.43.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.43.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.42.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['spreadsheet', 'shared', 'playground'],
    title: {
      en: 'Spreadsheet rules can now be local custom formulas',
      zh: '表格验证规则现在支持本地自定义公式',
    },
    summary: {
      en: 'A bounded custom-formula rule brings dependent local checks into the same accessible Data Validation workflow without introducing a remote service.',
      zh: '有界的自定义公式规则把本地依赖检查纳入同一个可访问数据验证流程，不引入远程服务。',
    },
    highlights: [
      {
        title: { en: 'Author where you validate', zh: '在验证处直接编写' },
        detail: {
          en: 'The shared dialog accepts an optional = prefix, hides irrelevant numeric operators, and keeps the authored formula visible as a first-class rule.',
          zh: '共享弹窗支持可选的 = 前缀，隐藏无关的数值运算符，并把公式作为一等规则持续展示。',
        },
      },
      {
        title: { en: 'Relative and local by design', zh: '相对且本地可控' },
        detail: {
          en: 'The proposed value is substituted before evaluation; references anchor to each selected range and common sheet-qualified cells or ranges stay inside the workbook.',
          zh: '求值前会替换为待提交值；引用以每个选定区域为锚点，常用的带工作表单元格或区域引用始终留在工作簿内。',
        },
      },
      {
        title: {
          en: 'Fail closed, round-trip native',
          zh: '失败关闭，原生往返',
        },
        detail: {
          en: 'A 255-character formula and 1,024-cell read budget reject unsafe or uncached references, while native XLSX import/export and the public template retain the rule.',
          zh: '255 个字符和 1,024 个单元格读取上限会拒绝不安全或无缓存引用，同时原生 XLSX 导入/导出和公开模板都会保留规则。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#custom-formulas',
          zh: './components/spreadsheet.html#自定义公式',
        },
        label: {
          en: 'Read the custom-formula validation guide',
          zh: '阅读自定义公式验证指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.42.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.42.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.41.0',
    date: '2026-09-02',
    kind: 'improved',
    surfaces: ['spreadsheet', 'shared', 'playground'],
    title: {
      en: 'Spreadsheet validation alerts now match Office decisions',
      zh: '表格数据验证警告现在与 Office 决策一致',
    },
    summary: {
      en: 'Invalid Spreadsheet edits now explain the rule and offer the same Stop, Warning, or Information decision instead of silently taking one browser path.',
      zh: '非法表格输入现在会解释规则，并提供与 Office 一致的停止、警告或信息决策，不再把所有情况合并成一条浏览器路径。',
    },
    highlights: [
      {
        title: { en: 'Stop means stop', zh: '停止就是阻止' },
        detail: {
          en: 'An accessible notice discards the invalid draft and keeps the original cell selected, including after Enter from the formula bar.',
          zh: '可访问提示会丢弃非法草稿并保持原单元格选中，即使输入来自公式栏的 Enter 提交。',
        },
      },
      {
        title: {
          en: 'Warning and Information are deliberate',
          zh: '警告与信息可明确决策',
        },
        detail: {
          en: 'Warning offers Continue input or Cancel; Information offers Keep input or Return to edit, with authored copy and the current value in context.',
          zh: '警告提供“继续输入/取消”，信息提供“保留输入/返回修改”，并在上下文中展示自定义文案与当前值。',
        },
      },
      {
        title: { en: 'One controlled commit', zh: '一次受控提交' },
        detail: {
          en: 'A confirmed invalid value uses a typed, single-use bypass, so selection, focus, Undo, collaboration, and native XLSX errorStyle fidelity remain bounded.',
          zh: '确认保留的非法值只使用一次类型化绕过，因此选区、焦点、撤销、协作和原生 XLSX errorStyle 保真度都保持有界。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/spreadsheet.html#office-style-error-alert-branches',
          zh: './components/spreadsheet.html#与-office-一致的错误警告分支',
        },
        label: {
          en: 'Read the Spreadsheet validation interaction guide',
          zh: '阅读表格数据验证交互指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.41.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.41.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.40.0',
    date: '2026-09-02',
    kind: 'new',
    surfaces: ['writer', 'shared', 'playground'],
    title: {
      en: 'Writer numbering changes are now reviewable',
      zh: 'Writer 编号变化现在可以完整审阅',
    },
    summary: {
      en: 'Ordered-list style and starting-number edits now form one atomic revision across review, collaboration, native DOCX, and Undo.',
      zh: '有序列表样式与起始编号修改现在会形成一条原子修订，贯通审核、协作、原生 DOCX 与撤销。',
    },
    highlights: [
      {
        title: { en: 'One list, one intent', zh: '一个列表，一个意图' },
        detail: {
          en: 'The Numbering card accepts the current list or restores the complete original style and start without changing list text.',
          zh: '“编号格式”卡可以接受当前列表，或恢复完整原样式与起始值，而不改动列表文字。',
        },
      },
      {
        title: { en: 'Native revision fidelity', zh: '原生修订保真' },
        detail: {
          en: 'Common single-level decimal, letter, and Roman w:numberingChange records import, export, and reopen with fail-closed diagnostics.',
          zh: '常见单层十进制、字母与罗马数字 w:numberingChange 可导入、导出和重开，异常形式会失败关闭。',
        },
      },
      {
        title: { en: 'Protected collaboration', zh: '受保护的协作' },
        detail: {
          en: 'Yjs/Yrs retain live numbering metadata and immutable decisions across persistence while suggest mode rejects metadata tampering.',
          zh: 'Yjs/Yrs 会跨持久化保留编号元数据与不可变决定，建议模式则拒绝篡改审核元数据。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/document.html#ordered-list-numbering-revisions',
          zh: './components/document.html#有序列表编号修订',
        },
        label: {
          en: 'Read the Writer numbering review guide',
          zh: '阅读 Writer 编号修订指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.40.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.40.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
  {
    version: '0.39.0',
    date: '2026-09-01',
    kind: 'new',
    surfaces: ['presentation', 'playground'],
    title: {
      en: 'Presentation objects can now enter and exit',
      zh: '演示对象现在可以完整登场与退场',
    },
    summary: {
      en: 'A slide object can combine one entrance and one exit effect through the same ordered authoring, playback, collaboration, and native PPTX model.',
      zh: '同一个幻灯片对象现在可以通过统一的排序、创作、放映、协作与原生 PPTX 模型组合一条进入和一条退出效果。',
    },
    highlights: [
      {
        title: { en: 'Object-centric authoring', zh: '对象式创作' },
        detail: {
          en: 'The Animation tab switches explicitly between Entrance and Exit, then keeps effects, directions, triggers, timing, ordering, and preview in context.',
          zh: '动画选项卡可明确切换进入与退出，并在当前上下文中统一管理效果、方向、触发、计时、顺序和预览。',
        },
      },
      {
        title: { en: 'Composable playback', zh: '可组合放映' },
        detail: {
          en: 'Eight bounded effects share one cue model; sequential effects on the same object compose, while overlapping intervals fail closed.',
          zh: '八种有界效果共用一套提示模型；同一对象的连续效果可以组合，重叠时间段则会在修改前拒绝。',
        },
      },
      {
        title: { en: 'Native PPTX evidence', zh: '原生 PPTX 证据' },
        detail: {
          en: 'Entrance and exit timing classes, in/out transitions, object targets, triggers, timing, and direction survive export, import, and a second reopen.',
          zh: '进入/退出计时类型、in/out transition、对象目标、触发、计时和方向都能通过导出、导入与二次重开。',
        },
      },
    ],
    links: [
      {
        href: {
          en: './components/presentation.html#entrance-and-exit-animations',
          zh: './components/presentation.html#进入与退出动画',
        },
        label: {
          en: 'Read the Presentation animation guide',
          zh: '阅读演示文稿动画指南',
        },
      },
      {
        href: {
          en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.39.0',
          zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.39.0',
        },
        label: { en: 'GitHub Release', zh: 'GitHub Release' },
      },
    ],
  },
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
