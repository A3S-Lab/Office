import type { Cell, CellMatrix, Sheet } from '@fortune-sheet/core';
import type {
  WorkArtifact,
  WorkArtifactContent,
  WorkArtifactKind,
  WorkPresentationContent,
  WorkSlide,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
  WorkTemplate,
} from './work-types';
import {
  DOCUMENT_RUN_BORDER_ATTRIBUTE,
  type DocumentRunBorder,
  documentRunBorderDomAttributes,
} from './work-document-run-border';
import {
  DOCUMENT_RUN_SHADING_ATTRIBUTE,
  type DocumentRunShading,
  documentRunShadingDomAttributes,
} from './work-document-run-shading';
import { documentProofingDomAttributes } from './work-document-proofing';
import type { WorkDocumentScriptFontSlot } from './work-document-script-fonts';

export const WORK_TEMPLATES: WorkTemplate[] = [
  {
    id: 'blank-document',
    kind: 'document',
    name: '空白文字',
    description: '从一张干净的 A4 页面开始',
    accent: '#2f6fed',
  },
  {
    id: 'project-brief',
    kind: 'document',
    name: '项目方案',
    description: '目标、范围、里程碑与风险',
    accent: '#536de2',
  },
  {
    id: 'text-effects',
    kind: 'document',
    name: '文字效果',
    description: '空心、阴影、阳文与阴文',
    accent: '#6b5bd2',
  },
  {
    id: 'run-borders',
    kind: 'document',
    name: '字符边框',
    description: '原生线型、颜色、宽度、间距与阴影',
    accent: '#4472c4',
  },
  {
    id: 'run-shading',
    kind: 'document',
    name: '字符底纹',
    description: '原生图案、前景色、背景色与显式重置',
    accent: '#70ad47',
  },
  {
    id: 'proofing-languages',
    kind: 'document',
    name: '校对语言',
    description: '拉丁、东亚、双向文字与校对排除',
    accent: '#2f6fed',
  },
  {
    id: 'blank-markdown',
    kind: 'markdown',
    name: '空白 Markdown',
    description: '用轻量标记编写结构化内容',
    accent: '#586574',
  },
  {
    id: 'blank-spreadsheet',
    kind: 'spreadsheet',
    name: '空白表格',
    description: '公式、表格、筛选与多工作表',
    accent: '#16a36a',
  },
  {
    id: 'quarterly-plan',
    kind: 'spreadsheet',
    name: '季度计划',
    description: '目标进度与预算跟踪',
    accent: '#168f72',
  },
  {
    id: 'data-validation',
    kind: 'spreadsheet',
    name: '数据验证',
    description: '下拉列表、输入提示与错误警告',
    accent: '#13795b',
  },
  {
    id: 'blank-presentation',
    kind: 'presentation',
    name: '空白演示',
    description: '16:9 宽屏演示文稿',
    accent: '#e16b3d',
  },
  {
    id: 'strategy-deck',
    kind: 'presentation',
    name: '策略汇报',
    description: '结论先行的三页汇报',
    accent: '#c85637',
  },
];

export function createWorkArtifact(templateId: string): WorkArtifact {
  const template =
    WORK_TEMPLATES.find((item) => item.id === templateId) ?? WORK_TEMPLATES[0];
  const now = Date.now();
  return {
    id: createWorkId('artifact'),
    kind: template.kind,
    title: initialTitle(template.id, template.kind),
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: contentForTemplate(template.id),
  };
}

export function createWorkId(prefix: string): string {
  const random =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `${prefix}-${random}`;
}

function initialTitle(templateId: string, kind: WorkArtifactKind): string {
  const titles: Record<string, string> = {
    'project-brief': '新项目方案',
    'text-effects': '文字效果示例',
    'run-borders': '字符边框示例',
    'run-shading': '字符底纹示例',
    'proofing-languages': '校对语言示例',
    'quarterly-plan': '季度执行计划',
    'data-validation': '数据验证示例',
    'strategy-deck': '业务策略汇报',
  };
  if (titles[templateId]) return titles[templateId];
  if (kind === 'document') return '无标题文字';
  if (kind === 'markdown') return '无标题 Markdown';
  if (kind === 'spreadsheet') return '无标题表格';
  return '无标题演示';
}

function contentForTemplate(templateId: string): WorkArtifactContent {
  if (templateId === 'project-brief') {
    return {
      type: 'document',
      pageSize: 'a4',
      html: [
        '<h1>新项目方案</h1>',
        '<p><strong>负责人：</strong>项目团队　　<strong>更新日期：</strong>今天</p>',
        '<blockquote><p>用一句话说明这项工作的目标，以及完成后会带来什么变化。</p></blockquote>',
        '<h2>背景与目标</h2>',
        '<p>描述当前情况、核心问题和可衡量的成功标准。</p>',
        '<h2>工作范围</h2>',
        '<ul><li><p>需要完成的关键交付物</p></li><li><p>明确不在本期范围内的事项</p></li></ul>',
        '<h2>里程碑</h2>',
        '<ol><li><p>方案确认</p></li><li><p>执行与评审</p></li><li><p>交付与复盘</p></li></ol>',
        '<h2>风险与决策</h2>',
        '<p>记录尚未解决的问题、依赖和决策负责人。</p>',
      ].join(''),
    };
  }
  if (templateId === 'text-effects') {
    return {
      type: 'document',
      pageSize: 'a4',
      html: [
        '<h1>原生文字效果</h1>',
        '<p>选择下面任一示例并打开字体高级设置，可以组合空心与阴影，或在互斥的阳文和阴文之间切换。</p>',
        '<h2>可组合效果</h2>',
        '<p><span data-office-legacy-text-outline="true" data-office-legacy-text-shadow="true">空心 + 阴影</span></p>',
        '<p><span data-office-legacy-text-outline="true">空心</span></p>',
        '<p><span data-office-legacy-text-shadow="true">阴影</span></p>',
        '<h2>互斥效果</h2>',
        '<p><span data-office-legacy-text-emboss="true">阳文</span></p>',
        '<p><span data-office-legacy-text-imprint="true">阴文</span></p>',
      ].join(''),
    };
  }
  if (templateId === 'run-borders') {
    return {
      type: 'document',
      pageSize: 'a4',
      html: [
        '<h1>原生字符边框</h1>',
        '<p>选择任一示例，可从开始选项卡直接开关边框，或在字体高级设置中编辑完整的原生线型、颜色、宽度、文字间距、阴影和框架属性。</p>',
        '<h2>常用线型</h2>',
        `<p>${runBorderTemplateSpan({ style: 'single', color: { value: '#4472c4' }, size: 4, space: 1 }, '单实线字符边框')}</p>`,
        `<p>${runBorderTemplateSpan({ style: 'double', color: { value: '#c00000' }, size: 8, space: 2 }, '双线字符边框')}</p>`,
        `<p>${runBorderTemplateSpan({ style: 'wave', color: { value: '#7030a0' }, size: 8, space: 2 }, '波浪字符边框')}</p>`,
        '<h2>高级属性</h2>',
        `<p>${runBorderTemplateSpan({ style: 'thinThickMediumGap', color: { value: '#0070c0' }, size: 16, space: 3, shadow: true, frame: true }, '阴影与框架字符边框')}</p>`,
      ].join(''),
    };
  }
  if (templateId === 'run-shading') {
    return {
      type: 'document',
      pageSize: 'a4',
      html: [
        '<h1>原生字符底纹</h1>',
        '<p>选择任一示例并打开字体高级设置，可以编辑完整的原生图案、前景色、背景色，或写入显式无底纹重置。</p>',
        '<h2>基本填充</h2>',
        `<p>${runShadingTemplateSpan({ pattern: 'clear', fill: { value: '#fff2cc' } }, '清除图案使用背景色')}</p>`,
        `<p>${runShadingTemplateSpan({ pattern: 'solid', color: { value: '#4472c4' } }, '实心图案使用前景色')}</p>`,
        '<h2>图案与密度</h2>',
        `<p>${runShadingTemplateSpan({ pattern: 'diagCross', color: { value: '#c00000' }, fill: { value: '#fce4d6' } }, '对角交叉字符底纹')}</p>`,
        `<p>${runShadingTemplateSpan({ pattern: 'pct25', color: { value: '#4472c4' }, fill: { value: '#ddebf7' } }, '25% 字符底纹')}</p>`,
        `<p>${runShadingTemplateSpan({ pattern: 'thinHorzStripe', color: { value: '#70ad47' }, fill: { value: '#e2f0d9' } }, '细水平条纹字符底纹')}</p>`,
      ].join(''),
    };
  }
  if (templateId === 'proofing-languages') {
    const languages = {
      latin: 'en-US',
      eastAsia: 'zh-CN',
      bidi: 'ar-SA',
    } as const;
    return {
      type: 'document',
      pageSize: 'a4',
      html: [
        '<h1>原生校对语言</h1>',
        '<p>在“审阅”选项卡中打开“设置校对语言”，可以独立编辑拉丁、东亚和双向文字语言，并决定文字是否参与拼写与语法检查。</p>',
        '<h2>按文字系统设置</h2>',
        `<p>${proofingLanguageTemplateSpan(languages, false, 'ascii', 'English proofing language')}</p>`,
        `<p>${proofingLanguageTemplateSpan(languages, false, 'eastAsia', '简体中文校对语言')}</p>`,
        `<p dir="rtl">${proofingLanguageTemplateSpan(languages, false, 'complexScript', 'لغة التدقيق العربية')}</p>`,
        '<h2>校对行为</h2>',
        `<p>${proofingLanguageTemplateSpan({ latin: 'en-GB' }, false, 'ascii', 'This text participates in proofing.')}</p>`,
        `<p>${proofingLanguageTemplateSpan({ latin: 'x-none' }, true, 'ascii', 'A3S-API-v2: this product identifier is excluded from proofing.')}</p>`,
      ].join(''),
    };
  }
  if (templateId === 'quarterly-plan') {
    return { type: 'spreadsheet', sheets: quarterlyPlanSheets() };
  }
  if (templateId === 'data-validation') {
    return { type: 'spreadsheet', sheets: dataValidationTemplateSheets() };
  }
  if (templateId === 'strategy-deck') {
    return strategyPresentation();
  }
  if (templateId === 'blank-spreadsheet') {
    return { type: 'spreadsheet', sheets: [blankSheet()] };
  }
  if (templateId === 'blank-markdown') {
    return { type: 'markdown', markdown: '' };
  }
  if (templateId === 'blank-presentation') {
    return { type: 'presentation', slides: [blankSlide()] };
  }
  return {
    type: 'document',
    pageSize: 'a4',
    html: '<p></p>',
  };
}

function runBorderTemplateSpan(
  border: DocumentRunBorder,
  text: string,
): string {
  const attributes = documentRunBorderDomAttributes(border);
  return `<span ${DOCUMENT_RUN_BORDER_ATTRIBUTE}='${attributes[DOCUMENT_RUN_BORDER_ATTRIBUTE]}' style="${attributes.style}">${text}</span>`;
}

function runShadingTemplateSpan(
  shading: DocumentRunShading,
  text: string,
): string {
  const attributes = documentRunShadingDomAttributes(shading);
  return `<span ${DOCUMENT_RUN_SHADING_ATTRIBUTE}='${attributes[DOCUMENT_RUN_SHADING_ATTRIBUTE]}' style="${attributes.style}">${text}</span>`;
}

function proofingLanguageTemplateSpan(
  languages: Record<string, string>,
  noProof: boolean,
  slot: WorkDocumentScriptFontSlot,
  text: string,
): string {
  const attributes = documentProofingDomAttributes(languages, noProof, slot);
  return `<span ${Object.entries(attributes)
    .map(([name, value]) => `${name}='${value}'`)
    .join(' ')}>${text}</span>`;
}

function blankSheet(): Sheet {
  return {
    id: createWorkId('sheet'),
    name: '工作表1',
    status: 1,
    order: 0,
    row: 60,
    column: 26,
    data: emptyMatrix(60, 26),
  };
}

function quarterlyPlanSheets(): Sheet[] {
  const data = emptyMatrix(40, 12);
  data[0][0] = styledCell('季度执行计划', {
    bl: 1,
    fs: 16,
    fc: '#ffffff',
    bg: '#168f72',
  });
  data[2][0] = headerCell('目标');
  data[2][1] = headerCell('负责人');
  data[2][2] = headerCell('一月');
  data[2][3] = headerCell('二月');
  data[2][4] = headerCell('三月');
  data[2][5] = headerCell('完成率');
  data[2][6] = headerCell('状态');
  const rows: Array<[string, string, number, number, number, string, string]> =
    [
      ['客户洞察报告', '林岚', 1, 1, 0, '=SUM(C4:E4)/3', '进行中'],
      ['新版发布', '周启', 0.8, 0.6, 0, '=AVERAGE(C5:E5)', '有风险'],
      ['渠道增长', '陈一', 1, 0.9, 0.7, '=AVERAGE(C6:E6)', '正常'],
      ['团队能力建设', '项目组', 1, 1, 1, '=AVERAGE(C7:E7)', '已完成'],
    ];
  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      data[rowIndex + 3][columnIndex] = styledCell(value, {
        bg: rowIndex % 2 ? '#f7faf9' : '#ffffff',
        ...(columnIndex >= 2 && columnIndex <= 5
          ? {
              ct: { fa: '0%', t: 'n' },
              ...(typeof value === 'number'
                ? { m: `${Math.round(value * 100)}%` }
                : {}),
            }
          : {}),
      });
    });
  });
  return [
    {
      id: createWorkId('sheet'),
      name: '执行看板',
      status: 1,
      order: 0,
      row: 40,
      column: 12,
      data,
      config: {
        columnlen: { 0: 180, 1: 96, 2: 76, 3: 76, 4: 76, 5: 96, 6: 96 },
        rowlen: { 0: 34, 2: 28 },
        merge: {
          '0_0': { r: 0, c: 0, rs: 1, cs: 7 },
        },
      },
    },
  ];
}

function dataValidationTemplateSheets(): WorkSpreadsheetSheet[] {
  const inputs = emptyMatrix(24, 8);
  ['Task', 'State', 'Due date', 'Priority', 'Owner'].forEach(
    (value, column) => {
      inputs[0][column] = headerCell(value);
    },
  );
  const rows: Array<[string, string, string, number, string]> = [
    ['Confirm requirements', 'Ready', '2026-09-05', 2, 'Avery'],
    ['Review integration', 'In review', '2026-09-12', 3, 'Morgan'],
    ['Resolve blockers', 'Blocked', '2026-09-18', 5, 'Riley'],
    ['Publish preview', 'Ready', '2026-09-24', 1, 'Jordan'],
    ['Ship release', 'In review', '2026-09-30', 4, 'Taylor'],
  ];
  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      inputs[rowIndex + 1][columnIndex] = styledCell(value, {
        bg: rowIndex % 2 ? '#f4faf7' : '#ffffff',
      });
    });
  });

  return [
    {
      id: createWorkId('sheet'),
      name: 'Inputs',
      status: 1,
      order: 0,
      row: 24,
      column: 8,
      data: inputs,
      dataValidationRanges: [
        {
          ranges: [{ row: [1, 5], column: [1, 1] }],
          item: dataValidationTemplateItem({
            type: 'dropdown',
            rangeTxt: 'B2:B6',
            value1: "'Lists'!A1:A3",
            allowBlank: false,
            showDropdownArrow: true,
            errorStyle: 'stop',
            errorTitle: 'Invalid state',
            errorMessage: 'Choose a state from the list.',
            hintShow: true,
            hintTitle: 'Workflow state',
            hintValue: 'Choose Ready, Blocked, or In review.',
          }),
        },
        {
          ranges: [{ row: [1, 5], column: [2, 2] }],
          item: dataValidationTemplateItem({
            type: 'date',
            type2: 'between',
            rangeTxt: 'C2:C6',
            value1: '2026-01-01',
            value2: '2026-12-31',
            errorStyle: 'information',
            errorTitle: 'Date outside 2026',
            errorMessage: 'Enter a date in calendar year 2026.',
            hintShow: true,
            hintTitle: 'Due date',
            hintValue: 'Use a date between 2026-01-01 and 2026-12-31.',
          }),
        },
        {
          ranges: [{ row: [1, 5], column: [3, 3] }],
          item: dataValidationTemplateItem({
            type: 'number_integer',
            type2: 'between',
            rangeTxt: 'D2:D6',
            value1: '1',
            value2: '5',
            allowBlank: false,
            errorStyle: 'warning',
            errorTitle: 'Priority outside range',
            errorMessage: 'Enter a whole number from 1 through 5.',
            hintShow: true,
            hintTitle: 'Priority',
            hintValue: '1 is highest priority; 5 is lowest.',
          }),
        },
      ],
      luckysheet_select_save: [
        {
          row: [1, 5],
          column: [1, 1],
          row_focus: 1,
          column_focus: 1,
        },
      ],
      config: {
        columnlen: { 0: 190, 1: 110, 2: 118, 3: 84, 4: 104 },
        rowlen: { 0: 30 },
      },
    },
    {
      id: createWorkId('sheet'),
      name: 'Lists',
      status: 0,
      order: 1,
      row: 12,
      column: 3,
      data: [
        [styledCell('Ready')],
        [styledCell('Blocked')],
        [styledCell('In review')],
      ],
      config: { columnlen: { 0: 120 } },
    },
  ];
}

function dataValidationTemplateItem(
  overrides: Partial<WorkSpreadsheetDataValidationItem>,
): WorkSpreadsheetDataValidationItem {
  return {
    type: 'dropdown',
    type2: '',
    rangeTxt: '',
    value1: '',
    value2: '',
    validity: '',
    remote: false,
    allowBlank: true,
    showDropdownArrow: true,
    prohibitInput: true,
    errorStyle: 'stop',
    errorTitle: '',
    errorMessage: '',
    hintShow: false,
    hintTitle: '',
    hintValue: '',
    checked: false,
    ...overrides,
  };
}

function emptyMatrix(rows: number, columns: number): CellMatrix {
  return Array.from({ length: rows }, () =>
    Array<Cell | null>(columns).fill(null),
  );
}

function styledCell(value: string | number, style: Partial<Cell> = {}): Cell {
  const formula =
    typeof value === 'string' && value.startsWith('=') ? value : undefined;
  return {
    v: formula ? undefined : value,
    m: formula ? '' : String(value),
    f: formula,
    ...style,
  };
}

function headerCell(value: string): Cell {
  return styledCell(value, {
    bl: 1,
    fc: '#215446',
    bg: '#dff3ec',
    ht: 0,
    vt: 0,
  });
}

function blankSlide(): WorkSlide {
  return {
    id: createWorkId('slide'),
    name: '标题幻灯片',
    background: '#ffffff',
    elements: [
      {
        id: createWorkId('element'),
        type: 'text',
        x: 12,
        y: 25,
        width: 76,
        height: 18,
        text: '',
        fontSize: 34,
        color: '#172033',
        fill: 'transparent',
        bold: true,
        align: 'center',
        placeholder: {
          key: 'title',
          type: 'title',
          prompt: '单击添加标题',
        },
      },
      {
        id: createWorkId('element'),
        type: 'text',
        x: 18,
        y: 49,
        width: 64,
        height: 10,
        text: '',
        fontSize: 17,
        color: '#727b8f',
        fill: 'transparent',
        bold: false,
        align: 'center',
        placeholder: {
          key: 'subtitle',
          type: 'subtitle',
          prompt: '添加副标题',
        },
      },
    ],
  };
}

function strategyPresentation(): WorkPresentationContent {
  const slides: WorkSlide[] = [
    {
      id: createWorkId('slide'),
      name: '封面',
      background: '#16213d',
      elements: [
        {
          id: createWorkId('element'),
          type: 'shape',
          x: 8,
          y: 12,
          width: 9,
          height: 3,
          text: '',
          fontSize: 12,
          color: '#ffffff',
          fill: '#ffb15a',
          bold: false,
          align: 'left',
          radius: 2,
        },
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 30,
          width: 72,
          height: 24,
          text: '业务策略汇报',
          fontSize: 38,
          color: '#ffffff',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 58,
          width: 62,
          height: 10,
          text: '把最重要的结论放在标题中',
          fontSize: 17,
          color: '#b8c4df',
          fill: 'transparent',
          bold: false,
          align: 'left',
        },
      ],
    },
    {
      id: createWorkId('slide'),
      name: '核心判断',
      background: '#f7f4ee',
      elements: [
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 10,
          width: 84,
          height: 11,
          text: '01　核心判断',
          fontSize: 15,
          color: '#b44e34',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 27,
          width: 76,
          height: 22,
          text: '用一句可以独立成立的话，说明我们看到了什么。',
          fontSize: 31,
          color: '#20273a',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: createWorkId('element'),
          type: 'shape',
          x: 8,
          y: 60,
          width: 84,
          height: 22,
          text: '关键证据或数据',
          fontSize: 18,
          color: '#ffffff',
          fill: '#b44e34',
          bold: true,
          align: 'center',
          radius: 3,
        },
      ],
    },
    {
      id: createWorkId('slide'),
      name: '下一步',
      background: '#ffffff',
      elements: [
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 10,
          width: 84,
          height: 12,
          text: '02　下一步',
          fontSize: 15,
          color: '#b44e34',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: createWorkId('element'),
          type: 'text',
          x: 8,
          y: 28,
          width: 84,
          height: 46,
          text: '1　确认优先级\n2　指定负责人\n3　设定可验证的里程碑',
          fontSize: 26,
          color: '#20273a',
          fill: '#f3f0ea',
          bold: false,
          align: 'left',
          radius: 3,
        },
      ],
    },
  ];
  return { type: 'presentation', slides };
}
