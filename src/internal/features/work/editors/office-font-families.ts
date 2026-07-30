export type OfficeFontFamilyGroup = '中文字体' | '西文字体' | '等宽字体';

export interface OfficeFontFamily {
  id: string;
  name: string;
  cssValue: string;
  cssFamily: string;
  label: string;
  group: OfficeFontFamilyGroup;
}

export const officeFontFamilies = [
  {
    id: 'microsoft-yahei',
    name: 'Microsoft YaHei',
    cssValue: '"Microsoft YaHei"',
    cssFamily: '"Microsoft YaHei", "PingFang SC", sans-serif',
    label: '微软雅黑',
    group: '中文字体',
  },
  {
    id: 'pingfang-sc',
    name: 'PingFang SC',
    cssValue: '"PingFang SC"',
    cssFamily: '"PingFang SC", "Microsoft YaHei", sans-serif',
    label: '苹方',
    group: '中文字体',
  },
  {
    id: 'simsun',
    name: 'SimSun',
    cssValue: 'SimSun',
    cssFamily: 'SimSun, "Songti SC", serif',
    label: '宋体',
    group: '中文字体',
  },
  {
    id: 'stsong',
    name: 'STSong',
    cssValue: 'STSong',
    cssFamily: 'STSong, "Songti SC", SimSun, serif',
    label: '华文宋体',
    group: '中文字体',
  },
  {
    id: 'simhei',
    name: 'SimHei',
    cssValue: 'SimHei',
    cssFamily: 'SimHei, "Heiti SC", sans-serif',
    label: '黑体',
    group: '中文字体',
  },
  {
    id: 'stheiti',
    name: 'STHeiti',
    cssValue: 'STHeiti',
    cssFamily: 'STHeiti, "Heiti SC", SimHei, sans-serif',
    label: '华文黑体',
    group: '中文字体',
  },
  {
    id: 'kaiti',
    name: 'KaiTi',
    cssValue: 'KaiTi',
    cssFamily: 'KaiTi, "Kaiti SC", serif',
    label: '楷体',
    group: '中文字体',
  },
  {
    id: 'fangsong',
    name: 'FangSong',
    cssValue: 'FangSong',
    cssFamily: 'FangSong, STFangsong, serif',
    label: '仿宋',
    group: '中文字体',
  },
  {
    id: 'hiragino-sans-gb',
    name: 'Hiragino Sans GB',
    cssValue: '"Hiragino Sans GB"',
    cssFamily: '"Hiragino Sans GB", "PingFang SC", sans-serif',
    label: '冬青黑体',
    group: '中文字体',
  },
  {
    id: 'aptos',
    name: 'Aptos',
    cssValue: 'Aptos',
    cssFamily: 'Aptos, sans-serif',
    label: 'Aptos',
    group: '西文字体',
  },
  {
    id: 'calibri',
    name: 'Calibri',
    cssValue: 'Calibri',
    cssFamily: 'Calibri',
    label: 'Calibri',
    group: '西文字体',
  },
  {
    id: 'arial',
    name: 'Arial',
    cssValue: 'Arial',
    cssFamily: 'Arial, sans-serif',
    label: 'Arial',
    group: '西文字体',
  },
  {
    id: 'helvetica',
    name: 'Helvetica',
    cssValue: 'Helvetica',
    cssFamily: 'Helvetica, Arial, sans-serif',
    label: 'Helvetica',
    group: '西文字体',
  },
  {
    id: 'times-new-roman',
    name: 'Times New Roman',
    cssValue: '"Times New Roman"',
    cssFamily: '"Times New Roman", Times, serif',
    label: 'Times New Roman',
    group: '西文字体',
  },
  {
    id: 'georgia',
    name: 'Georgia',
    cssValue: 'Georgia',
    cssFamily: 'Georgia, "Times New Roman", serif',
    label: 'Georgia',
    group: '西文字体',
  },
  {
    id: 'cambria',
    name: 'Cambria',
    cssValue: 'Cambria',
    cssFamily: 'Cambria, Georgia, serif',
    label: 'Cambria',
    group: '西文字体',
  },
  {
    id: 'garamond',
    name: 'Garamond',
    cssValue: 'Garamond',
    cssFamily: 'Garamond, Georgia, serif',
    label: 'Garamond',
    group: '西文字体',
  },
  {
    id: 'verdana',
    name: 'Verdana',
    cssValue: 'Verdana',
    cssFamily: 'Verdana, Arial, sans-serif',
    label: 'Verdana',
    group: '西文字体',
  },
  {
    id: 'tahoma',
    name: 'Tahoma',
    cssValue: 'Tahoma',
    cssFamily: 'Tahoma, Verdana, sans-serif',
    label: 'Tahoma',
    group: '西文字体',
  },
  {
    id: 'trebuchet-ms',
    name: 'Trebuchet MS',
    cssValue: '"Trebuchet MS"',
    cssFamily: '"Trebuchet MS", Arial, sans-serif',
    label: 'Trebuchet MS',
    group: '西文字体',
  },
  {
    id: 'sf-mono',
    name: 'SFMono-Regular',
    cssValue: 'SFMono-Regular',
    cssFamily: 'SFMono-Regular, Menlo, Consolas, monospace',
    label: 'SF Mono',
    group: '等宽字体',
  },
  {
    id: 'menlo',
    name: 'Menlo',
    cssValue: 'Menlo',
    cssFamily: 'Menlo, SFMono-Regular, Consolas, monospace',
    label: 'Menlo',
    group: '等宽字体',
  },
  {
    id: 'consolas',
    name: 'Consolas',
    cssValue: 'Consolas',
    cssFamily: 'Consolas, "Courier New", monospace',
    label: 'Consolas',
    group: '等宽字体',
  },
  {
    id: 'courier-new',
    name: 'Courier New',
    cssValue: '"Courier New"',
    cssFamily: '"Courier New", Courier, monospace',
    label: 'Courier New',
    group: '等宽字体',
  },
  {
    id: 'monaco',
    name: 'Monaco',
    cssValue: 'Monaco',
    cssFamily: 'Monaco, Menlo, Consolas, monospace',
    label: 'Monaco',
    group: '等宽字体',
  },
] as const satisfies readonly OfficeFontFamily[];

export function officeFontFamilyLabel(value: string): string {
  const firstFamily = value.split(',')[0]?.trim() || value.trim();
  return firstFamily.replace(/^(['"])(.*)\1$/, '$2');
}

export function normalizeOfficeFontFamily(value: string): string {
  return officeFontFamilyLabel(value).toLocaleLowerCase();
}
