export const officeFontFamilies = [
  {
    id: 'aptos',
    name: 'Aptos',
    cssValue: 'Aptos',
    cssFamily: 'Aptos, sans-serif',
    label: 'Aptos',
  },
  {
    id: 'microsoft-yahei',
    name: 'Microsoft YaHei',
    cssValue: '"Microsoft YaHei"',
    cssFamily: '"Microsoft YaHei", "PingFang SC", sans-serif',
    label: '微软雅黑',
  },
  {
    id: 'simsun',
    name: 'SimSun',
    cssValue: 'SimSun',
    cssFamily: 'SimSun, "Songti SC", serif',
    label: '宋体',
  },
  {
    id: 'simhei',
    name: 'SimHei',
    cssValue: 'SimHei',
    cssFamily: 'SimHei, "Heiti SC", sans-serif',
    label: '黑体',
  },
  {
    id: 'kaiti',
    name: 'KaiTi',
    cssValue: 'KaiTi',
    cssFamily: 'KaiTi, "Kaiti SC", serif',
    label: '楷体',
  },
  {
    id: 'arial',
    name: 'Arial',
    cssValue: 'Arial',
    cssFamily: 'Arial, sans-serif',
    label: 'Arial',
  },
  {
    id: 'times-new-roman',
    name: 'Times New Roman',
    cssValue: '"Times New Roman"',
    cssFamily: '"Times New Roman", serif',
    label: 'Times New Roman',
  },
] as const;

export type OfficeFontFamily = (typeof officeFontFamilies)[number];

export function officeFontFamilyLabel(value: string): string {
  const firstFamily = value.split(',')[0]?.trim() || value.trim();
  return firstFamily.replace(/^(['"])(.*)\1$/, '$2');
}

export function normalizeOfficeFontFamily(value: string): string {
  return officeFontFamilyLabel(value).toLocaleLowerCase();
}
