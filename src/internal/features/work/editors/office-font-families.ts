import type { OfficeMessageCatalog, OfficeMessageKey } from '../../../i18n/office-messages';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';

export type OfficeFontFamilyGroupId = 'cjk' | 'latin' | 'mono';

export interface OfficeFontFamily {
  id: string;
  name: string;
  cssValue: string;
  cssFamily: string;
  group: OfficeFontFamilyGroupId;
  labelKey?: OfficeMessageKey;
}

const FONT_FAMILY_LABEL_KEYS: Record<string, OfficeMessageKey> = {
  'microsoft-yahei': 'document.font.family.microsoftYahei',
  'pingfang-sc': 'document.font.family.pingfangSc',
  simsun: 'document.font.family.simsun',
  stsong: 'document.font.family.stsong',
  simhei: 'document.font.family.simhei',
  stheiti: 'document.font.family.stheiti',
  kaiti: 'document.font.family.kaiti',
  fangsong: 'document.font.family.fangsong',
  'hiragino-sans-gb': 'document.font.family.hiraginoSansGb',
};

export const officeFontFamilies = [
  {
    id: 'microsoft-yahei',
    name: 'Microsoft YaHei',
    cssValue: '"Microsoft YaHei"',
    cssFamily: '"Microsoft YaHei", "PingFang SC", sans-serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS['microsoft-yahei'],
  },
  {
    id: 'pingfang-sc',
    name: 'PingFang SC',
    cssValue: '"PingFang SC"',
    cssFamily: '"PingFang SC", "Microsoft YaHei", sans-serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS['pingfang-sc'],
  },
  {
    id: 'simsun',
    name: 'SimSun',
    cssValue: 'SimSun',
    cssFamily: 'SimSun, "Songti SC", serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.simsun,
  },
  {
    id: 'stsong',
    name: 'STSong',
    cssValue: 'STSong',
    cssFamily: 'STSong, "Songti SC", SimSun, serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.stsong,
  },
  {
    id: 'simhei',
    name: 'SimHei',
    cssValue: 'SimHei',
    cssFamily: 'SimHei, "Heiti SC", sans-serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.simhei,
  },
  {
    id: 'stheiti',
    name: 'STHeiti',
    cssValue: 'STHeiti',
    cssFamily: 'STHeiti, "Heiti SC", SimHei, sans-serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.stheiti,
  },
  {
    id: 'kaiti',
    name: 'KaiTi',
    cssValue: 'KaiTi',
    cssFamily: 'KaiTi, "Kaiti SC", serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.kaiti,
  },
  {
    id: 'fangsong',
    name: 'FangSong',
    cssValue: 'FangSong',
    cssFamily: 'FangSong, STFangsong, serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS.fangsong,
  },
  {
    id: 'hiragino-sans-gb',
    name: 'Hiragino Sans GB',
    cssValue: '"Hiragino Sans GB"',
    cssFamily: '"Hiragino Sans GB", "PingFang SC", sans-serif',
    group: 'cjk',
    labelKey: FONT_FAMILY_LABEL_KEYS['hiragino-sans-gb'],
  },
  {
    id: 'aptos',
    name: 'Aptos',
    cssValue: 'Aptos',
    cssFamily: 'Aptos, sans-serif',
    group: 'latin',
  },
  {
    id: 'calibri',
    name: 'Calibri',
    cssValue: 'Calibri',
    cssFamily: 'Calibri',
    group: 'latin',
  },
  {
    id: 'arial',
    name: 'Arial',
    cssValue: 'Arial',
    cssFamily: 'Arial, sans-serif',
    group: 'latin',
  },
  {
    id: 'helvetica',
    name: 'Helvetica',
    cssValue: 'Helvetica',
    cssFamily: 'Helvetica, Arial, sans-serif',
    group: 'latin',
  },
  {
    id: 'times-new-roman',
    name: 'Times New Roman',
    cssValue: '"Times New Roman"',
    cssFamily: '"Times New Roman", Times, serif',
    group: 'latin',
  },
  {
    id: 'georgia',
    name: 'Georgia',
    cssValue: 'Georgia',
    cssFamily: 'Georgia, "Times New Roman", serif',
    group: 'latin',
  },
  {
    id: 'cambria',
    name: 'Cambria',
    cssValue: 'Cambria',
    cssFamily: 'Cambria, Georgia, serif',
    group: 'latin',
  },
  {
    id: 'garamond',
    name: 'Garamond',
    cssValue: 'Garamond',
    cssFamily: 'Garamond, Georgia, serif',
    group: 'latin',
  },
  {
    id: 'verdana',
    name: 'Verdana',
    cssValue: 'Verdana',
    cssFamily: 'Verdana, Arial, sans-serif',
    group: 'latin',
  },
  {
    id: 'tahoma',
    name: 'Tahoma',
    cssValue: 'Tahoma',
    cssFamily: 'Tahoma, Verdana, sans-serif',
    group: 'latin',
  },
  {
    id: 'trebuchet-ms',
    name: 'Trebuchet MS',
    cssValue: '"Trebuchet MS"',
    cssFamily: '"Trebuchet MS", Arial, sans-serif',
    group: 'latin',
  },
  {
    id: 'sf-mono',
    name: 'SFMono-Regular',
    cssValue: 'SFMono-Regular',
    cssFamily: 'SFMono-Regular, Menlo, Consolas, monospace',
    group: 'mono',
  },
  {
    id: 'menlo',
    name: 'Menlo',
    cssValue: 'Menlo',
    cssFamily: 'Menlo, SFMono-Regular, Consolas, monospace',
    group: 'mono',
  },
  {
    id: 'consolas',
    name: 'Consolas',
    cssValue: 'Consolas',
    cssFamily: 'Consolas, "Courier New", monospace',
    group: 'mono',
  },
  {
    id: 'courier-new',
    name: 'Courier New',
    cssValue: '"Courier New"',
    cssFamily: '"Courier New", Courier, monospace',
    group: 'mono',
  },
  {
    id: 'monaco',
    name: 'Monaco',
    cssValue: 'Monaco',
    cssFamily: 'Monaco, Menlo, Consolas, monospace',
    group: 'mono',
  },
] as const satisfies readonly OfficeFontFamily[];

export function officeFontFamilyGroupLabel(
  group: OfficeFontFamilyGroupId,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (group === 'cjk') {
    return officeMessage(messages, 'document.font.familyGroup.cjk');
  }
  if (group === 'mono') {
    return officeMessage(messages, 'document.font.familyGroup.mono');
  }
  return officeMessage(messages, 'document.font.familyGroup.latin');
}

export function officeFontFamilyLocalizedLabel(
  family: Pick<OfficeFontFamily, 'labelKey' | 'name'>,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  return family.labelKey
    ? officeMessage(messages, family.labelKey)
    : family.name;
}

export function officeFontFamilyLabel(value: string): string {
  const firstFamily = value.split(',')[0]?.trim() || value.trim();
  return firstFamily.replace(/^(['"])(.*)\1$/, '$2');
}

export function normalizeOfficeFontFamily(value: string): string {
  return officeFontFamilyLabel(value).toLocaleLowerCase();
}
