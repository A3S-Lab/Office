export const WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS = 32_767;

export function workSpreadsheetFilterTextCharacters(value: string): number {
  let characters = 0;
  for (const _character of value) characters += 1;
  return characters;
}

export function workSpreadsheetFilterTextIsBounded(value: string): boolean {
  const characters = workSpreadsheetFilterTextCharacters(value);
  return (
    characters >= 1 && characters <= WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS
  );
}
