export function spreadsheetImportNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

export function recordSpreadsheetImportMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, unknown>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing diagnostics must never affect spreadsheet import.
  }
}
