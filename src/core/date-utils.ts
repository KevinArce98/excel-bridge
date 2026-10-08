import { invalidInput } from './errors';
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAYS_BETWEEN_DATE_SYSTEMS = 1462;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?Z?)?$/;

export function dateToExcelSerial(date: Date): number {
  const utc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds()
  );

  return (utc - EXCEL_EPOCH_UTC) / MS_PER_DAY;
}

export function excelSerialToDate(serial: number, date1904 = false): Date {
  const days = date1904 ? serial + DAYS_BETWEEN_DATE_SYSTEMS : serial;
  const utc = new Date(EXCEL_EPOCH_UTC + days * MS_PER_DAY);

  return new Date(
    utc.getUTCFullYear(),
    utc.getUTCMonth(),
    utc.getUTCDate(),
    utc.getUTCHours(),
    utc.getUTCMinutes(),
    utc.getUTCSeconds(),
    utc.getUTCMilliseconds()
  );
}

export function parseIsoDate(text: string): Date | undefined {
  const match = ISO_DATE.exec(text);
  if (!match) return undefined;

  const [, year, month, day, hours = '0', minutes = '0', seconds = '0', fraction = ''] = match;
  const date = new Date(2000, 0, 1);
  date.setFullYear(Number(year), Number(month) - 1, Number(day));
  date.setHours(
    Number(hours),
    Number(minutes),
    Number(seconds),
    Number(fraction.slice(0, 3).padEnd(3, '0'))
  );

  const sameCalendarDay =
    date.getFullYear() === Number(year) &&
    date.getMonth() === Number(month) - 1 &&
    date.getDate() === Number(day);
  return sameCalendarDay && Number(hours) < 24 && Number(minutes) < 60 && Number(seconds) < 60
    ? date
    : undefined;
}

export function isDate(value: unknown): value is Date {
  return value instanceof Date && !isNaN(value.getTime());
}

const BUILTIN_DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

export function isDateNumFmtId(
  numFmtId: number,
  customFormats: Record<number, string> = {}
): boolean {
  if (BUILTIN_DATE_NUMFMT_IDS.has(numFmtId)) {
    return true;
  }
  const code = customFormats[numFmtId];
  return code ? isDateFormatCode(code) : false;
}

export function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(stripped);
}

export const EXCEL_LIMITS = {
  MAX_ROWS: 1048576,
  MAX_COLS: 16384,
  MAX_CELL_LENGTH: 32767,
  MAX_HYPERLINKS: 65530,
  MAX_HYPERLINK_LENGTH: 2079,
} as const;

export function validateRowIndex(row: number): void {
  if (row < 0 || row >= EXCEL_LIMITS.MAX_ROWS) {
    throw invalidInput(`Row index ${row} exceeds Excel limit (0-${EXCEL_LIMITS.MAX_ROWS - 1})`);
  }
}

export function validateColIndex(col: number): void {
  if (col < 0 || col >= EXCEL_LIMITS.MAX_COLS) {
    throw invalidInput(`Column index ${col} exceeds Excel limit (0-${EXCEL_LIMITS.MAX_COLS - 1})`);
  }
}

export function validateCellValue(value: string): void {
  if (value.length > EXCEL_LIMITS.MAX_CELL_LENGTH) {
    throw invalidInput(
      `Cell value length ${value.length} exceeds Excel limit (${EXCEL_LIMITS.MAX_CELL_LENGTH})`
    );
  }
}
