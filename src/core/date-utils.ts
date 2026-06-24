/**
 * Convert JavaScript Date to Excel serial date number
 * Excel dates are the number of days since January 1, 1900
 * Note: Excel incorrectly treats 1900 as a leap year, so we account for that
 */
const MS_PER_DAY = 24 * 60 * 60 * 1000;
// December 30, 1899 in UTC. Excel is timezone-naive, so we anchor in UTC and use
// the Date's local calendar fields to avoid historical timezone-offset drift.
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

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

/**
 * Convert Excel serial date number to JavaScript Date
 */
export function excelSerialToDate(serial: number): Date {
  const utc = new Date(EXCEL_EPOCH_UTC + serial * MS_PER_DAY);

  // Rebuild as a local Date from the UTC calendar fields so the result is naive.
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

/**
 * Check if a value is a Date object
 */
export function isDate(value: unknown): value is Date {
  return value instanceof Date && !isNaN(value.getTime());
}

/** Built-in Excel number-format ids that represent dates or times. */
const BUILTIN_DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/**
 * Determine whether an Excel number-format id refers to a date/time format.
 * `customFormats` maps custom numFmtId (>= 164) to its format code string.
 */
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

/**
 * Heuristic: does a custom format code contain date/time tokens?
 * Strips quoted literals and bracketed sections first to avoid false positives.
 */
export function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(stripped);
}

/**
 * Validate Excel limits
 */
export const EXCEL_LIMITS = {
  MAX_ROWS: 1048576,
  MAX_COLS: 16384,
  MAX_CELL_LENGTH: 32767,
} as const;

export function validateRowIndex(row: number): void {
  if (row < 0 || row >= EXCEL_LIMITS.MAX_ROWS) {
    throw new Error(`Row index ${row} exceeds Excel limit (0-${EXCEL_LIMITS.MAX_ROWS - 1})`);
  }
}

export function validateColIndex(col: number): void {
  if (col < 0 || col >= EXCEL_LIMITS.MAX_COLS) {
    throw new Error(`Column index ${col} exceeds Excel limit (0-${EXCEL_LIMITS.MAX_COLS - 1})`);
  }
}

export function validateCellValue(value: string): void {
  if (value.length > EXCEL_LIMITS.MAX_CELL_LENGTH) {
    throw new Error(
      `Cell value length ${value.length} exceeds Excel limit (${EXCEL_LIMITS.MAX_CELL_LENGTH})`
    );
  }
}
