/**
 * Convert JavaScript Date to Excel serial date number
 * Excel dates are the number of days since January 1, 1900
 * Note: Excel incorrectly treats 1900 as a leap year, so we account for that
 */
export function dateToExcelSerial(date: Date): number {
  const epoch = new Date(1899, 11, 30); // December 30, 1899 (Excel's epoch)
  const msPerDay = 24 * 60 * 60 * 1000;

  const diff = date.getTime() - epoch.getTime();
  const days = diff / msPerDay;

  return days;
}

/**
 * Convert Excel serial date number to JavaScript Date
 */
export function excelSerialToDate(serial: number): Date {
  const epoch = new Date(1899, 11, 30);
  const msPerDay = 24 * 60 * 60 * 1000;

  return new Date(epoch.getTime() + serial * msPerDay);
}

/**
 * Check if a value is a Date object
 */
export function isDate(value: any): value is Date {
  return value instanceof Date && !isNaN(value.getTime());
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
