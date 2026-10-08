export type ExcelBridgeErrorCode =
  'INVALID_INPUT' | 'INVALID_FILE' | 'LIMIT_EXCEEDED' | 'UNSUPPORTED';

export type ReaderLimitName = 'maxCells' | 'maxPartBytes' | 'maxTotalBytes' | 'maxSheets';

export interface ExcelBridgeErrorOptions {
  cause?: unknown;
  limit?: ReaderLimitName;
}

export class ExcelBridgeError extends Error {
  readonly code: ExcelBridgeErrorCode;
  declare readonly limit?: ReaderLimitName;

  constructor(code: ExcelBridgeErrorCode, message: string, options?: ExcelBridgeErrorOptions) {
    super(message, options);
    this.name = 'ExcelBridgeError';
    this.code = code;
    if (options?.limit) this.limit = options.limit;
  }
}

export const isExcelBridgeError = (error: unknown): error is ExcelBridgeError =>
  error instanceof Error && error.name === 'ExcelBridgeError';

export const invalidInput = (message: string): ExcelBridgeError =>
  new ExcelBridgeError('INVALID_INPUT', message);

export const limitExceeded = (
  limit: ReaderLimitName,
  max: number,
  what: string
): ExcelBridgeError =>
  new ExcelBridgeError('LIMIT_EXCEEDED', `${what}, over the limit ${limit} of ${max}`, { limit });
