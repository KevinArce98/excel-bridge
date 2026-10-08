import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { XLSX_CONTENT_TYPE } from './constants';
import { ExcelBridgeError, limitExceeded } from './errors';

export interface PartBudget {
  maxPartBytes: number;
  maxTotalBytes: number;
  total: number;
}

export interface ExcelFiles {
  [path: string]: string;
}

export const createExcelBlob = (files: ExcelFiles): Blob => {
  const zipConfig: Record<string, Uint8Array> = {};

  for (const [path, content] of Object.entries(files)) {
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    zipConfig[cleanPath] = strToU8(content);
  }

  const zipped = zipSync(zipConfig, { level: 6 });
  const zippedArray = new Uint8Array(zipped);

  if (typeof Blob !== 'undefined') {
    return new Blob([zippedArray], { type: XLSX_CONTENT_TYPE });
  }

  throw new ExcelBridgeError(
    'UNSUPPORTED',
    'Blob is not available in this environment. Use createExcelBuffer() for Node.js.'
  );
};

export const createExcelBuffer = (files: ExcelFiles): Uint8Array => {
  const zipConfig: Record<string, Uint8Array> = {};

  for (const [path, content] of Object.entries(files)) {
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    zipConfig[cleanPath] = strToU8(content);
  }

  const result = zipSync(zipConfig, { level: 6 });

  return new Uint8Array(result);
};

const takeFromBudget = (budget: PartBudget, name: string, size: number): void => {
  if (size > budget.maxPartBytes) {
    throw limitExceeded(
      'maxPartBytes',
      budget.maxPartBytes,
      `Part "${name}" inflates to ${size} bytes`
    );
  }
  budget.total += size;
  if (budget.total > budget.maxTotalBytes) {
    throw limitExceeded(
      'maxTotalBytes',
      budget.maxTotalBytes,
      `The workbook inflates to ${budget.total} bytes`
    );
  }
};

export const extractParts = (
  buffer: Uint8Array,
  select?: (path: string) => boolean,
  budget?: PartBudget
): ExcelFiles => {
  try {
    const unzipped = unzipSync(buffer, {
      filter: ({ name, originalSize }) => {
        if (select && !select(name)) return false;
        if (budget) takeFromBudget(budget, name, originalSize);
        return true;
      },
    });
    const files: ExcelFiles = {};

    for (const [path, content] of Object.entries(unzipped)) {
      files[path] = strFromU8(content);
    }

    return files;
  } catch (cause) {
    throw cause instanceof ExcelBridgeError
      ? cause
      : new ExcelBridgeError('INVALID_FILE', 'Invalid Excel file: Unable to extract ZIP contents', {
          cause,
        });
  }
};

export const extractExcelFiles = (buffer: Uint8Array): ExcelFiles => extractParts(buffer);

export const validateExcelStructure = (files: ExcelFiles): boolean => {
  const requiredFiles = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'];

  const hasRequired = requiredFiles.every(file => files[file]);
  const hasWorksheet = Object.keys(files).some(path => /^xl\/worksheets\/.+\.xml$/.test(path));

  return hasRequired && hasWorksheet;
};
