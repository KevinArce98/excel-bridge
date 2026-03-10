import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';

export interface ExcelFiles {
  [path: string]: string;
}

export const createExcelBlob = (files: ExcelFiles): Blob => {
  const zipConfig: Record<string, Uint8Array> = {};

  for (const [path, content] of Object.entries(files)) {
    zipConfig[path] = strToU8(content);
  }

  const zipped = zipSync(zipConfig);
  const zippedArray = new Uint8Array(zipped);

  // Check if Blob is available (browser environment)
  if (typeof Blob !== 'undefined') {
    return new Blob([zippedArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  // Fallback for Node.js - create a Buffer-like object
  // Note: In Node.js, users should use createExcelBuffer instead
  throw new Error(
    'Blob is not available in this environment. Use createExcelBuffer() for Node.js.'
  );
};

export const createExcelBuffer = (files: ExcelFiles): Uint8Array => {
  const zipConfig: Record<string, Uint8Array> = {};

  for (const [path, content] of Object.entries(files)) {
    zipConfig[path] = strToU8(content);
  }

  return zipSync(zipConfig);
};

export const extractExcelFiles = (buffer: Uint8Array): ExcelFiles => {
  try {
    const unzipped = unzipSync(buffer);
    const files: ExcelFiles = {};

    for (const [path, content] of Object.entries(unzipped)) {
      files[path] = strFromU8(content);
    }

    return files;
  } catch {
    throw new Error('Invalid Excel file: Unable to extract ZIP contents');
  }
};

export const validateExcelStructure = (files: ExcelFiles): boolean => {
  const requiredFiles = [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
  ];

  return requiredFiles.every(file => files[file]);
};
