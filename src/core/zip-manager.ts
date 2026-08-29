import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';

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
    return new Blob([zippedArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  throw new Error(
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
  const requiredFiles = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'];

  const hasRequired = requiredFiles.every(file => files[file]);
  const hasWorksheet = Object.keys(files).some(path => /^xl\/worksheets\/.+\.xml$/.test(path));

  return hasRequired && hasWorksheet;
};
