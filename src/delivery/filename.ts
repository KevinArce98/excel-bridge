const FORBIDDEN_IN_FILENAME = /[\\/:*?"<>|\x00-\x1F\x7F]+/g;
const NOT_PRINTABLE_ASCII = /[^\x20-\x7E]|["\\]/g;
const RESERVED_BY_RFC_5987 = /['()*]/g;

export const xlsxFilename = (filename = ''): string => {
  const stem = filename
    .replace(FORBIDDEN_IN_FILENAME, '_')
    .replace(/\.xlsx$/i, '')
    .trim();
  return `${stem || 'workbook'}.xlsx`;
};

export const attachmentHeader = (filename: string): string => {
  const encoded = encodeURIComponent(filename).replace(
    RESERVED_BY_RFC_5987,
    char => `%${char.charCodeAt(0).toString(16).toUpperCase()}`
  );
  return `attachment; filename="${filename.replace(NOT_PRINTABLE_ASCII, '_')}"; filename*=UTF-8''${encoded}`;
};
