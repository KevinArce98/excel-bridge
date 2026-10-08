const FORBIDDEN_IN_FILENAME =
  /(?:[\\/:*?"<>|\x00-\x1F\x7F\u200E\u200F\u202A-\u202E\u2066-\u2069]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF])+/g;
const NOT_PRINTABLE_ASCII = /[^\x20-\x7E]|["\\]/g;
const RESERVED_BY_RFC_5987 = /['()*]/g;

export const xlsxFilename = (filename = ''): string => {
  const stem = filename
    .trim()
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
