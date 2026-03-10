export const XML_NS = {
  spreadsheetml: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  relationships: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  main_rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  content_types: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const CONTENT_TYPES = {
  XML: 'application/xml',
  RELATIONSHIPS: 'application/vnd.openxmlformats-package.relationships+xml',
  SPREADSHEETML: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  WORKSHEET: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  SHARED_STRINGS: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  STYLES: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
} as const;

export const RELATIONSHIP_TYPES = {
  WORKSHEET: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  SHARED_STRINGS:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  STYLES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
} as const;

export const CELL_TYPES = {
  INLINE_STRING: 'inlineStr',
  SHARED_STRING: 's',
  NUMBER: 'n',
  BOOLEAN: 'b',
  ERROR: 'e',
} as const;
