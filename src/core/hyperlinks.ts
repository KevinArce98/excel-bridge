import { CellStyle, Hyperlink } from './types';
import { EXCEL_LIMITS } from './date-utils';
import { parseRange, formatRange } from './cell-ref';

export interface PreparedHyperlink {
  ref: string;
  anchor: string;
  rId?: string;
  target?: string;
  location?: string;
  tooltip?: string;
  display?: string;
}

export const HYPERLINK_STYLE: CellStyle = { color: '#0563C1', underline: true };

const ALLOWED_URL = /^(?:https?:\/\/|mailto:)[^\s"<>\\^`{|}\x00-\x1F\x7F]+$/i;
const WEB_URL = /^https?:/i;
const CONTROL_CHARS = /[\x00-\x1F\x7F]/;

const checkLength = (kind: 'url' | 'location', value: string): void => {
  if (value.length > EXCEL_LIMITS.MAX_HYPERLINK_LENGTH) {
    throw new Error(
      `Hyperlink ${kind} length ${value.length} exceeds Excel limit (${EXCEL_LIMITS.MAX_HYPERLINK_LENGTH})`
    );
  }
};

export const prepareHyperlink = (link: Hyperlink): PreparedHyperlink => {
  const range = parseRange(link.range);
  const ref = formatRange(range);
  const url = typeof link.url === 'string' ? link.url : '';
  const location = typeof link.location === 'string' ? link.location.replace(/^#/, '') : '';

  if (Boolean(url) === Boolean(location)) {
    throw new Error(`Hyperlink at ${ref} must set exactly one of url or location`);
  }

  const prepared: PreparedHyperlink = { ref, anchor: `${range.start.row}-${range.start.col}` };

  if (url) {
    checkLength('url', url);
    if (!ALLOWED_URL.test(url)) {
      throw new Error(
        `Unsupported hyperlink url at ${ref}: ${url} (use http, https or mailto; percent-encode spaces and quotes)`
      );
    }

    const hashIndex = url.indexOf('#');
    if (WEB_URL.test(url) && hashIndex !== -1 && hashIndex < url.length - 1) {
      prepared.target = url.slice(0, hashIndex);
      prepared.location = url.slice(hashIndex + 1);
    } else {
      prepared.target = url;
    }
  } else {
    checkLength('location', location);
    if (CONTROL_CHARS.test(location)) {
      throw new Error(`Invalid hyperlink location at ${ref}: ${location}`);
    }
    prepared.location = location;
  }

  if (link.tooltip !== undefined) prepared.tooltip = String(link.tooltip);
  if (link.display !== undefined) prepared.display = String(link.display);

  return prepared;
};

export const prepareHyperlinks = (links: Hyperlink[] = []): PreparedHyperlink[] => {
  if (links.length > EXCEL_LIMITS.MAX_HYPERLINKS) {
    throw new Error(
      `Hyperlink count ${links.length} exceeds Excel limit (${EXCEL_LIMITS.MAX_HYPERLINKS})`
    );
  }

  const seen = new Set<string>();
  let nextRelationshipId = 1;

  return links.map(link => {
    const prepared = prepareHyperlink(link);
    if (seen.has(prepared.ref)) {
      throw new Error(`Duplicate hyperlink at ${prepared.ref}`);
    }
    seen.add(prepared.ref);

    if (prepared.target !== undefined) {
      prepared.rId = `rId${nextRelationshipId++}`;
    }
    return prepared;
  });
};

export const withHyperlinkStyles = (
  styles: Record<string, CellStyle> | undefined,
  links: PreparedHyperlink[]
): Record<string, CellStyle> | undefined => {
  if (links.length === 0) return styles;

  const merged = { ...styles };
  for (const link of links) {
    const style = merged[link.anchor];
    merged[link.anchor] = {
      ...style,
      color: style?.color ?? HYPERLINK_STYLE.color,
      underline: style?.underline ?? HYPERLINK_STYLE.underline,
    };
  }
  return merged;
};
