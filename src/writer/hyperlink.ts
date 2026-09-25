import { ExternalHyperlink, InternalHyperlink } from '../core/types';
import { parseRange, formatRange, quoteSheetName } from '../core/cell-ref';

export type HyperlinkOptions = Pick<ExternalHyperlink, 'tooltip' | 'display'>;

export const hyperlink = {
  url(range: string, url: string, options: HyperlinkOptions = {}): ExternalHyperlink {
    return { range, url, ...options };
  },

  email(
    range: string,
    address: string,
    options: HyperlinkOptions & { subject?: string } = {}
  ): ExternalHyperlink {
    const { subject, ...rest } = options;
    const query = subject !== undefined ? `?subject=${encodeURIComponent(subject)}` : '';
    return { range, url: `mailto:${address}${query}`, ...rest };
  },

  internal(
    range: string,
    sheet: string,
    cell = 'A1',
    options: HyperlinkOptions = {}
  ): InternalHyperlink {
    return {
      range,
      location: `${quoteSheetName(sheet)}!${formatRange(parseRange(cell))}`,
      ...options,
    };
  },
};
