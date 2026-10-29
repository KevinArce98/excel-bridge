import { rowIndexes } from './rows';
import { splitCell } from './cells';
import { CellValue } from './types';

const displayedText = (cell: CellValue): string => {
  const { formula, value } = splitCell(cell);
  const shown = value ?? (formula === undefined ? undefined : `=${formula}`);
  return (typeof shown === 'object' && 'error' in shown ? shown.error : shown)?.toString() || '';
};

export function calculateColumnWidths(data: CellValue[][]): number[] {
  const widths: number[] = [];
  let maxCols = 0;

  rowIndexes(data).forEach(index => {
    const row = data[index];
    maxCols = Math.max(maxCols, row.length);
    row.forEach((cell, colIndex) => {
      const cellText = displayedText(cell);
      widths[colIndex] = Math.max(widths[colIndex] ?? 0, estimateTextWidth(cellText));
    });
  });

  return Array.from({ length: maxCols }, (_, colIndex) =>
    Math.min(Math.max(widths[colIndex] ?? 0, 8), 50)
  );
}

function estimateTextWidth(text: string): number {
  if (!text) return 8;

  let width = text.length * 1.2;

  const wideChars = text.match(/[WMm@]/g);
  if (wideChars) {
    width += wideChars.length * 0.5;
  }

  width += 2;

  return Math.ceil(width);
}

const DEFAULT_COLUMN_WIDTH = 9.140625;

export function generateColsXml(
  widths: number[],
  { hiddenColumns = [] }: { hiddenColumns?: Iterable<number> } = {}
): string {
  const hidden = new Set(hiddenColumns);
  const columns = widths.slice();
  hidden.forEach(index => {
    columns[index] ??= DEFAULT_COLUMN_WIDTH;
  });

  if (columns.length === 0) return '';

  const colsXml = columns
    .map((width, index) => {
      const colNum = index + 1;
      const hiddenAttr = hidden.has(index) ? ' hidden="1"' : '';
      return `    <col min="${colNum}" max="${colNum}" width="${width}" customWidth="1"${hiddenAttr}/>`;
    })
    .join('\n');

  return `  <cols>\n${colsXml}\n  </cols>`;
}
