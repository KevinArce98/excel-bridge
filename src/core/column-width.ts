import { rowIndexes } from './rows';

export function calculateColumnWidths(data: any[][]): number[] {
  const widths: number[] = [];
  let maxCols = 0;

  rowIndexes(data).forEach(index => {
    const row = data[index];
    maxCols = Math.max(maxCols, row.length);
    row.forEach((cell, colIndex) => {
      const cellText = cell?.toString() || '';
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

export function generateColsXml(widths: number[]): string {
  if (widths.length === 0) return '';

  const colsXml = widths
    .map((width, index) => {
      const colNum = index + 1;
      return `    <col min="${colNum}" max="${colNum}" width="${width}" customWidth="1"/>`;
    })
    .join('\n');

  return `  <cols>\n${colsXml}\n  </cols>`;
}
