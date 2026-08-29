export function calculateColumnWidths(data: any[][]): number[] {
  if (data.length === 0) return [];

  const maxCols = Math.max(...data.map(row => row.length));
  const widths: number[] = new Array(maxCols).fill(0);

  data.forEach(row => {
    row.forEach((cell, colIndex) => {
      const cellText = cell?.toString() || '';
      const cellWidth = estimateTextWidth(cellText);
      widths[colIndex] = Math.max(widths[colIndex], cellWidth);
    });
  });

  return widths.map(w => Math.min(Math.max(w, 8), 50));
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
