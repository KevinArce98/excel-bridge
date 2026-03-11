/**
 * Calculate the maximum width needed for each column based on cell content
 */
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

  // Apply min/max constraints
  return widths.map(w => Math.min(Math.max(w, 8), 50));
}

/**
 * Estimate text width in Excel units
 * Excel width units are approximately the width of '0' in the default font
 */
function estimateTextWidth(text: string): number {
  if (!text) return 8;

  // Base calculation: character count + padding
  let width = text.length * 1.2;

  // Add extra width for wide characters
  const wideChars = text.match(/[WMm@]/g);
  if (wideChars) {
    width += wideChars.length * 0.5;
  }

  // Add padding
  width += 2;

  return Math.ceil(width);
}

/**
 * Generate <cols> XML for column widths
 */
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
