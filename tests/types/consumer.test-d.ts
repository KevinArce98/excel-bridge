import { describe, expectTypeOf, it } from 'vitest';
import { ExcelWriter, StyleManager } from '../../src';
import type { CellStyle, CellValue } from '../../src';

describe('1.5 consumer code keeps compiling', () => {
  it('accepts literal rows and style records', () => {
    const rows = [['a', 1, true, new Date(), null]];
    const styles: Record<string, CellStyle> = { '0-0': { bold: true, border: true } };
    new ExcelWriter().createWorkbookBuffer([{ data: rows, styles }]);
    expectTypeOf(new StyleManager().getStyleId({ bold: true })).toBeNumber();
  });

  it('keeps the input widening one-way', () => {
    expectTypeOf<string>().toMatchTypeOf<CellValue>();
    // @ts-expect-error an object without formula, text or error is not a cell
    const bad: CellValue = { result: 1 };
    void bad;
  });
});
