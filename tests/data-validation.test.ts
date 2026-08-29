import { describe, it, expect } from 'vitest';
import { ExcelWriter, dataValidation } from '../src';
import { unzipSync, strFromU8 } from 'fflate';

const sheetXml = (buffer: Uint8Array): string => {
  const files = unzipSync(buffer);
  return strFromU8(files['xl/worksheets/sheet1.xml']);
};

const write = (validations: ReturnType<typeof dataValidation.list>[]) =>
  sheetXml(new ExcelWriter().createWorkbookBuffer([{ data: [['H']], validations }]));

describe('dataValidation builders', () => {
  it('builds a list (dropdown) validation', () => {
    const v = dataValidation.list('A2:A10', ['Yes', 'No', 'Maybe']);
    expect(v).toEqual({ range: 'A2:A10', type: 'list', options: 'Yes,No,Maybe' });

    const xml = write([v]);
    expect(xml).toContain('type="list"');
    expect(xml).toContain('sqref="A2:A10"');
    expect(xml).toContain('<formula1>"Yes,No,Maybe"</formula1>');
  });

  it('builds a whole-number between validation', () => {
    const v = dataValidation.wholeNumber('B2:B10', 'between', 1, 100);
    expect(v).toMatchObject({ type: 'whole', operator: 'between', formula1: '1', formula2: '100' });

    const xml = write([v]);
    expect(xml).toContain('type="whole"');
    expect(xml).toContain('operator="between"');
    expect(xml).toContain('<formula1>1</formula1>');
    expect(xml).toContain('<formula2>100</formula2>');
  });

  it('builds a decimal validation with a single bound', () => {
    const v = dataValidation.decimal('C2:C10', 'greaterThan', 0);
    expect(v).toMatchObject({ type: 'decimal', operator: 'greaterThan', formula1: '0' });
    expect(v.formula2).toBeUndefined();

    const xml = write([v]);
    expect(xml).toContain('type="decimal"');
    expect(xml).toContain('operator="greaterThan"');
    expect(xml).toContain('<formula1>0</formula1>');
    expect(xml).not.toContain('<formula2>');
  });

  it('builds a text-length validation', () => {
    const v = dataValidation.textLength('D2:D10', 'lessThanOrEqual', 50);
    const xml = write([v]);
    expect(xml).toContain('type="textLength"');
    expect(xml).toContain('operator="lessThanOrEqual"');
    expect(xml).toContain('<formula1>50</formula1>');
  });

  it('builds a date-between validation using Excel serials', () => {
    const v = dataValidation.dateBetween('E2:E10', new Date(2024, 0, 1), new Date(2024, 11, 31));
    expect(v).toMatchObject({ type: 'date', operator: 'between', formula1: '45292', formula2: '45657' });

    const xml = write([v]);
    expect(xml).toContain('type="date"');
    expect(xml).toContain('<formula1>45292</formula1>');
    expect(xml).toContain('<formula2>45657</formula2>');
  });

  it('keeps the legacy { range, options } list shape working', () => {
    const xml = write([{ range: 'A2:A5', options: 'Low,High' }]);
    expect(xml).toContain('type="list"');
    expect(xml).toContain('<formula1>"Low,High"</formula1>');
  });
});
