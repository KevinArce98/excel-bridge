import { describe, it, expect } from 'vitest';
import { ExcelBridge } from '../src';

describe('Excel Read/Write Integration', () => {
  it('should write and read back the same data', () => {
    const originalData = [
      ['Name', 'Age', 'Active'],
      ['John', 25, true],
      ['Jane', 30, false],
      ['Bob', 35, true],
    ];

    const buffer = ExcelBridge.writeBuffer(originalData);
    const result = ExcelBridge.read(buffer);

    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].name).toBe('Sheet1');

    const sheet = result.sheets[0];
    expect(sheet.data).toHaveLength(4);

    expect(sheet.data[0][0].value).toBe('Name');
    expect(sheet.data[0][0].type).toBe('string');
    expect(sheet.data[0][1].value).toBe('Age');
    expect(sheet.data[0][2].value).toBe('Active');

    expect(sheet.data[1][0].value).toBe('John');
    expect(sheet.data[1][1].value).toBe(25);
    expect(sheet.data[1][1].type).toBe('number');
    expect(sheet.data[1][2].value).toBe(true);
    expect(sheet.data[1][2].type).toBe('boolean');

    expect(sheet.data[2][0].value).toBe('Jane');
    expect(sheet.data[2][1].value).toBe(30);
    expect(sheet.data[2][2].value).toBe(false);
  });

  it('should handle mixed data types correctly', () => {
    const data = [
      ['String', 'Number', 'Boolean', 'Empty'],
      ['Hello', 123, true, null],
      ['World', 456.78, false, undefined],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];

    expect(sheet.data[1][0].value).toBe('Hello');
    expect(sheet.data[1][0].type).toBe('string');
    expect(sheet.data[1][1].value).toBe(123);
    expect(sheet.data[1][1].type).toBe('number');
    expect(sheet.data[1][2].value).toBe(true);
    expect(sheet.data[1][2].type).toBe('boolean');

    expect(sheet.data[2][1].value).toBe(456.78);
    expect(sheet.data[2][1].type).toBe('number');
  });
});
