import { describe, it, expect } from 'vitest';
import { ExcelBridge } from '../src';

describe('Excel Numbers Only', () => {
  it('should create valid Excel file with only numbers', () => {
    const data = [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    expect(result.sheets).toHaveLength(1);
    const sheet = result.sheets[0];

    expect(sheet.data[0][0].value).toBe(1);
    expect(sheet.data[0][0].type).toBe('number');
    expect(sheet.data[1][1].value).toBe(5);
    expect(sheet.data[2][2].value).toBe(9);
  });

  it('should create valid Excel file with mixed numbers and booleans', () => {
    const data = [
      [100, true],
      [200, false],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];

    expect(sheet.data[0][0].value).toBe(100);
    expect(sheet.data[0][0].type).toBe('number');
    expect(sheet.data[0][1].value).toBe(true);
    expect(sheet.data[0][1].type).toBe('boolean');
  });
});
