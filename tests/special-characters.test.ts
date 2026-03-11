import { describe, it, expect } from 'vitest';
import { ExcelBridge } from '../src';

describe('Special Characters Handling', () => {
  it('should handle XML special characters correctly', () => {
    const data = [
      ['Company', 'Description'],
      ['R&D Corp', 'Research & Development'],
      ['Tech <Solutions>', 'Software & Hardware'],
      ['Data "Analytics"', "It's great!"],
      ['Ampersand & Co', 'Less <than> Greater >'],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];
    
    // Verify all special characters are preserved
    expect(sheet.data[1][0].value).toBe('R&D Corp');
    expect(sheet.data[1][1].value).toBe('Research & Development');
    expect(sheet.data[2][0].value).toBe('Tech <Solutions>');
    expect(sheet.data[2][1].value).toBe('Software & Hardware');
    expect(sheet.data[3][0].value).toBe('Data "Analytics"');
    expect(sheet.data[3][1].value).toBe("It's great!");
    expect(sheet.data[4][0].value).toBe('Ampersand & Co');
    expect(sheet.data[4][1].value).toBe('Less <than> Greater >');
  });

  it('should handle Unicode characters (tildes, ñ, etc.)', () => {
    const data = [
      ['Nombre', 'Ciudad', 'País'],
      ['José', 'México', 'México'],
      ['María', 'São Paulo', 'Brasil'],
      ['François', 'Montréal', 'Canadá'],
      ['Müller', 'München', 'Alemania'],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];
    
    expect(sheet.data[1][0].value).toBe('José');
    expect(sheet.data[1][1].value).toBe('México');
    expect(sheet.data[2][0].value).toBe('María');
    expect(sheet.data[2][1].value).toBe('São Paulo');
    expect(sheet.data[3][0].value).toBe('François');
    expect(sheet.data[3][1].value).toBe('Montréal');
    expect(sheet.data[4][0].value).toBe('Müller');
    expect(sheet.data[4][1].value).toBe('München');
  });

  it('should handle whitespace and special formatting', () => {
    const data = [
      ['Name', 'Value'],
      ['Normal', 'Normal text'],
      ['With Space', 'Text with spaces'],
      ['TabChar', 'Contains\ttab'],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];
    
    expect(sheet.data[1][0].value).toBe('Normal');
    expect(sheet.data[1][1].value).toBe('Normal text');
    expect(sheet.data[2][0].value).toBe('With Space');
    expect(sheet.data[3][0].value).toBe('TabChar');
  });

  it('should handle mixed content with numbers and special chars', () => {
    const data = [
      ['Product', 'Price', 'Description'],
      ['Widget & Co', 29.99, 'Best <product> ever!'],
      ['Gadget "Pro"', 149.50, 'It\'s amazing & affordable'],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const result = ExcelBridge.read(buffer);

    const sheet = result.sheets[0];
    
    expect(sheet.data[1][0].value).toBe('Widget & Co');
    expect(sheet.data[1][1].value).toBe(29.99);
    expect(sheet.data[1][1].type).toBe('number');
    expect(sheet.data[1][2].value).toBe('Best <product> ever!');
    
    expect(sheet.data[2][0].value).toBe('Gadget "Pro"');
    expect(sheet.data[2][1].value).toBe(149.50);
    expect(sheet.data[2][2].value).toBe("It's amazing & affordable");
  });
});
