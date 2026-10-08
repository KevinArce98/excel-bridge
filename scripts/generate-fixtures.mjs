process.env.TZ = 'UTC';

import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'tests', 'fixtures');
const workDir = mkdtempSync(join(tmpdir(), 'excel-bridge-fixtures-'));

const LEAP_DAY = new Date(Date.UTC(2024, 1, 29));
const FIXED_TIME = new Date(Date.UTC(2024, 0, 1));
const YELLOW = 'FFFF00';

const buildExceljsReport = async ExcelJS => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'excel-bridge fixtures';
  workbook.created = FIXED_TIME;
  workbook.modified = FIXED_TIME;
  const report = workbook.addWorksheet('Report', { views: [{ state: 'frozen', ySplit: 1 }] });
  report.getColumn(1).width = 14;
  report.getColumn(2).width = 12;
  report.getColumn(3).width = 12;
  for (const [ref, text] of [
    ['A1', 'Region'],
    ['B1', 'Revenue'],
    ['C1', 'Date'],
  ]) {
    report.getCell(ref).value = text;
    report.getCell(ref).font = { bold: true };
    report.getCell(ref).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${YELLOW}` },
    };
  }
  report.getCell('A3').value = 'North';
  report.getCell('B3').value = 1200.5;
  report.getCell('B3').numFmt = '0.00';
  report.getCell('B3').dataValidation = {
    type: 'whole',
    operator: 'between',
    allowBlank: true,
    formulae: [1, 10000],
  };
  report.getCell('C3').value = LEAP_DAY;
  report.getCell('C3').numFmt = 'yyyy-mm-dd';
  report.getCell('A5').value = 'Total';
  report.getCell('A5').font = { bold: true };
  report.getCell('B5').value = { formula: 'SUM(B1:B4)', result: 1200.5 };
  report.getCell('A7').value = 'Footnote';
  report.mergeCells('A7:C7');
  workbook.addWorksheet('Hidden', { state: 'hidden' }).getCell('A1').value = 'secret';
  return Buffer.from(await workbook.xlsx.writeBuffer());
};

const buildExceljsText = async ExcelJS => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'excel-bridge fixtures';
  workbook.created = FIXED_TIME;
  workbook.modified = FIXED_TIME;
  const sheet = workbook.addWorksheet('Text');
  const values = [
    {
      richText: [{ text: 'bold ', font: { bold: true } }, { text: ' plain' }],
    },
    '  padded on both sides  ',
    `& < > " '`,
    'line one\nline two',
    '😀 日本語 مرحبا',
    '00123',
    'x'.repeat(300),
    '&#233; &amp;',
  ];
  values.forEach((value, index) => {
    sheet.getCell(`A${index + 1}`).value = value;
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
};

const buildSheetjsReport = XLSX => {
  const sheet = {
    A1: { t: 's', v: 'Region' },
    B1: { t: 's', v: 'Revenue' },
    C1: { t: 's', v: 'Date' },
    A3: { t: 's', v: 'North' },
    B3: { t: 'n', v: 1200.5, z: '0.00' },
    C3: { t: 'd', v: LEAP_DAY, z: 'yyyy-mm-dd' },
    A5: { t: 's', v: 'Total' },
    B5: { t: 'n', f: 'SUM(B1:B4)', v: 1200.5 },
    A7: { t: 's', v: 'Footnote' },
    '!ref': 'A1:C7',
    '!merges': [XLSX.utils.decode_range('A7:C7')],
    '!cols': [{ wch: 14 }, { wch: 12 }, { wch: 12 }],
  };
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Report');
  XLSX.utils.book_append_sheet(workbook, { A1: { t: 's', v: 'secret' }, '!ref': 'A1' }, 'Hidden');
  workbook.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', cellDates: true });
};

const buildHucreReport = writeXlsx => {
  const bold = { font: { bold: true } };
  const header = {
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: YELLOW } },
  };
  return writeXlsx({
    sheets: [
      {
        name: 'Report',
        columns: [{ width: 14 }, { width: 12 }, { width: 12 }],
        rows: [
          [
            { value: 'Region', style: header },
            { value: 'Revenue', style: header },
            { value: 'Date', style: header },
          ],
          [],
          [
            'North',
            { value: 1200.5, style: { numFmt: '0.00' } },
            { value: LEAP_DAY, style: { numFmt: 'yyyy-mm-dd' } },
          ],
          [],
          [
            { value: 'Total', style: bold },
            { formula: 'SUM(B1:B4)', value: 1200.5 },
          ],
          [],
          ['Footnote'],
        ],
        merges: ['A7:C7'],
        freezePane: { rows: 1 },
      },
      { name: 'Hidden', rows: [['secret']], hidden: true },
    ],
  });
};

try {
  writeFileSync(join(workDir, 'package.json'), JSON.stringify({ name: 'fixtures', private: true }));
  execFileSync(
    'npm',
    [
      'install',
      'exceljs@4.4.0',
      'xlsx@0.18.5',
      'hucre@1.2.0',
      '--no-audit',
      '--no-fund',
      '--ignore-scripts',
    ],
    { cwd: workDir, stdio: 'inherit' }
  );

  const require = createRequire(join(workDir, 'index.js'));
  const ExcelJS = require('exceljs');
  const XLSX = require('xlsx');
  const { writeXlsx } = require('hucre');

  const files = {
    'exceljs-report.xlsx': await buildExceljsReport(ExcelJS),
    'exceljs-text.xlsx': await buildExceljsText(ExcelJS),
    'sheetjs-report.xlsx': buildSheetjsReport(XLSX),
    'hucre-report.xlsx': await buildHucreReport(writeXlsx),
  };

  mkdirSync(outDir, { recursive: true });
  for (const [name, bytes] of Object.entries(files)) writeFileSync(join(outDir, name), bytes);
} finally {
  rmSync(workDir, { recursive: true, force: true });
}
