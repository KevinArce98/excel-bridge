import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'tests', 'fixtures');
const workDir = mkdtempSync(join(tmpdir(), 'excel-bridge-fixtures-'));

const SUMMER_DAY = new Date(Date.UTC(2024, 1, 29));

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

  mkdirSync(outDir, { recursive: true });

  const exceljs = new ExcelJS.Workbook();
  exceljs.creator = 'excel-bridge fixtures';
  const report = exceljs.addWorksheet('Report', { views: [{ state: 'frozen', ySplit: 1 }] });
  report.getColumn(1).width = 14;
  report.getColumn(2).width = 12;
  report.getColumn(3).width = 12;
  report.getCell('A1').value = 'Region';
  report.getCell('B1').value = 'Revenue';
  report.getCell('C1').value = 'Date';
  for (const ref of ['A1', 'B1', 'C1']) {
    report.getCell(ref).font = { bold: true };
    report.getCell(ref).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' },
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
  report.getCell('C3').value = SUMMER_DAY;
  report.getCell('C3').numFmt = 'yyyy-mm-dd';
  report.getCell('A5').value = 'Total';
  report.getCell('A5').font = { bold: true };
  report.getCell('B5').value = { formula: 'SUM(B1:B4)', result: 1200.5 };
  report.getCell('A7').value = 'Footnote';
  report.mergeCells('A7:C7');
  exceljs.addWorksheet('Hidden', { state: 'hidden' }).getCell('A1').value = 'secret';
  await exceljs.xlsx.writeFile(join(outDir, 'exceljs-report.xlsx'));

  const sheet = {
    A1: { t: 's', v: 'Region' },
    B1: { t: 's', v: 'Revenue' },
    C1: { t: 's', v: 'Date' },
    A3: { t: 's', v: 'North' },
    B3: { t: 'n', v: 1200.5, z: '0.00' },
    C3: { t: 'd', v: SUMMER_DAY, z: 'yyyy-mm-dd' },
    A5: { t: 's', v: 'Total' },
    B5: { t: 'n', f: 'SUM(B1:B4)', v: 1200.5 },
    A7: { t: 's', v: 'Footnote' },
    '!ref': 'A1:C7',
    '!merges': [XLSX.utils.decode_range('A7:C7')],
    '!cols': [{ wch: 14 }, { wch: 12 }, { wch: 12 }],
  };
  const sheetjs = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(sheetjs, sheet, 'Report');
  XLSX.utils.book_append_sheet(sheetjs, { A1: { t: 's', v: 'secret' }, '!ref': 'A1' }, 'Hidden');
  sheetjs.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };
  writeFileSync(
    join(outDir, 'sheetjs-report.xlsx'),
    XLSX.write(sheetjs, { bookType: 'xlsx', type: 'buffer', cellDates: true })
  );

  const bold = { font: { bold: true } };
  const header = { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: '#FFFF00' } };
  const hucre = await writeXlsx({
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
            { value: SUMMER_DAY, style: { numFmt: 'yyyy-mm-dd' } },
          ],
          [],
          [{ value: 'Total', style: bold }, { formula: 'SUM(B1:B4)', value: 1200.5 }],
          [],
          ['Footnote'],
        ],
        merges: ['A7:C7'],
        freezePane: { rows: 1 },
      },
      { name: 'Hidden', rows: [['secret']], hidden: true },
    ],
  });
  writeFileSync(join(outDir, 'hucre-report.xlsx'), hucre);
} finally {
  rmSync(workDir, { recursive: true, force: true });
}
