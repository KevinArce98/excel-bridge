import { existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const distDir = join(__dirname, '..', 'dist');
const distEntry = join(distDir, 'index.mjs');

if (!existsSync(distEntry)) {
  console.error('dist/ not found. Build first:\n  pnpm run build');
  process.exit(1);
}

const ROWS = Number(process.env.ROWS ?? 50_000);
const COLS = 10;

const buildData = () => {
  const rows = [Array.from({ length: COLS }, (_, c) => `Col ${c + 1}`)];
  for (let r = 0; r < ROWS; r++) {
    const row = [];
    for (let c = 0; c < COLS; c++) {
      row.push(c % 2 === 0 ? `Cell ${r}-${c}` : (r * c) % 1000);
    }
    rows.push(row);
  }
  return rows;
};

const median = (xs) => {
  const s = [...xs].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

const time = async (fn, runs = 3) => {
  let bytes = 0;
  const times = [];
  for (let i = 0; i < runs; i++) {
    const t0 = performance.now();
    const out = await fn();
    times.push(performance.now() - t0);
    bytes = out?.length ?? out?.byteLength ?? 0;
  }
  return { ms: median(times), bytes };
};

const fmtMs = (ms) => `${ms.toFixed(0).padStart(6)} ms`;
const fmtMB = (b) => `${(b / 1024 / 1024).toFixed(2).padStart(6)} MB`;

const load = async (name, spec) => {
  try {
    return await import(spec);
  } catch {
    console.log(`  (skipped ${name} — run "pnpm add -D ${spec}" to include it)`);
    return null;
  }
};

const main = async () => {
  console.log(`\nexcel-bridge benchmark — ${ROWS.toLocaleString()} rows × ${COLS} cols\n`);

  const data = buildData();
  const results = [];

  const eb = await import(distEntry);

  results.push([
    'excel-bridge (write)',
    await time(() => eb.ExcelBridge.writeBuffer(data)),
  ]);

  results.push([
    'excel-bridge (stream)',
    await time(async () => {
      function* rows() {
        for (const row of data) yield row;
      }
      return eb.streamToBuffer(eb.createExcelWorkbookStream([{ name: 'Sheet1', rows: rows() }]));
    }),
  ]);

  const exceljs = await load('exceljs', 'exceljs');
  if (exceljs) {
    const ExcelJS = exceljs.default ?? exceljs;
    results.push([
      'exceljs (write)',
      await time(async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addRows(data);
        return wb.xlsx.writeBuffer();
      }),
    ]);
  }

  const xlsx = await load('xlsx (SheetJS)', 'xlsx');
  if (xlsx) {
    const XLSX = xlsx.default ?? xlsx;
    results.push([
      'xlsx / SheetJS (write)',
      await time(() => {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      }),
    ]);
  }

  const label = 'Library';
  console.log(`\n${label.padEnd(24)} ${'Time'.padStart(9)}   ${'Output'.padStart(9)}`);
  console.log('-'.repeat(48));
  for (const [name, { ms, bytes }] of results) {
    console.log(`${name.padEnd(24)} ${fmtMs(ms)}   ${fmtMB(bytes)}`);
  }

  console.log('\nBundle (dist, shipped to consumers):');
  for (const file of ['index.js', 'index.mjs']) {
    const p = join(distDir, file);
    if (existsSync(p)) console.log(`  ${file.padEnd(12)} ${(statSync(p).size / 1024).toFixed(1)} KB`);
  }
  console.log('');
};

main();
