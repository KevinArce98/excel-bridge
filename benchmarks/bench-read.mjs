import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const distEntry = join(__dirname, '..', 'dist', 'index.mjs');

if (!existsSync(distEntry)) {
  console.error('dist/ not found. Build first:\n  pnpm run build');
  process.exit(1);
}

const ROWS = Number(process.env.ROWS ?? 50_000);
const RUNS = Number(process.env.RUNS ?? 5);
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

const median = values => {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};

const main = async () => {
  const { ExcelBridge } = await import(distEntry);
  const bytes = ExcelBridge.writeBuffer(buildData());

  const times = [];
  let cells = 0;
  for (let run = 0; run < RUNS; run++) {
    globalThis.gc?.();
    const started = performance.now();
    const workbook = ExcelBridge.read(bytes);
    times.push(performance.now() - started);
    cells = workbook.sheets[0].data.reduce((count, row) => count + row.length, 0);
  }

  const peakMiB = process.resourceUsage().maxRSS / 1024;
  console.log(
    `\nexcel-bridge read — ${ROWS.toLocaleString()} rows × ${COLS} cols, ${(bytes.length / 1024 / 1024).toFixed(2)} MiB file`
  );
  console.log(
    `  median of ${RUNS}: ${median(times).toFixed(0)} ms (${cells.toLocaleString()} cells)`
  );
  console.log(`  peak resident memory of this process: ${peakMiB.toFixed(0)} MiB\n`);
};

main();
