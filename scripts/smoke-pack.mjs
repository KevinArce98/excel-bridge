import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const required = new Set((process.env.SMOKE_REQUIRE_RUNTIMES ?? '').split(',').filter(Boolean));

if (!existsSync(join(root, 'dist', 'index.mjs'))) {
  console.error('dist/ not found. Build first:\n  pnpm run build');
  process.exit(1);
}

const body = `
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

const writer = new ExcelWriter();
const bytes = writer.createWorkbookBuffer([
  { data: [['name', 'qty'], ['pen', 2]], options: { name: 'Stock' } },
]);
assert(bytes instanceof Uint8Array && bytes.length > 100, 'writer returned no bytes');

const parsed = new ExcelReader().parseFromBuffer(bytes);
assert(parsed.sheets[0].name === 'Stock', 'sheet name lost');
assert(parsed.sheets[0].data[1][1].value === 2, 'cell value lost');

const workbook = Workbook.fromBuffer(bytes);
workbook.setCellValue('Stock', 0, 0, 'item');
const edited = new ExcelReader().parseFromBuffer(workbook.toBuffer());
assert(edited.sheets[0].data[0][0].value === 'item', 'workbook edit lost');

const streamed = await streamToBuffer(
  createExcelWorkbookStream([{ name: 'Rows', rows: [['id', 'total'], [1, 10]] }])
);
const streamedSheet = new ExcelReader().parseFromBuffer(streamed).sheets[0];
assert(streamedSheet.data[1][1].value === 10, 'streamed cell lost');

console.log('ok');
`;

const names = 'ExcelWriter, ExcelReader, Workbook, createExcelWorkbookStream, streamToBuffer';
const files = {
  'smoke.mjs': `import { ${names} } from 'excel-bridge';\n${body}`,
  'smoke.cjs': `const { ${names} } = require('excel-bridge');\n(async () => {\n${body}\n})().catch(error => {\n  console.error(error);\n  process.exit(1);\n});\n`,
};

const runtimes = [
  { name: 'node (esm)', command: process.execPath, args: ['smoke.mjs'], runtime: 'node' },
  { name: 'node (cjs)', command: process.execPath, args: ['smoke.cjs'], runtime: 'node' },
  { name: 'bun', command: 'bun', args: ['smoke.mjs'], runtime: 'bun' },
  { name: 'deno', command: 'deno', args: ['run', '--allow-read', 'smoke.mjs'], runtime: 'deno' },
];

const workDir = mkdtempSync(join(tmpdir(), 'excel-bridge-smoke-'));
const failures = [];

try {
  const [{ filename }] = JSON.parse(
    execFileSync('npm', ['pack', '--json', '--pack-destination', workDir], {
      cwd: root,
      encoding: 'utf8',
    })
  );

  const appDir = join(workDir, 'app');
  mkdirSync(appDir);
  writeFileSync(join(appDir, 'package.json'), JSON.stringify({ name: 'smoke', private: true }));
  execFileSync(
    'npm',
    ['install', join(workDir, filename), '--no-audit', '--no-fund', '--ignore-scripts'],
    { cwd: appDir, stdio: 'inherit' }
  );
  for (const [name, content] of Object.entries(files)) writeFileSync(join(appDir, name), content);

  for (const { name, command, args, runtime } of runtimes) {
    const result = spawnSync(command, args, { cwd: appDir, encoding: 'utf8' });

    if (result.error?.code === 'ENOENT') {
      if (required.has(runtime)) failures.push(`${name}: ${command} is required but not installed`);
      console.log(`skip ${name} (${command} not installed)`);
    } else if (result.status === 0 && result.stdout.trim() === 'ok') {
      console.log(`pass ${name}`);
    } else {
      failures.push(`${name}: exit ${result.status}\n${result.stdout}${result.stderr}`);
      console.log(`FAIL ${name}`);
    }
  }
} finally {
  rmSync(workDir, { recursive: true, force: true });
}

if (failures.length) {
  console.error(`\n${failures.join('\n\n')}`);
  process.exit(1);
}
