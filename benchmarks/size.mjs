import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { gzipSync } from 'node:zlib';
import { build, version as esbuildVersion } from 'esbuild';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const distEntry = join(root, 'dist', 'index.mjs');

if (!existsSync(distEntry)) {
  console.error('dist/ not found. Build first:\n  pnpm run build');
  process.exit(1);
}

const SELF = 'excel-bridge';

const CASES = [
  [SELF, 'createExcelWorkbookStream'],
  [SELF, 'ExcelWriter'],
  [SELF, 'ExcelReader'],
  [SELF, 'Workbook'],
  [SELF, 'ExcelBridge'],
  [SELF, '*'],
  ['hucre', 'writeXlsx'],
  ['hucre', 'readXlsx'],
  ['hucre/xlsx', 'XlsxStreamWriter'],
  ['hucre/xlsx', 'streamXlsxRows'],
  ['@mitresthen/excelents', '*'],
  ['read-excel-file/browser', 'default'],
  ['write-excel-file/universal', 'default'],
  ['xlsx', 'utils, write'],
  ['exceljs', 'default'],
];

const competitorModules = process.env.COMPETITORS_DIR
  ? join(resolve(process.env.COMPETITORS_DIR), 'node_modules')
  : undefined;
const moduleRoots = [join(root, 'node_modules'), ...(competitorModules ? [competitorModules] : [])];

const packageOf = specifier =>
  specifier.split('/').slice(0, specifier.startsWith('@') ? 2 : 1).join('/');

const readVersion = packageJson =>
  existsSync(packageJson) ? JSON.parse(readFileSync(packageJson, 'utf8')).version : null;

const versionOf = name =>
  name === SELF
    ? readVersion(join(root, 'package.json'))
    : moduleRoots
        .map(modules => readVersion(join(modules, packageOf(name), 'package.json')))
        .find(Boolean) ?? null;

const entryFor = (name, names) => {
  const specifier = JSON.stringify(name === SELF ? distEntry : name);
  return names === '*' ? `export * from ${specifier};` : `export { ${names} } from ${specifier};`;
};

const measure = async contents => {
  const result = await build({
    stdin: { contents, resolveDir: root },
    nodePaths: moduleRoots,
    bundle: true,
    minify: true,
    platform: 'browser',
    format: 'esm',
    write: false,
    logLevel: 'silent',
  });
  const bytes = result.outputFiles[0].contents;
  return { min: bytes.length, gzip: gzipSync(bytes).length };
};

const fmtKB = bytes => `${(bytes / 1000).toFixed(1)} KB`;

const README_TOLERANCE_BYTES = 100;

const toKB = bytes => (bytes / 1000).toFixed(1);

const importLabel = names => (names === '*' ? '* (everything)' : `{ ${names} }`);

const tableRowPattern = names =>
  names === '*'
    ? /^\| (?:Everything|Todo) \| ([\d.]+) KB \|$/m
    : new RegExp(`^\\| \`${names}\`[^|]*\\| ([\\d.]+) KB \\|$`, 'm');

const benchmarkRowPattern = names =>
  new RegExp(
    `^\\| excel-bridge@[\\d.]+ \\(dist\\) \\| \`${escapeRegExp(importLabel(names))}\` \\| [\\d.]+ KB \\| ([\\d.]+) KB \\|$`,
    'm'
  );

const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const SIZE_TABLES = [
  ['docs-src/en/performance.md', tableRowPattern],
  ['docs-src/es/performance.md', tableRowPattern],
  ['benchmarks/README.md', benchmarkRowPattern],
];

const SIZE_MENTIONS = [
  ['README.md', /ExcelWriter-([\d.]+)%20KB/, ['ExcelWriter'], 'the README size badge'],
  ['README.md', /the writer adds ([\d.]+) KB min\+gzip/, ['ExcelWriter'], 'the README tagline'],
  [
    'README.md',
    /`ExcelWriter` adds ([\d.]+) KB and `ExcelReader` ([\d.]+) KB min\+gzip/,
    ['ExcelWriter', 'ExcelReader'],
    'the README highlights',
  ],
  [
    'docs-src/en/getting-started.md',
    /the writer adds ([\d.]+) KB min\+gzip/,
    ['ExcelWriter'],
    'the English getting started page',
  ],
  [
    'docs-src/es/getting-started.md',
    /el escritor suma ([\d.]+) KB min\+gzip/,
    ['ExcelWriter'],
    'the Spanish getting started page',
  ],
  [
    'docs-src/en/performance.md',
    /`ExcelBridge` is\s+([\d.]+) KB against\s+([\d.]+) KB for `ExcelWriter`/,
    ['ExcelBridge', 'ExcelWriter'],
    'the English note on the ExcelBridge object',
  ],
  [
    'docs-src/en/performance.md',
    /hyperlinks\s+in\s+([\d.]+) KB/,
    ['ExcelWriter'],
    'the English paragraph on what ExcelWriter covers',
  ],
  [
    'docs-src/en/performance.md',
    /\| Bundle size to write a file ¹ \| \*\*([\d.]+) KB\*\*/,
    ['ExcelWriter'],
    'the English comparison table',
  ],
  [
    'docs-src/es/performance.md',
    /`ExcelBridge` pesa\s+([\d.]+) KB frente a\s+([\d.]+) KB/,
    ['ExcelBridge', 'ExcelWriter'],
    'the Spanish note on the ExcelBridge object',
  ],
  [
    'docs-src/es/performance.md',
    /hipervínculos\s+en\s+([\d.]+) KB/,
    ['ExcelWriter'],
    'the Spanish paragraph on what ExcelWriter covers',
  ],
  [
    'docs-src/es/performance.md',
    /\| Tamaño del paquete para escribir un archivo ¹ \| \*\*([\d.]+) KB\*\*/,
    ['ExcelWriter'],
    'the Spanish comparison table',
  ],
];

const checkReadme = async () => {
  const failures = [];
  const documented = {};

  for (const [, names] of CASES.filter(([name]) => name === SELF)) {
    const { gzip } = await measure(entryFor(SELF, names));
    const label = names === '*' ? 'everything' : names;
    const measuredKB = toKB(gzip);
    documented[names] = measuredKB;

    for (const [file, rowPattern] of SIZE_TABLES) {
      const match = rowPattern(names).exec(readFileSync(join(root, file), 'utf8'));
      if (!match) {
        failures.push(`${label}: no row in ${file}`);
        continue;
      }

      const drift = gzip - Number(match[1]) * 1000;
      const status = Math.abs(drift) <= README_TOLERANCE_BYTES ? 'ok' : 'FAIL';
      console.log(
        `${status.padEnd(4)} ${label.padEnd(28)} measured ${gzip} B, ${file} ${match[1]} KB (${drift >= 0 ? '+' : ''}${Math.round(drift)} B)`
      );
      if (status === 'FAIL') {
        failures.push(`${label}: ${file} says ${match[1]} KB, the build measures ${measuredKB} KB (${gzip} B)`);
      }
    }
  }

  for (const [file, pattern, subjects, where] of SIZE_MENTIONS) {
    const match = pattern.exec(readFileSync(join(root, file), 'utf8'));
    if (!match) {
      failures.push(`${where} (${file}): the sentence with the size was not found`);
      continue;
    }
    subjects.forEach((subject, index) => {
      const measured = documented[subject];
      if (measured !== undefined && Math.abs(Number(match[index + 1]) - Number(measured)) > 0.05) {
        failures.push(
          `${subject}: ${where} (${file}) says ${match[index + 1]} KB, the build measures ${measured} KB`
        );
      }
    });
  }

  if (failures.length) {
    console.error(`\nBundle size and documentation disagree:\n  ${failures.join('\n  ')}`);
    console.error('\nUpdate each figure above, then run: pnpm run build && pnpm run size:check');
    process.exit(1);
  }
};

const main = async () => {
  if (process.argv.includes('--check')) {
    await checkReadme();
    return;
  }

  console.log(
    `\nBundle size — esbuild ${esbuildVersion} (--bundle --minify --platform=browser --format=esm), gzip via Node ${process.version} zlib, 1 KB = 1,000 bytes\n`
  );

  const rows = [];
  const skipped = new Set();

  for (const [name, names] of CASES) {
    const version = versionOf(name);
    if (!version) {
      skipped.add(name);
      continue;
    }
    const { min, gzip } = await measure(entryFor(name, names));
    const label = name === SELF ? `${name}@${version} (dist)` : `${name}@${version}`;
    const imported = names === '*' ? '* (everything)' : `{ ${names} }`;
    rows.push(`| ${label} | \`${imported}\` | ${fmtKB(min)} | ${fmtKB(gzip)} |`);
  }

  console.log('| Package | Import | Min | Min+gzip |');
  console.log('| --- | --- | ---: | ---: |');
  for (const row of rows) console.log(row);

  if (skipped.size) {
    const names = [...skipped];
    console.log(
      `\n  (not installed: ${names.join(', ')} — install them in a scratch directory and run with COMPETITORS_DIR=<that directory>)`
    );
  }
  console.log('');
};

main();
