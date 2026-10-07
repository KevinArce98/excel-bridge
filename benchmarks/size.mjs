import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
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
  ['xlsx', 'utils, write'],
  ['exceljs', 'default'],
];

const readVersion = packageJson =>
  existsSync(packageJson) ? JSON.parse(readFileSync(packageJson, 'utf8')).version : null;

const versionOf = name =>
  name === SELF
    ? readVersion(join(root, 'package.json'))
    : readVersion(join(root, 'node_modules', name, 'package.json'));

const entryFor = (name, names) => {
  const specifier = JSON.stringify(name === SELF ? distEntry : name);
  return names === '*' ? `export * from ${specifier};` : `export { ${names} } from ${specifier};`;
};

const measure = async contents => {
  const result = await build({
    stdin: { contents, resolveDir: root },
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

const readmeRowPattern = names =>
  names === '*'
    ? /^\| Everything \| ([\d.]+) KB \|$/m
    : new RegExp(`^\\| \`${names}\`[^|]*\\| ([\\d.]+) KB \\|$`, 'm');

const benchmarkRowPattern = names =>
  new RegExp(
    `^\\| excel-bridge@[\\d.]+ \\(dist\\) \\| \`${escapeRegExp(importLabel(names))}\` \\| [\\d.]+ KB \\| ([\\d.]+) KB \\|$`,
    'm'
  );

const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const checkReadme = async () => {
  const readme = readFileSync(join(root, 'README.md'), 'utf8');
  const benchmarks = readFileSync(join(root, 'benchmarks', 'README.md'), 'utf8');
  const failures = [];
  const documented = {};

  for (const [name, names] of CASES.filter(([name]) => name === SELF)) {
    const { gzip } = await measure(entryFor(name, names));
    const label = names === '*' ? 'everything' : names;
    const measuredKB = toKB(gzip);

    for (const [file, pattern] of [
      ['README.md bundle size table', readmeRowPattern(names)],
      ['benchmarks/README.md reference results', benchmarkRowPattern(names)],
    ]) {
      const match = pattern.exec(file.startsWith('README') ? readme : benchmarks);
      if (!match) {
        failures.push(`${label}: no row in the ${file}`);
        continue;
      }
      if (file.startsWith('README')) documented[names] = match[1];

      const drift = gzip - Number(match[1]) * 1000;
      const status = Math.abs(drift) <= README_TOLERANCE_BYTES ? 'ok' : 'FAIL';
      console.log(
        `${status.padEnd(4)} ${label.padEnd(28)} measured ${gzip} B, ${file} ${match[1]} KB (${drift >= 0 ? '+' : ''}${Math.round(drift)} B)`
      );
      if (status === 'FAIL') {
        failures.push(
          `${label}: the ${file} says ${match[1]} KB, the build measures ${measuredKB} KB (${gzip} B)`
        );
      }
    }
  }

  const mentions = [
    [/ExcelWriter-([\d.]+)%20KB/, ['ExcelWriter'], 'the README size badge'],
    [/`ExcelWriter` alone is ([\d.]+) KB/, ['ExcelWriter'], 'the README highlights'],
    [
      /> ([\d.]+) KB min\+gzip, against ([\d.]+) KB for `ExcelWriter`/,
      ['ExcelBridge', 'ExcelWriter'],
      'the README note on the ExcelBridge object',
    ],
    [
      /\| Bundle size to write a file ¹ \| \*\*([\d.]+) KB\*\*/,
      ['ExcelWriter'],
      'the README comparison table',
    ],
  ];

  for (const [pattern, subjects, where] of mentions) {
    const match = pattern.exec(readme);
    if (!match) {
      failures.push(`${where}: the sentence with the size was not found`);
      continue;
    }
    subjects.forEach((subject, index) => {
      if (documented[subject] !== undefined && match[index + 1] !== documented[subject]) {
        failures.push(
          `${subject}: ${where} says ${match[index + 1]} KB, the bundle size table says ${documented[subject]} KB`
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
      `\n  (not installed: ${names.join(', ')} — add with "pnpm add -D ${names.join(' ')}")`
    );
  }
  console.log('');
};

main();
