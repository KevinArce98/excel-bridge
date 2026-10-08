import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { transformSync } from 'esbuild';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const sourceDir = join(root, 'docs-src');
const typesEntry = join(root, 'dist', 'index.d.mts');
const runtimeEntry = pathToFileURL(join(root, 'dist', 'index.mjs')).href;

if (!existsSync(typesEntry)) {
  console.error('dist/ not found. Build first:\n  pnpm run build');
  process.exit(1);
}

const listMarkdown = dir =>
  readdirSync(dir).flatMap(name => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? listMarkdown(path) : name.endsWith('.md') ? [path] : [];
  });

const blocks = listMarkdown(sourceDir).flatMap(file =>
  [...readFileSync(file, 'utf8').matchAll(/```(ts|typescript)([^\n]*)\n([\s\S]*?)```/g)].map((match, index) => ({
    id: `${relative(sourceDir, file).replace(/[\\/]/g, '__').replace(/\.md$/, '')}__${index}`,
    file: relative(root, file),
    info: match[2],
    code: match[3],
  }))
);

const placeholders = [
  ['buffer', 'declare const buffer: Uint8Array;'],
  ['file', 'declare const file: File;'],
  ['response', 'declare const response: Response;'],
];

const prepare = block => {
  const header = placeholders
    .filter(([name]) => new RegExp(`\\b${name}\\b`).test(block.code) && !new RegExp(`\\b(const|let|var|function)\\s+${name}\\b`).test(block.code))
    .map(([, declaration]) => declaration)
    .join('\n');
  return `${header}${header ? '\n' : ''}${block.code}\nexport {};\n`;
};

const work = mkdtempSync(join(tmpdir(), 'excel-bridge-docs-'));
const failures = [];

try {
  const checked = blocks.filter(block => !/\bcheck=false\b/.test(block.info));
  mkdirSync(join(work, 'src'));
  writeFileSync(join(work, 'shim.d.ts'), `declare module 'node:fs' {\n  const fs: {\n    readFileSync(path: string): Uint8Array;\n    writeFileSync(path: string, data: Uint8Array | string): void;\n    createWriteStream(path: string): { write(chunk: Uint8Array): void; end(): void };\n  };\n  export default fs;\n}\n`);
  const files = checked.map(block => {
    const name = `src/${block.id}.ts`;
    writeFileSync(join(work, name), prepare(block));
    return name;
  });
  writeFileSync(
    join(work, 'tsconfig.json'),
    JSON.stringify({
      compilerOptions: {
        target: 'ES2022',
        module: 'ESNext',
        moduleResolution: 'Bundler',
        strict: true,
        noEmit: true,
        skipLibCheck: true,
        lib: ['ES2022', 'DOM', 'DOM.Iterable'],
        types: [],
        baseUrl: '.',
        paths: { 'excel-bridge': [typesEntry] },
      },
      files: [...files, 'shim.d.ts'],
    })
  );
  const compile = spawnSync(join(root, 'node_modules', '.bin', 'tsc'), ['-p', join(work, 'tsconfig.json')], { encoding: 'utf8' });
  const diagnostics = compile.stdout.split('\n').filter(line => /\.ts\(\d+,\d+\)/.test(line));
  diagnostics.forEach(line => {
    const id = /src\/([^(]+)\.ts/.exec(line)?.[1];
    const block = checked.find(entry => entry.id === id);
    failures.push(`types  ${block?.file ?? id}: ${line.replace(/^src\/[^(]+\.ts/, '').trim()}`);
  });

  const seed = join(work, 'seed.xlsx');
  const { ExcelWriter } = await import(runtimeEntry);
  writeFileSync(seed, new ExcelWriter().createWorkbookBuffer([{ data: [['Name', 'Qty'], ['a', 1]], options: { name: 'Sales' } }]));
  for (const name of ['data.xlsx', 'sales.xlsx']) writeFileSync(join(work, name), readFileSync(seed));

  let ran = 0;
  for (const block of checked) {
    if (/\brun=false\b/.test(block.info)) continue;
    const needsSeed = /\b(buffer|file)\b/.test(block.code) && !/\b(const|let|var)\s+(buffer|file)\b/.test(block.code);
    const prelude = needsSeed
      ? `import fs0 from 'node:fs';\nconst buffer = fs0.readFileSync('seed.xlsx');\nconst file = new Blob([buffer]);\n`
      : '';
    const source = (prelude + block.code).replaceAll("'excel-bridge'", `'${runtimeEntry}'`).replace(/(?<=const rowCount = )[\d_]+/, '2_000').replace('1_000_000', '2_000');
    const { code } = transformSync(source, { loader: 'ts', format: 'esm', target: 'es2022' });
    const path = join(work, `${block.id}.mjs`);
    writeFileSync(path, code);
    const run = spawnSync(process.execPath, [path], { cwd: work, encoding: 'utf8', timeout: 60_000, env: { ...process.env, TZ: 'UTC' } });
    ran += 1;
    if (run.status !== 0) {
      const reason = run.stderr.split('\n').find(line => /Error/.test(line)) ?? run.stderr.slice(0, 160);
      failures.push(`run    ${block.file} (block ${block.id.split('__').pop()}): ${reason}`);
    }
  }

  console.log(`docs examples: ${checked.length} type-checked, ${ran} run, ${blocks.length - checked.length} skipped`);
} finally {
  rmSync(work, { recursive: true, force: true });
}

const exportedNames = readFileSync(typesEntry, 'utf8')
  .split('\n')
  .filter(line => line.startsWith('export {'))
  .flatMap(line => [...line.matchAll(/(?:type\s+)?([A-Za-z_$][\w$]*)(?:\s+as\s+[\w$]+)?(?=,|\s*\})/g)].map(match => match[1]))
  .filter(name => name !== 'type' && name !== 'export');

for (const language of ['en', 'es']) {
  const page = join(sourceDir, language, 'api.md');
  if (!existsSync(page)) continue;
  const text = readFileSync(page, 'utf8');
  exportedNames
    .filter(name => !new RegExp(`(?<![\\w$])${name.replace(/\$/g, '\\$')}(?![\\w$])`).test(text))
    .forEach(name => failures.push(`api    ${relative(root, page)}: export \`${name}\` is not documented`));
}

if (failures.length) {
  console.error(`\n${failures.length} example(s) fail:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
