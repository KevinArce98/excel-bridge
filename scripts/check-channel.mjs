import { execFileSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const OPEN = '<!-- channel:prerelease -->';
const CLOSE = '<!-- /channel:prerelease -->';
const MARKER = 'channel:prerelease';
const INSTALL_COMMAND = 'excel-bridge@next';
const BANNER_FILES = [
  'README.md',
  'docs-src/en/getting-started.md',
  'docs-src/es/getting-started.md',
];
const SCANNED = ['README.md', 'UPGRADING.md', 'docs-src', 'docs'];
const TEXT_FILE = /\.(md|html)$/;

const argument = process.argv.find(value => value.startsWith('--channel='));
const currentBranch = () => {
  try {
    return execFileSync('git', ['rev-parse', '--abbrev-ref', 'HEAD'], { encoding: 'utf8' }).trim();
  } catch {
    return '';
  }
};
const channel =
  argument?.slice('--channel='.length) ||
  process.env.GITHUB_BASE_REF ||
  process.env.GITHUB_REF_NAME ||
  currentBranch();

const textFilesUnder = path => {
  if (!existsSync(path)) return [];
  if (!statSync(path).isDirectory()) return TEXT_FILE.test(path) ? [path] : [];
  return readdirSync(path).flatMap(name => textFilesUnder(join(path, name)));
};

const problems = [];

if (channel === 'next') {
  for (const file of BANNER_FILES) {
    const text = readFileSync(file, 'utf8');
    const start = text.indexOf(OPEN);
    const end = text.indexOf(CLOSE);
    if (start === -1 || end < start || text.indexOf(OPEN, start + 1) !== -1) {
      problems.push(`${file}: needs exactly one prerelease block between ${OPEN} and ${CLOSE}`);
    } else if (!text.slice(start, end).includes(INSTALL_COMMAND)) {
      problems.push(
        `${file}: the prerelease block must tell readers to install ${INSTALL_COMMAND}`
      );
    }
  }
} else if (channel === 'main') {
  for (const file of SCANNED.flatMap(textFilesUnder)) {
    if (readFileSync(file, 'utf8').includes(MARKER)) {
      problems.push(`${file}: remove the prerelease block before this reaches main`);
    }
  }
} else {
  console.log(`Channel "${channel}" is neither main nor next. Nothing to check.`);
  process.exit(0);
}

if (problems.length > 0) {
  console.error(`Release channel "${channel}" check failed:`);
  for (const problem of problems) console.error(`  ${problem}`);
  process.exit(1);
}
console.log(`Release channel "${channel}" is consistent.`);
