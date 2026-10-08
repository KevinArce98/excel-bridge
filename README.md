<div align="center">

<img src="https://raw.githubusercontent.com/KevinArce98/excel-bridge/main/assets/banner.svg" alt="excel-bridge — the lightweight Excel toolkit for TypeScript" width="100%" />

<br />

**Read and write styled `.xlsx` files in the browser and Node.js, without shipping ExcelJS or SheetJS.**

Import a class and your bundler ships only that part: the writer adds 12.8 KB min+gzip.

<br />

[![Live demo](https://img.shields.io/badge/demo-live-22c55e?labelColor=1e293b)](https://kevinarce98.github.io/excel-bridge/)
[![npm version](https://img.shields.io/npm/v/excel-bridge?logo=npm&label=npm&color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![ExcelWriter size](https://img.shields.io/badge/ExcelWriter-12.8%20KB%20min%2Bgzip-22c55e?labelColor=1e293b)](https://kevinarce98.github.io/excel-bridge/performance/)
[![CI](https://img.shields.io/github/actions/workflow/status/KevinArce98/excel-bridge/ci.yml?branch=main&label=CI&logo=github)](https://github.com/KevinArce98/excel-bridge/actions)
[![types](https://img.shields.io/npm/types/excel-bridge?color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![license](https://img.shields.io/npm/l/excel-bridge?color=22c55e)](./LICENSE)

<sub>[Documentation](https://kevinarce98.github.io/excel-bridge/getting-started/) · [Documentación en español](https://kevinarce98.github.io/excel-bridge/es/getting-started/) · [Live demo](https://kevinarce98.github.io/excel-bridge/) · [API reference](https://kevinarce98.github.io/excel-bridge/api/) · [Upgrading](./UPGRADING.md)</sub>

</div>

---

## Why excel-bridge

- **Small and tree-shakeable.** `ExcelWriter` adds 12.8 KB and `ExcelReader` 9.5 KB min+gzip. Import only the class you need. ESM and CJS builds.
- **Read and write.** Cell styles with per-side borders, formulas with stored results, dates, merged cells, freeze panes, row heights, conditional formatting, data validation, autofilters, hyperlinks, hidden sheets and multi-sheet workbooks.
- **One API for the browser and Node.js.** `Blob` and `File` in the browser, `Uint8Array` in Node.js.
- **TypeScript-first.** Typed cells, typed styles and typed errors.
- **Built for large files.** A streaming writer for million-row exports, and reader limits for files you did not make.
- **One dependency.** `fflate`. The reader has its own small XML tokenizer.
- **Signed releases.** Releases from 1.2.0 on are published to npm with [provenance](https://docs.npmjs.com/generating-provenance-statements).

## Install

```bash
npm install excel-bridge
```

`pnpm add excel-bridge`, `yarn add excel-bridge` and `bun add excel-bridge` work the same way.

## Write and read

```typescript
import fs from 'node:fs';
import { ExcelReader, ExcelWriter } from 'excel-bridge';

const header = { bold: true, background: '#4472C4', color: '#FFFFFF' };

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Product', 'Price', 'Qty', 'Total'],
      ['Laptop', 999.99, 5, { formula: 'B2*C2' }],
      ['Mouse', 29.99, 20, { formula: 'B3*C3' }],
    ],
    styles: { '0-0': header, '0-1': header, '0-2': header, '0-3': header },
    options: { name: 'Sales', freezePane: { row: 1 }, autoWidth: true },
  },
]);

fs.writeFileSync('sales.xlsx', bytes);

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('sales.xlsx'));
workbook.sheets[0].data.forEach(row => row.forEach(cell => console.log(cell.coordinate, cell.type, cell.value)));
```

In the browser, `createWorkbook` returns a `Blob` and `parseFromFile` reads a `File`. The writers return a plain `Uint8Array`, not a `Buffer`.

## Documentation

The full documentation, in English and Spanish, is at **[kevinarce98.github.io/excel-bridge](https://kevinarce98.github.io/excel-bridge/getting-started/)**:

- [Getting started](https://kevinarce98.github.io/excel-bridge/getting-started/) · [Write values and formulas](https://kevinarce98.github.io/excel-bridge/guide/values/) · [Style cells](https://kevinarce98.github.io/excel-bridge/guide/styling/) · [Lay out sheets](https://kevinarce98.github.io/excel-bridge/guide/layout/) · [Add rules](https://kevinarce98.github.io/excel-bridge/guide/rules/)
- [Read a workbook](https://kevinarce98.github.io/excel-bridge/guide/reading/) · [Rows as objects](https://kevinarce98.github.io/excel-bridge/guide/objects/) · [Edit a workbook](https://kevinarce98.github.io/excel-bridge/guide/workbook/) · [Stream big exports](https://kevinarce98.github.io/excel-bridge/guide/streaming/) · [Handle errors and limits](https://kevinarce98.github.io/excel-bridge/guide/errors/)
- [API reference](https://kevinarce98.github.io/excel-bridge/api/) · [Performance and size](https://kevinarce98.github.io/excel-bridge/performance/) · [Limits and alternatives](https://kevinarce98.github.io/excel-bridge/limitations/) · [Upgrading](./UPGRADING.md)

## When another library fits better

excel-bridge writes reports and reads them back. It is not the right tool for:

- **Images, charts, comments, pivot tables or editing a file while keeping everything you did not touch.** ExcelJS and hucre cover many of these.
- **Very large files you must read without loading them whole.** `ExcelReader` has no streaming mode. hucre has `streamXlsxRows`.
- **CSV, ODS or legacy `.xls`.** Use hucre or SheetJS.

It is not the smallest on every row of the [size comparison](https://kevinarce98.github.io/excel-bridge/performance/), and nothing in CI opens a file in Excel, LibreOffice or Google Sheets. [Limits and alternatives](https://kevinarce98.github.io/excel-bridge/limitations/) has the full list.

## Security

Report vulnerabilities privately through [GitHub security advisories](https://github.com/KevinArce98/excel-bridge/security/advisories/new) (see [SECURITY.md](./SECURITY.md)). Strings are always written as text, so user data cannot become a formula. The reader has [limits](https://kevinarce98.github.io/excel-bridge/guide/errors/) but is not hardened for untrusted files beyond them.

## Contributing

Issues and pull requests are welcome. To work on the library locally:

```bash
pnpm install
pnpm run build          # bundle ESM + CJS + types
pnpm run test:run       # run the test suite once
pnpm run lint           # ESLint
pnpm run format:check   # Prettier
pnpm run check:package  # publint and arethetypeswrong on the built package
pnpm run size:check     # the documented bundle sizes against the build
pnpm run smoke:pack     # install the packed tarball and run it on Node, Bun and Deno
pnpm run docs:build     # regenerate docs/ from docs-src/
pnpm run docs:check     # fail if docs/ is out of date
pnpm run docs:examples  # type-check and run every example in docs-src/
```

Commits follow [Conventional Commits](https://www.conventionalcommits.org/); releases are published automatically by semantic-release. See the [CHANGELOG](./CHANGELOG.md) for release notes.

## License

[MIT](./LICENSE) © [Kevin Arias](https://github.com/KevinArce98)

Microsoft Excel is a trademark of Microsoft. This project is not affiliated with Microsoft.
