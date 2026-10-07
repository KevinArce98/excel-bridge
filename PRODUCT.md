# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Static, self-contained `docs/index.html` (HTML + CSS + a single ES module), deployed via GitHub Pages from the `docs/` folder on `main`. No build step. The library is loaded at runtime as browser ESM from jsDelivr (`https://cdn.jsdelivr.net/npm/excel-bridge@<version>/+esm`), so the page runs the actual published npm package. Decided for this surface (a demo/landing page for a GitHub-hosted OSS project); a heavier framework would add a build step with no benefit for a single static page.

## Users

Developers choosing an `.xlsx` read/write library for TypeScript/JavaScript that runs in both the browser (`File`/`Blob`) and Node.js (`Uint8Array`). Typically evaluating against ExcelJS or SheetJS and looking for something lighter and tree-shakeable to drop into a front-end bundle. They arrive from npm, GitHub, or a search, and decide quickly based on size, capability, and proof.

## Product Purpose

`excel-bridge` reads and writes `.xlsx` spreadsheets in the browser and Node.js without pulling in ExcelJS or SheetJS. This surface — the project's demo/landing page — exists to let a developer understand what the library is and try it for real in under a minute, then install it. Success = the visitor generates or reads a real `.xlsx` in the browser and leaves with the install command.

## Positioning

One tree-shakeable package: import a class and a bundler ships only that part. Two direct runtime dependencies (`fflate` for zip, `fast-xml-parser` for reading; the writer bundles no parser code), ESM **and** CJS, TypeScript-first, synchronous. One API spans browser and Node. Read **and** write — cell styles, formulas, dates, merged cells, freeze panes, conditional formatting, data validation, autofilters, hyperlinks, hidden sheets, multi-sheet — plus a streaming writer for very large exports and a `Workbook` load/edit/save API for files the library can model. The defensible claim is the combination, stated with its limits: SheetJS gates styling/conditional-formatting/streaming behind its Pro edition; ExcelJS carries a large CJS-first runtime; newer libraries such as hucre cover more features at 3x the writer size. It is not the smallest on every row and does not do images, charts, comments, tables, per-side borders or file-preserving edits.

## Operating Context

Evaluation happens at a desk, in a browser, alongside npm, GitHub, and Bundlephobia tabs. Developers compare install size, skim a README, and want to try before adopting (the README already points to a RunKit sandbox). The demo therefore runs the real published package client-side: build a styled workbook and download it (open in Excel to verify), or drop in an existing `.xlsx` and see it parsed. The page is the trial.

## Capabilities and Constraints

Confirmed capabilities (v1.4.0):

- `ExcelBridge.read` / `readFromFile` / `write` / `writeBuffer` entry points.
- Classes: `Workbook` (load/edit/save), `ExcelReader`, `ExcelWriter`, `StyleManager`.
- Cell styling: background, bold, italic, underline, color, fontSize, fontName, align, verticalAlign, wrapText, numberFormat, border.
- Formulas (strings starting with `=`), `Date` → Excel serial conversion, merged cells, freeze panes, column widths / autoWidth.
- Conditional formatting: `cellValue`, `expression`, `colorScale` — written and read back, and kept by a `Workbook` round trip.
- Data validation builders: `list`, `wholeNumber`, `decimal`, `textLength`, `dateBetween`; the types `time` and `custom` are written from plain objects. Rules are read back in full and kept by a `Workbook` round trip.
- AutoFilter (`options.autoFilter`) and hyperlinks (`hyperlinks` array, `hyperlink.url` / `email` / `internal` builders; `http`, `https`, `mailto` and in-workbook locations) — written by both writers, read back, and kept by a `Workbook` round trip.
- Streaming writer: `createExcelWorkbookStream` / `streamToBuffer` for million-row exports from sync/async iterables.
- Shared strings opt-in (ExcelWriter only); coordinate + date helper utilities; hidden and very hidden sheets (`state`).
- Input checks that throw: sheet names (Excel's rules), hex colours, finite numbers and valid dates.

Constraints:

- Inline strings by default (enable `sharedStrings: true` for smaller files with repeated text).
- Formula cells are written without a cached value; Excel recalculates on open (`fullCalcOnLoad`).
- Node engines `^20.19.0 || ^22.13.0 || >=24`; browsers need ES2022 + `File`/`Blob`; full Excel compatibility targets Excel 2016+.
- The streaming writer does not support `autoWidth`, `validations`, `conditionalFormats`, `sharedStrings` or sheet `state`.
- Strings starting with `=` are always written as formulas; date cells cannot be styled; there are no per-side borders, row heights or hidden rows/columns.
- `Workbook` rebuilds the file from its model: images, charts, comments, tables, defined names, print setup, themes and macros of a loaded file are dropped.
- The reader holds the whole file in memory, rejects out-of-grid references and caps empty padding cells at 5,000,000; it is not hardened for untrusted files beyond that. Known read gaps: 1904 dates, shared-formula followers, ISO date cells, prefixed namespaces, split panes.

## Brand Commitments

- Name `excel-bridge`, always lowercase, hyphenated. Wordmark treatment: `excel` near-white, `-` slate, `bridge` green.
- Existing brand asset: [assets/banner.svg](assets/banner.svg).
- Palette in use across the banner and npm badges: ground navy `#0F172A → #1E293B`, surface `#1B2336`, hairline `#334155`, accent green `#22C55E` with `#34D399` highlight; slate text scale `#F8FAFC / #CBD5E1 / #94A3B8 / #64748B`.
- Type: JetBrains Mono for the wordmark and code; Inter for prose.
- License MIT © Kevin Arias (github.com/KevinArce98). Repository: github.com/KevinArce98/excel-bridge.

## Evidence on Hand

- Real, published npm package `excel-bridge@1.4.0` (provenance-signed) — the demo loads it live from jsDelivr. The repository's main branch is ahead of it (reader limits, round-trip and writer validation fixes); the demo page keeps showing the 1.4.0 figures below until it moves to the next release.
- Real benchmark (README, `benchmarks/`): writing 50,000 rows × 10 columns, median of 5 runs, Node 24.19, Apple M4 — excel-bridge **560 ms / 2.41 MiB**, hucre 390 ms / 2.79 MiB, exceljs 1490 ms / 2.82 MiB, xlsx (SheetJS) 510 ms / 18.23 MiB uncompressed (570 ms / 6.26 MiB with `compression: true`). Reproducible via `pnpm run bench`. Never claim "fastest".
- Real bundle-size measurements (README "Bundle size", `benchmarks/size.mjs`): esbuild browser ESM bundles, min+gzip. Published 1.4.0 (measured 2026-09-24, esbuild 0.27.3): `ExcelWriter` **12.0 KB**, `ExcelReader` 27.6 KB, full import 41.5 KB. Repository main (measured 2026-10-07, esbuild 0.28.2): `ExcelWriter` **12.5 KB**, `ExcelReader` 28.2 KB, full import 42.5 KB. Others (2026-10-07): hucre 1.2.0 `writeXlsx` 41.6 KB and `XlsxStreamWriter` 12.1 KB, excelents 1.0.1 12.2 KB (reader and writer, no conditional formatting), SheetJS `utils` + `write` 95.8 KB, ExcelJS 272.1 KB. Reproducible via `pnpm run size`; show them with their method footnote and never say "smallest" without the qualifier.
- Real capability comparison table (README) vs ExcelJS and SheetJS community edition.
- Real API surface (`src/index.ts`) and banner asset.
- **No** testimonials, named customers, download counts, or endorsements exist — these must not be fabricated. Dynamic badges (npm version and downloads) are the only live third-party numbers and should be linked, not hardcoded. Download counts follow releases, so they say little about use.

## Product Principles

- Lightweight over feature-maximalism: every feature must still earn a place in a front-end bundle.
- One API across browser and Node — never make the caller choose an environment-specific path.
- Typed everything; IntelliSense for every public export.
- Prove with real numbers and a working trial; never invent claims, customers, or benchmarks.
- Round-trips keep the features the library writes, and the docs say plainly what a load/edit/save drops.

## Accessibility & Inclusion

No product-specific standard was established. Baseline for this public page: full keyboard operability for the interactive demo, visible focus, WCAG-AA contrast on the dark ground, honored `prefers-reduced-motion`, and file-drop that also works via a standard file picker.
