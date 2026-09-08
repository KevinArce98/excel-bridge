# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Static, self-contained `docs/index.html` (HTML + CSS + a single ES module), deployed via GitHub Pages from the `docs/` folder on `main`. No build step. The library is loaded at runtime as browser ESM from jsDelivr (`https://cdn.jsdelivr.net/npm/excel-bridge@<version>/+esm`), so the page runs the actual published npm package. Decided for this surface (a demo/landing page for a GitHub-hosted OSS project); a heavier framework would add a build step with no benefit for a single static page.

## Users

Developers choosing an `.xlsx` read/write library for TypeScript/JavaScript that runs in both the browser (`File`/`Blob`) and Node.js (`Buffer`). Typically evaluating against ExcelJS or SheetJS and looking for something lighter and tree-shakeable to drop into a front-end bundle. They arrive from npm, GitHub, or a search, and decide quickly based on size, capability, and proof.

## Product Purpose

`excel-bridge` reads and writes `.xlsx` spreadsheets in the browser and Node.js without pulling in ExcelJS or SheetJS. This surface — the project's demo/landing page — exists to let a developer understand what the library is and try it for real in under a minute, then install it. Success = the visitor generates or reads a real `.xlsx` in the browser and leaves with the install command.

## Positioning

A micro-package architecture that ships only what you import: zero heavy runtime dependencies (just `fflate` + `fast-xml-parser`), ESM **and** CJS, tree-shakeable, TypeScript-first. One API spans browser and Node. Full read **and** write — styling, fonts, borders, formulas, dates, merged cells, freeze panes, conditional formatting, data validation, multi-sheet — plus a streaming writer for very large exports and a high-level `Workbook` load/edit/save API. This is the combination a neighbor cannot truthfully claim: SheetJS gates styling/conditional-formatting/streaming behind its Pro edition; ExcelJS carries a large CJS-first runtime.

## Operating Context

Evaluation happens at a desk, in a browser, alongside npm, GitHub, and Bundlephobia tabs. Developers compare install size, skim a README, and want to try before adopting (the README already points to a RunKit sandbox). The demo therefore runs the real published package client-side: build a styled workbook and download it (open in Excel to verify), or drop in an existing `.xlsx` and see it parsed. The page is the trial.

## Capabilities and Constraints

Confirmed capabilities (v1.3.0):

- `ExcelBridge.read` / `readFromFile` / `write` / `writeBuffer` entry points.
- Classes: `Workbook` (load/edit/save), `ExcelReader`, `ExcelWriter`, `StyleManager`.
- Cell styling: background, bold, italic, underline, color, fontSize, fontName, align, verticalAlign, wrapText, numberFormat, border.
- Formulas (strings starting with `=`), `Date` → Excel serial conversion, merged cells, freeze panes, column widths / autoWidth.
- Conditional formatting: `cellValue`, `expression`, `colorScale` — written and read back losslessly.
- Data validation builders: `list`, `wholeNumber`, `decimal`, `textLength`, `dateBetween`.
- Streaming writer: `createExcelWorkbookStream` / `streamToBuffer` for million-row exports from sync/async iterables.
- Shared strings opt-in; coordinate + date helper utilities.

Constraints:

- Inline strings by default (enable `sharedStrings: true` for smaller files with repeated text).
- Formula cells are written without a cached value; Excel recalculates on open (`fullCalcOnLoad`).
- Node engines `^20.19.0 || ^22.13.0 || >=24`; browsers need ES2022 + `File`/`Blob`; full Excel compatibility targets Excel 2016+.
- The streaming writer does not support `autoWidth`, `validations`, or `conditionalFormats`.

## Brand Commitments

- Name `excel-bridge`, always lowercase, hyphenated. Wordmark treatment: `excel` near-white, `-` slate, `bridge` green.
- Existing brand asset: [assets/banner.svg](assets/banner.svg).
- Palette in use across the banner and npm badges: ground navy `#0F172A → #1E293B`, surface `#1B2336`, hairline `#334155`, accent green `#22C55E` with `#34D399` highlight; slate text scale `#F8FAFC / #CBD5E1 / #94A3B8 / #64748B`.
- Type: JetBrains Mono for the wordmark and code; Inter for prose.
- License MIT © Kevin Arias (github.com/KevinArce98). Repository: github.com/KevinArce98/excel-bridge.

## Evidence on Hand

- Real, published npm package `excel-bridge@1.3.0` (provenance-signed) — the demo loads it live from jsDelivr.
- Real benchmark (README, `benchmarks/`): writing 50,000 rows × 10 columns, median of 3 runs, Node 22, Apple Silicon — excel-bridge **662 ms / 2.41 MB**, exceljs 1667 ms / 2.82 MB, xlsx (SheetJS) 578 ms / 18.23 MB. Reproducible via `pnpm run bench`.
- Real capability comparison table (README) vs ExcelJS and SheetJS community edition.
- Real API surface (`src/index.ts`) and banner asset.
- **No** testimonials, named customers, download counts, or endorsements exist — these must not be fabricated. Dynamic badges (npm version/downloads, bundlephobia size) are the only live third-party numbers and should be linked, not hardcoded.

## Product Principles

- Lightweight over feature-maximalism: every feature must still earn a place in a front-end bundle.
- One API across browser and Node — never make the caller choose an environment-specific path.
- Typed everything; IntelliSense for every public export.
- Prove with real numbers and a working trial; never invent claims, customers, or benchmarks.
- Lossless round-trips for the features the library writes.

## Accessibility & Inclusion

No product-specific standard was established. Baseline for this public page: full keyboard operability for the interactive demo, visible focus, WCAG-AA contrast on the dark ground, honored `prefers-reduced-motion`, and file-drop that also works via a standard file picker.
