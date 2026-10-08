---
title: Performance and size
description: Bundle size of every entry point next to other libraries, write and read speed, and how each number was measured.
group: Reference
groupOrder: 3
order: 2
---

# Performance and size

Every number on this page comes from a script in the repository, and the method sits next to it. Run the scripts to check them on your own machine.

## Bundle size

Bundlers keep the exports you import and drop the rest. This is what each entry point adds to a browser bundle:

| Import from `excel-bridge` | min+gzip |
| --- | ---: |
| `createExcelWorkbookStream` | 12.4 KB |
| `ExcelWriter` | 12.9 KB |
| `ExcelReader` | 9.9 KB |
| `Workbook` (reader + writer) | 22.0 KB |
| `ExcelBridge` (convenience object) | 22.1 KB |
| Everything | 25.8 KB |

The same measurement for other libraries (2026-10-07):

| Library | Import | min+gzip |
| --- | --- | ---: |
| hucre 1.2.0 | `writeXlsx` | 41.6 KB |
| hucre 1.2.0 | `readXlsx` | 41.1 KB |
| hucre 1.2.0 | `XlsxStreamWriter` (from `hucre/xlsx`) | 12.1 KB |
| hucre 1.2.0 | `streamXlsxRows` (from `hucre/xlsx`) | 17.4 KB |
| @mitresthen/excelents 1.0.1 | whole main entry, reader and writer | 12.2 KB |
| read-excel-file 9.3.10 | `read-excel-file/browser`, read only | 16.7 KB |
| write-excel-file 4.1.1 | `write-excel-file/universal`, write only | 19.8 KB |
| SheetJS (`xlsx` 0.18.5 on npm) | `utils` + `write` | 95.8 KB |
| ExcelJS 4.4.0 | default browser build, not tree-shakeable | 272.1 KB |

**excel-bridge is not the smallest on every row.** `ExcelWriter` writes styles, borders, conditional formatting, data validation, autofilters and hyperlinks in 12.9 KB. Some libraries are smaller for one job, or give up a feature to get there: the `excelents` entry has no conditional formatting at 1.0.1, and reads and writes in less than the whole of excel-bridge. Compare what you get for the bytes, not the bytes alone.

Tree-shaking uses the ESM build, which bundlers pick for `import`. `require('excel-bridge')` loads the whole CommonJS build. `ExcelBridge` is 22.1 KB against 12.9 KB for `ExcelWriter`, because bundlers keep an object whole: even one `ExcelBridge.write` call ships the reader too.

**How these numbers were measured**

- **Toolchain:** esbuild 0.28.2 (`--bundle --minify --platform=browser --format=esm`). The excel-bridge rows were measured on 2026-10-08 and the other libraries on 2026-10-07.
- **Entry:** each row bundles a one-line `export { … } from '<package>'` file, then gzips the output with Node's zlib at the default level. 1 KB = 1,000 bytes.
- **Sources:** the excel-bridge rows come from this repository's build. The other libraries were installed in a scratch directory.
- **Margin:** gzip implementations differ by about 1%. macOS `gzip` comes out slightly smaller.
- **Not measured:** newer SheetJS Community Edition builds, which ship from the SheetJS CDN.
- **Reproduce:** run `pnpm run size`. CI runs `pnpm run size:check`, which fails when an excel-bridge row drifts more than 100 bytes from the measured size.

## How it compares

| | **excel-bridge** | ExcelJS | SheetJS (community `xlsx`) |
| --- | :---: | :---: | :---: |
| Read `.xlsx` | ✅ | ✅ | ✅ |
| Write `.xlsx` | ✅ | ✅ | ✅ |
| Cell styling (color, font, per-side borders) | ✅ | ✅ | ⚠️ Pro edition |
| Conditional formatting | ✅ | ✅ | ⚠️ Pro edition |
| Formulas | ✅ | ✅ | ✅ |
| Merged cells | ✅ | ✅ | ✅ |
| Freeze panes | ✅ | ✅ | ❌ not in `xlsx` 0.18.5 |
| Streaming writer | ✅ | ✅ | ⚠️ Pro edition |
| First-class TypeScript types | ✅ | ✅ | ✅ |
| ESM **and** CJS, tree-shakeable | ✅ | ⚠️ CJS-first | ✅ |
| Direct runtime dependencies ² | 1 | 9 | 7 |
| Bundle size to write a file ¹ | **12.9 KB** | 272.1 KB | 95.8 KB |

¹ Minified and gzipped code that a browser bundle needs to write an `.xlsx`: `ExcelWriter`, ExcelJS's default browser build (not tree-shakeable) and SheetJS `utils` + `write` from npm `xlsx@0.18.5`, bundled with esbuild. Bundlephobia measures each whole package with its own toolchain, so its numbers differ: [excel-bridge](https://bundlephobia.com/package/excel-bridge), [exceljs](https://bundlephobia.com/package/exceljs) and [xlsx](https://bundlephobia.com/package/xlsx).

² Declared in each `package.json` (`exceljs@4.4.0`, `xlsx@0.18.5` on npm), checked 2026-10-07. `xlsx@0.18.5` is the last version published to npm and has known security advisories.

## Write speed

Writing 50,000 rows × 10 columns (Node 24.19, Apple M4, 2026-10-08). Each time is the median of 5 runs inside one invocation, taken over three invocations:

| Library | Write time | Output size |
| --- | ---: | ---: |
| **excel-bridge** | **560 ms** | **2.41 MiB** |
| excel-bridge (streaming) | 580 ms | 2.48 MiB |
| hucre | 390 ms | 2.79 MiB |
| xlsx / SheetJS (default, no compression) | 510 ms | 18.23 MiB |
| xlsx / SheetJS (`compression: true`) | 580 ms | 6.26 MiB |
| exceljs | 1430 ms | 2.82 MiB |

- **Against ExcelJS:** about 2.6× faster.
- **Against hucre:** hucre writes this workload faster, in about 390 ms.
- **Against SheetJS:** similar time, 510 ms with its default uncompressed output and 580 ms with `compression: true`. Its file is about 7.5× larger by default and about 2.6× larger with compression.

Times vary by 10% or more between runs and by machine, and ratios hold better than absolute times. Reproduce them with `pnpm run bench` after installing the other libraries (see the benchmarks README).

## Read speed

Reading a file that excel-bridge wrote, 50,000 rows × 10 columns and 200,000 rows × 10 columns (Node 24.19, Apple M4, 2026-10-08, median of 7 and 3 runs):

| File | 1.6.0 | 2.0 | Peak memory 1.6.0 | Peak memory 2.0 |
| --- | ---: | ---: | ---: | ---: |
| 50,000 × 10 | 1503 ms | 270 ms | 859 MiB | 451 MiB |
| 200,000 × 10 | 6521 ms | 1435 ms | 2679 MiB | 1611 MiB |

Version 2.0 reads about 4.5 to 5.6 times faster and peaks at about half to two thirds of the memory. The reader parses the XML with its own small tokenizer instead of a general-purpose parser, which is also why `ExcelReader` went from 28.8 to 9.5 KB.

The peak memory is that of the benchmark process, which also builds the file to read. The reader holds the whole file in memory and has no streaming mode. Reproduce the figures with `node --expose-gc benchmarks/bench-read.mjs`.
