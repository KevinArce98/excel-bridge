# Benchmarks

Two scripts, both run against the local `dist/` build:

- [Write speed](#write-speed) — `bench.mjs` compares write time and output size against `exceljs`
  and `xlsx` (SheetJS).
- [Bundle size](#bundle-size) — `size.mjs` measures what each import adds to a browser bundle, next
  to `hucre`, `xlsx` and `exceljs`.

Competitors are **optional** — both scripts skip any that aren't installed, so they are not part
of this package's dependencies.

## Write speed

### Run it

```bash
pnpm run bench
```

To include the comparison against ExcelJS and SheetJS:

```bash
pnpm add -D exceljs xlsx
pnpm run bench
```

Tune the dataset size with the `ROWS` env var (default `50000`):

```bash
ROWS=200000 pnpm run bench
```

### Reference results

Writing **50,000 rows × 10 columns** (mixed strings and numbers), median of 3 runs
on an Apple Silicon laptop, Node 22. Numbers vary by machine — run it yourself for
your own hardware.

| Library | Write time | Output size |
| --- | ---: | ---: |
| **excel-bridge** (write) | **662 ms** | **2.41 MB** |
| **excel-bridge** (stream) | **641 ms** | 2.48 MB |
| exceljs (write) | 1667 ms | 2.82 MB |
| xlsx / SheetJS (write) | 578 ms | 18.23 MB |

Takeaways for this workload:

- **~2.5× faster than ExcelJS** with a smaller file.
- **Comparable speed to SheetJS but ~7.5× smaller output** at each library's
  default settings.
- The streaming writer matches the in-memory writer's speed while keeping memory
  flat — the win grows as row counts rise.

> Every library was driven with its documented defaults; no per-library tuning was
> applied. Treat these as directional, not absolute.

## Bundle size

Each row bundles a one-line entry such as `export { ExcelWriter } from 'excel-bridge'` with
esbuild (`--bundle --minify --platform=browser --format=esm`), then gzips the output with Node's
zlib at the default level. The bundler keeps that export and everything it references, which is
what the import costs a browser app.

### Run it

```bash
pnpm run size
```

To include the comparison against hucre, SheetJS and ExcelJS:

```bash
pnpm add -D hucre xlsx exceljs
pnpm run size
```

The script prints a Markdown table, ready to paste into the main README.

### Reference results

Measured 2026-09-24 with esbuild 0.27.3 and Node 22; 1 KB = 1,000 bytes.

| Package | Import | Min | Min+gzip |
| --- | --- | ---: | ---: |
| excel-bridge@1.4.0 (dist) | `{ createExcelWorkbookStream }` | 30.1 KB | 11.4 KB |
| excel-bridge@1.4.0 (dist) | `{ ExcelWriter }` | 31.8 KB | 12.0 KB |
| excel-bridge@1.4.0 (dist) | `{ ExcelReader }` | 79.1 KB | 27.6 KB |
| excel-bridge@1.4.0 (dist) | `{ Workbook }` | 112.6 KB | 38.9 KB |
| excel-bridge@1.4.0 (dist) | `{ ExcelBridge }` | 113.0 KB | 39.0 KB |
| excel-bridge@1.4.0 (dist) | `* (everything)` | 120.2 KB | 41.5 KB |
| hucre@1.1.0 | `{ writeXlsx }` | | ~40 KB |
| hucre@1.1.0 | `{ readXlsx }` | | ~40 KB |
| xlsx@0.18.5 | `{ utils, write }` | 287.4 KB | 95.8 KB |
| exceljs@4.4.0 | `{ default }` | 947.0 KB | 272.1 KB |

Notes:

- `ExcelBridge` is a single object, and bundlers keep an object whole: any `ExcelBridge.*` call
  ships the reader and the writer, about as much as `Workbook`.
- The hucre figures were measured separately with the same esbuild flags and the `gzip` CLI, so
  they are rounded; run the script with `hucre` installed for exact numbers.
- `exceljs` resolves to its prebuilt browser bundle (`dist/exceljs.min.js`), which can't be
  tree-shaken.
- `xlsx@0.18.5` is the SheetJS package on npm. Newer Community Edition builds ship from the SheetJS
  CDN and aren't measured here.
- gzip implementations differ by about 1% (macOS `gzip` comes out slightly smaller than Node's
  zlib), so compare numbers from the same run.
