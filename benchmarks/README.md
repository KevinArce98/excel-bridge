# Benchmarks

Three scripts, all run against the local `dist/` build:

- [Write speed](#write-speed) — `bench.mjs` compares write time and output size against `exceljs`,
  `xlsx` (SheetJS) and `hucre`.
- [Read speed](#read-speed) — `bench-read.mjs` times `ExcelBridge.read` on a file excel-bridge wrote.
- [Bundle size](#bundle-size) — `size.mjs` measures what each import adds to a browser bundle, next
  to `hucre`, `@mitresthen/excelents`, `read-excel-file`, `write-excel-file`, `xlsx` and `exceljs`.

Competitors are **optional** — both scripts skip any that aren't installed, so they are not part
of this package's dependencies.

## Write speed

### Run it

```bash
pnpm run bench
```

To include the comparison against ExcelJS, SheetJS and hucre, install them where the script can
import them (this edits your local `package.json`, so do not commit it):

```bash
pnpm add -D exceljs xlsx hucre
pnpm run bench
```

Each library is timed `RUNS` times (default `5`) and the median is reported. hucre declares Node >=24 in its `engines`.

Tune the dataset size with the `ROWS` env var (default `50000`):

```bash
ROWS=200000 pnpm run bench
```

### Reference results

Writing **50,000 rows × 10 columns** (mixed strings and numbers), median of 5 runs, three
invocations on an Apple M4, Node 24.19 (2026-10-08). Numbers vary by machine and by about 10%
between runs — run it yourself for your own hardware.

| Library | Write time | Output size |
| --- | ---: | ---: |
| **excel-bridge** (write) | **560 ms** | **2.41 MiB** |
| excel-bridge (stream) | 580 ms | 2.48 MiB |
| hucre (write) | 390 ms | 2.79 MiB |
| xlsx / SheetJS (write, no compression) | 510 ms | 18.23 MiB |
| xlsx / SheetJS (write, `compression: true`) | 580 ms | 6.26 MiB |
| exceljs (write) | 1430 ms | 2.82 MiB |

Takeaways for this workload:

- **~2.6× faster than ExcelJS** with a smaller file.
- hucre is faster (about 390 ms). SheetJS is on par: 510 ms with its default output, which is
  ~7.5× larger because it does not compress, and 580 ms with `compression: true`, ~2.6× larger.
- The streaming writer is not faster than the in-memory writer; it keeps memory lower (roughly
  250 MB against 375 MB peak resident memory for 50,000 rows, collecting the stream with
  `streamToBuffer`), and the gap grows with row count.

> Every library was driven with its documented defaults; no per-library tuning was
> applied. Treat these as directional, not absolute.

## Read speed

### Run it

```bash
pnpm run build
node --expose-gc benchmarks/bench-read.mjs
```

The script writes a file with excel-bridge, reads it `RUNS` times (default `5`) and reports the median and the peak memory of the process. Tune it with `ROWS` (default `50000`) and `RUNS`.

### Reference results

Reading a file excel-bridge wrote, 10 columns of mixed strings and numbers, on an Apple M4 with Node 24.19 (2026-10-08). The 1.6.0 column runs the same script against the published 1.6.0 package.

| Rows | 1.6.0 | 2.0 | Peak memory 1.6.0 | Peak memory 2.0 |
| --- | ---: | ---: | ---: | ---: |
| 50,000 (median of 7) | 1503 ms | 270 ms | 859 MiB | 451 MiB |
| 200,000 (median of 3) | 6521 ms | 1435 ms | 2679 MiB | 1611 MiB |

The peak memory is that of the whole process, which also builds the file it reads.

## Bundle size

Each row bundles a one-line entry such as `export { ExcelWriter } from 'excel-bridge'` with
esbuild (`--bundle --minify --platform=browser --format=esm`), then gzips the output with Node's
zlib at the default level. The bundler keeps that export and everything it references, which is
what the import costs a browser app.

### Run it

```bash
pnpm run size
```

To include the other libraries, install them in a scratch directory (not in this repository) and
point the script at it:

```bash
mkdir /tmp/competitors && cd /tmp/competitors && npm init -y
npm install hucre@1.2.0 xlsx@0.18.5 exceljs@4.4.0 @mitresthen/excelents@1.0.1 \
  read-excel-file@9.3.10 write-excel-file@4.1.1 --ignore-scripts
cd - && COMPETITORS_DIR=/tmp/competitors pnpm run size
```

The script prints a Markdown table, ready to paste into the main README.

To check the main README against the build without printing the table:

```bash
pnpm run build && pnpm run size:check
```

It measures only the excel-bridge entries and exits with an error when one differs by more than 100
bytes from the figure in the README bundle size table. CI runs it on every push and pull request.

### Reference results

The excel-bridge rows were measured 2026-10-08 and the other rows 2026-10-07, with esbuild 0.28.2 and Node 22; 1 KB = 1,000 bytes. The excel-bridge rows come from this repository's build with the dependency versions in `pnpm-lock.yaml`; the version in the first column is the release the build belongs to.

| Package | Import | Min | Min+gzip |
| --- | --- | ---: | ---: |
| excel-bridge@2.0.0 (dist) | `{ createExcelWorkbookStream }` | 30.5 KB | 12.3 KB |
| excel-bridge@2.0.0 (dist) | `{ ExcelWriter }` | 31.8 KB | 12.9 KB |
| excel-bridge@2.0.0 (dist) | `{ ExcelReader }` | 22.6 KB | 9.5 KB |
| excel-bridge@2.0.0 (dist) | `{ Workbook }` | 56.5 KB | 21.5 KB |
| excel-bridge@2.0.0 (dist) | `{ ExcelBridge }` | 56.9 KB | 21.7 KB |
| excel-bridge@2.0.0 (dist) | `* (everything)` | 66.6 KB | 25.4 KB |
| hucre@1.2.0 | `{ writeXlsx }` | 136.2 KB | 41.6 KB |
| hucre@1.2.0 | `{ readXlsx }` | 136.6 KB | 41.1 KB |
| hucre/xlsx@1.2.0 | `{ XlsxStreamWriter }` | 36.8 KB | 12.1 KB |
| hucre/xlsx@1.2.0 | `{ streamXlsxRows }` | 51.8 KB | 17.4 KB |
| @mitresthen/excelents@1.0.1 | `* (everything)` | 37.9 KB | 12.2 KB |
| read-excel-file/browser@9.3.10 | `{ default }` | 53.6 KB | 16.7 KB |
| write-excel-file/universal@4.1.1 | `{ default }` | 72.0 KB | 19.8 KB |
| xlsx@0.18.5 | `{ utils, write }` | 287.4 KB | 95.8 KB |
| exceljs@4.4.0 | `{ default }` | 947.0 KB | 272.1 KB |

Notes:

- `ExcelBridge` is a single object, and bundlers keep an object whole: any `ExcelBridge.*` call
  ships the reader and the writer, about as much as `Workbook`.
- hucre's own README quotes 34 KB for `readXlsx` and 68 KB for read plus write, measured with
  rolldown; those figures did not reproduce here with esbuild.
- `@mitresthen/excelents` is measured as a whole entry (reader and writer) and has no conditional
  formatting at 1.0.1, so it is not a like-for-like row for `ExcelWriter`.
- `exceljs` resolves to its prebuilt browser bundle (`dist/exceljs.min.js`), which can't be
  tree-shaken.
- `xlsx@0.18.5` is the SheetJS package on npm. Newer Community Edition builds ship from the SheetJS
  CDN and aren't measured here.
- gzip implementations differ by about 1% (macOS `gzip` comes out slightly smaller than Node's
  zlib), so compare numbers from the same run.
