# Benchmarks

Compares `excel-bridge` against `exceljs` and `xlsx` (SheetJS) for writing a
large workbook. Competitors are **optional** — the benchmark skips any that
aren't installed, so they are not part of this package's dependencies.

## Run it

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

## Reference results

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
