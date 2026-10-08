---
title: Handle errors and limits
description: Branch on error codes instead of messages, and set reader limits for files you did not make.
group: Guide
groupOrder: 2
order: 9
---

# Handle errors and limits

Every error the library raises on purpose is an `ExcelBridgeError`. It carries a `code`, so you can decide what to do without reading the message.

```ts title="catch.ts"
import { ExcelReader, ExcelWriter, isExcelBridgeError } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([{ data: [['a', 1], ['b', 2], ['c', 3]] }]);

try {
  new ExcelReader({ maxCells: 2 }).parseFromBuffer(bytes);
} catch (error) {
  if (isExcelBridgeError(error)) {
    console.log(error.code, error.limit, error.message);
  } else {
    throw error;
  }
}
```

## The four codes

| Code | Meaning | A server might answer |
| --- | --- | --- |
| `INVALID_INPUT` | You passed a value the library rejects: a `NaN` cell, a bad sheet name, a bad colour, a row beyond the grid. Fix the call or the data. | 500 if your code built the value, 422 if a user supplied it |
| `INVALID_FILE` | The bytes are not what was asked for: not a zip, no workbook part, XML that does not parse, a cell outside Excel's grid, or a column you required that the sheet lacks. | 400 or 422 |
| `LIMIT_EXCEEDED` | A reader limit was hit, a default one included. `error.limit` names it. | 413 |
| `UNSUPPORTED` | The environment lacks what the call needs, for example `downloadXlsx` without a `document`. | 500 |

Excel's own format limits are `INVALID_INPUT` when you write, for example a row past 1,048,576 or text over 32,767 characters in a cell. When you read, a row or column outside the grid is `INVALID_FILE`. The reader does not check the length of cell text, so `Workbook` throws `INVALID_INPUT` when it saves such a cell. They are not `LIMIT_EXCEEDED`, which is for the reader limits below, so a server that answers 413 for `LIMIT_EXCEEDED` does not answer 413 for a corrupt file.

New codes can be added in a minor release, so give a `switch` over `error.code` a `default`.

> [!NOTE]
> Use `isExcelBridgeError(error)` or `error.code`, not `instanceof`. The ESM and CommonJS builds of the package are separate copies of the class, so `instanceof` can be false across them.

An error from `parseFromBuffer` keeps its message prefix, `Failed to parse Excel file:`, and the original error is on `cause`.

## Reader limits

`ExcelReader`, `parseExcel`, `ExcelBridge.read` and `Workbook.fromBuffer` take the same options. Set them for any file you did not make.

```ts title="limits.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const reader = new ExcelReader({
  maxCells: 1_000_000,
  maxPartBytes: 64 * 1024 * 1024,
  maxTotalBytes: 128 * 1024 * 1024,
  maxSheets: 20,
});

const workbook = reader.parseFromBuffer(fs.readFileSync('data.xlsx'));
```

| Option | Counts | Default |
| --- | --- | --- |
| `maxCells` | Every cell the reader creates in the workbook, including the empty cells that pad short rows. | 5,000,000 |
| `maxPartBytes` | The declared uncompressed size of any one part of the file. A worksheet is a part, so this is the sheet XML limit. | 256 MiB |
| `maxTotalBytes` | The sum of the declared sizes of every part that is inflated. | 512 MiB |
| `maxSheets` | The sheets listed in the workbook. | 1,000 |

Pass `Infinity` to turn a limit off. The size limits are checked from the file's directory before anything is inflated, so a file that declares 2 GiB for one part is refused without allocating it.

## Files you did not make

> [!WARNING]
> These limits do not make an untrusted file safe. A sheet of 4.5 million cells (205 MiB of XML) passes every default. Reading it took 3 to 5 s, and the process peaked at 2.1 to 2.5 GB of resident memory in repeated runs on Node 22 with its default heap: about 10 to 12 times the sheet XML. A zip header that understates a part's size can cut the part short, which the reader usually reports as an invalid file.

Cap the upload size before you parse, set `maxCells` and `maxPartBytes` to what your product needs, and for public uploads parse in a worker or a separate process that can be restarted.
