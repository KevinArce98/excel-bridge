---
title: Stream big exports
description: Write million-row files without holding them in memory, and send them to a browser, an Express handler or a fetch-style server.
group: Guide
groupOrder: 2
order: 8
---

# Stream big exports

For exports too large to hold in memory, `createExcelWorkbookStream` yields the `.xlsx` as `Uint8Array` chunks. Rows come from a sync or async iterable, so they never all live in memory at once.

```ts title="stream-to-disk.ts"
import fs from 'node:fs';
import { createExcelWorkbookStream, streamToBuffer } from 'excel-bridge';

function* generateRows() {
  yield ['Id', 'Name', 'Value'];
  for (let id = 1; id <= 1_000; id++) yield [id, `Row ${id}`, id * 2];
}

const out = fs.createWriteStream('big.xlsx');
for await (const chunk of createExcelWorkbookStream([{ name: 'Data', rows: generateRows() }])) {
  out.write(chunk);
}
out.end();

const bytes = await streamToBuffer(createExcelWorkbookStream([{ name: 'Data', rows: generateRows() }]));
```

Pipe the chunks straight to disk, or collect them with `streamToBuffer` when you still need one buffer. The example makes 1,000 rows. Change the loop to `1_000_000` for a million: the rows are produced one at a time.

## What a streaming sheet supports

A `StreamingSheetInput` takes `name`, `rows`, `styles`, `freezePane`, `columnWidths`, `rowHeights`, `hiddenRows`, `hiddenColumns`, `mergeCells`, `autoFilter` and `hyperlinks`. Pass the final filter range and links up front: they are validated, and the range is written to the workbook part, before the first row.

It does not support `autoWidth`, `validations`, `conditionalFormats`, `sharedStrings` or a sheet `state`. Use `ExcelWriter` or `Workbook` when you need those.

To stream a list of objects with a header row, see [Rows as objects](../objects/).

## Send a file to the browser

`downloadXlsx(data, filename?)` starts a download from a `Blob` or a `Uint8Array`. It works only where there is a `document`, and it throws an `UNSUPPORTED` error elsewhere. Nothing is touched when you import the package.

```ts title="download.ts" run=false
import { ExcelWriter, downloadXlsx } from 'excel-bridge';

const sheet = { data: [['Product', 'Price'], ['Laptop', 999.99]] };

document.querySelector('button')!.addEventListener('click', () => {
  downloadXlsx(new ExcelWriter().createWorkbook([sheet]), 'Sales 2024');
});
```

The file name is made safe for you: path separators, `: * ? " < > |` and control characters become `_` (a run becomes one `_`), and `.xlsx` is appended once.

## Send a file from a server

For fetch-style servers (Cloudflare Workers, Deno, Bun, Next.js route handlers, Hono), `xlsxResponse(body, filename?, init?)` returns a `Response` with the right `Content-Type` and a `Content-Disposition` that handles non-ASCII names. Give it the stream and it waits for the first chunk, so anything the writer checks up front, such as sheet names and styles, rejects before the response exists and you can still answer with an error status.

```ts title="route-handler.ts" run=false
import { createExcelWorkbookStream, xlsxResponse } from 'excel-bridge';

declare function loadRows(): AsyncIterable<(string | number)[]>;

export async function GET() {
  return xlsxResponse(createExcelWorkbookStream([{ name: 'Data', rows: loadRows() }]), 'report');
}
```

> [!NOTE]
> Once the response has started, a failure while later rows are read cannot change the status. The body errors and the client sees an aborted download.

`toReadableStream(chunks)` turns any async iterable of `Uint8Array` into a `ReadableStream`, if you build the `Response` yourself.

### Express and plain Node

The streaming writer yields `Uint8Array` chunks, which `Readable.from` accepts.

```ts title="express.ts" check=false
import express from 'express';
import { Readable } from 'node:stream';
import { XLSX_CONTENT_TYPE, createExcelWorkbookStream } from 'excel-bridge';

declare function loadRows(): AsyncIterable<(string | number)[]>;

const app = express();

app.get('/report.xlsx', (req, res) => {
  res.setHeader('Content-Type', XLSX_CONTENT_TYPE);
  res.attachment('report.xlsx');
  Readable.from(createExcelWorkbookStream([{ name: 'Data', rows: loadRows() }])).pipe(res);
});
```

> [!NOTE]
> The writers return a plain `Uint8Array`, not a `Buffer`. With a buffer in Express 4, use `res.send(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength))`, because Express 4 sends a `Uint8Array` as JSON.
