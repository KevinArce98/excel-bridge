---
title: Exporta archivos grandes con streaming
description: Escribe archivos de un millón de filas sin mantenerlos en memoria y envíalos a un navegador, a un handler de Express o a un servidor de estilo fetch.
group: Guía
groupOrder: 2
order: 8
---

# Exporta archivos grandes con streaming

Para exportaciones demasiado grandes para mantenerlas en memoria, `createExcelWorkbookStream` produce el `.xlsx` en fragmentos `Uint8Array`. Las filas provienen de un iterable síncrono o asíncrono, así que nunca están todas en memoria a la vez.

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

Envía los fragmentos directamente al disco, o recógelos con `streamToBuffer` cuando aún necesites un solo buffer. El ejemplo genera 1,000 filas. Cambia el bucle a `1_000_000` para un millón: las filas se producen una a la vez.

## Qué admite una hoja de streaming

Un `StreamingSheetInput` acepta `name`, `rows`, `styles`, `freezePane`, `columnWidths`, `rowHeights`, `hiddenRows`, `hiddenColumns`, `mergeCells`, `autoFilter` y `hyperlinks`. Pasa desde el principio el rango final del filtro y los vínculos: se validan, y el rango se escribe en la parte del libro, antes de la primera fila.

No admite `autoWidth`, `validations`, `conditionalFormats`, `sharedStrings` ni un `state` de hoja. Usa `ExcelWriter` o `Workbook` cuando los necesites.

Para enviar por streaming una lista de objetos con una fila de encabezado, consulta [Filas como objetos](../objects/).

## Enviar un archivo al navegador

`downloadXlsx(data, filename?)` inicia una descarga a partir de un `Blob` o un `Uint8Array`. Solo funciona donde hay un `document`, y en otros entornos lanza un error `UNSUPPORTED`. No se toca nada al importar el paquete.

```ts title="download.ts" run=false
import { ExcelWriter, downloadXlsx } from 'excel-bridge';

const sheet = { data: [['Product', 'Price'], ['Laptop', 999.99]] };

document.querySelector('button')!.addEventListener('click', () => {
  downloadXlsx(new ExcelWriter().createWorkbook([sheet]), 'Sales 2024');
});
```

El nombre del archivo se hace seguro por ti: los separadores de ruta, `: * ? " < > |` y los caracteres de control se convierten en `_` (una secuencia se convierte en un solo `_`), y `.xlsx` se agrega una sola vez.

## Enviar un archivo desde un servidor

Para servidores de estilo fetch (Cloudflare Workers, Deno, Bun, route handlers de Next.js, Hono), `xlsxResponse(body, filename?, init?)` devuelve una `Response` con el `Content-Type` correcto y un `Content-Disposition` que maneja nombres con caracteres no ASCII. Pásale el stream y espera el primer fragmento, así que todo lo que el escritor comprueba de antemano, como los nombres de hoja y los estilos, se rechaza antes de que exista la respuesta y aún puedes responder con un código de estado de error.

```ts title="route-handler.ts" run=false
import { createExcelWorkbookStream, xlsxResponse } from 'excel-bridge';

declare function loadRows(): AsyncIterable<(string | number)[]>;

export async function GET() {
  return xlsxResponse(createExcelWorkbookStream([{ name: 'Data', rows: loadRows() }]), 'report');
}
```

> [!NOTE]
> Una vez que la respuesta ha comenzado, un fallo al leer filas posteriores no puede cambiar el código de estado. El cuerpo falla y el cliente ve una descarga interrumpida.

`toReadableStream(chunks)` convierte cualquier iterable asíncrono de `Uint8Array` en un `ReadableStream`, por si construyes tú la `Response`.

### Express y Node puro

El escritor de streaming produce fragmentos `Uint8Array`, que `Readable.from` acepta.

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
> Los escritores devuelven un `Uint8Array` simple, no un `Buffer`. Con un buffer en Express 4, usa `res.send(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength))`, porque Express 4 envía un `Uint8Array` como JSON.
