---
title: Maneja errores y límites
description: Ramifica según los códigos de error en lugar de los mensajes, y define límites del lector para archivos que no creaste tú.
group: Guía
groupOrder: 2
order: 9
---

# Maneja errores y límites

Todo error que la biblioteca lanza a propósito es un `ExcelBridgeError`. Lleva un `code`, así que puedes decidir qué hacer sin leer el mensaje.

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

## Los cuatro códigos

| Código | Significado | Un servidor podría responder |
| --- | --- | --- |
| `INVALID_INPUT` | Pasaste un valor que la biblioteca rechaza: una celda `NaN`, un nombre de hoja incorrecto, un color incorrecto, una fila fuera de la cuadrícula. Corrige la llamada o los datos. | 500 si tu código armó el valor, 422 si lo envió un usuario |
| `INVALID_FILE` | Los bytes no son lo que se pidió: no es un zip, falta la parte del libro, XML que no se puede analizar, una celda fuera de la cuadrícula de Excel, o una columna que exigiste y que la hoja no tiene. | 400 o 422 |
| `LIMIT_EXCEEDED` | Se alcanzó un límite del lector, también uno predeterminado. `error.limit` lo nombra. | 413 |
| `UNSUPPORTED` | Al entorno le falta lo que necesita la llamada, por ejemplo `downloadXlsx` sin un `document`. | 500 |

Los límites propios del formato de Excel son `INVALID_INPUT` cuando escribes, por ejemplo una fila posterior a la 1,048,576 o más de 32,767 caracteres de texto en una celda. Cuando lees, una fila o columna fuera de la cuadrícula es `INVALID_FILE`. El lector no comprueba la longitud del texto de una celda, así que `Workbook` lanza `INVALID_INPUT` al guardar una celda así. No son `LIMIT_EXCEEDED`, que corresponde a los límites del lector descritos abajo, así que un servidor que responde 413 ante `LIMIT_EXCEEDED` no responde 413 ante un archivo corrupto.

Se pueden agregar códigos nuevos en una versión menor, así que agrega un `default` a todo `switch` sobre `error.code`.

> [!NOTE]
> Usa `isExcelBridgeError(error)` o `error.code`, no `instanceof`. Las compilaciones ESM y CommonJS del paquete son copias separadas de la clase, así que `instanceof` puede dar falso entre ellas.

Un error de `parseFromBuffer` conserva el prefijo de su mensaje, `Failed to parse Excel file:`, y el error original está en `cause`.

## Límites del lector

`ExcelReader`, `parseExcel`, `ExcelBridge.read` y `Workbook.fromBuffer` aceptan las mismas opciones. Defínelas para cualquier archivo que no hayas creado tú.

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

| Opción | Qué cuenta | Valor predeterminado |
| --- | --- | --- |
| `maxCells` | Todas las celdas que el lector crea en el libro, incluidas las celdas vacías que rellenan las filas cortas. | 5,000,000 |
| `maxPartBytes` | El tamaño sin comprimir declarado de cualquier parte del archivo. Una hoja de cálculo es una parte, así que este es el límite del XML de la hoja. | 256 MiB |
| `maxTotalBytes` | La suma de los tamaños declarados de todas las partes que se descomprimen. | 512 MiB |
| `maxSheets` | Las hojas enumeradas en el libro. | 1,000 |

Pasa `Infinity` para desactivar un límite. Los límites de tamaño se comprueban a partir del directorio del archivo antes de descomprimir nada, así que un archivo que declara 2 GiB para una parte se rechaza sin reservar memoria para ella.

## Archivos que no creaste tú

> [!WARNING]
> Estos límites no hacen seguro un archivo no confiable. Una hoja de 4.5 millones de celdas (205 MiB de XML) pasa todos los valores predeterminados. Leerla tardó de 3 a 5 s, y el proceso llegó a entre 2.1 y 2.5 GB de memoria residente en ejecuciones repetidas en Node 22 con su heap predeterminado: unas 10 a 12 veces el XML de la hoja. Un encabezado zip que subestima el tamaño de una parte puede truncarla, y el lector suele informarlo como un archivo no válido.

Limita el tamaño de la carga antes de analizar, define `maxCells` y `maxPartBytes` según lo que necesite tu producto y, para cargas públicas, analiza en un worker o en un proceso separado que se pueda reiniciar.
