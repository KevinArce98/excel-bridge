---
title: Escribe valores y fórmulas
description: Qué puede contener una celda, cómo se escriben las fórmulas, el texto literal y los errores, y qué comprueba la biblioteca antes de escribir.
group: Guía
groupOrder: 2
order: 1
---

# Escribe valores y fórmulas

Una hoja es un arreglo de filas, y una fila es un arreglo de celdas. La mayoría de las celdas son valores simples. Unas pocas necesitan un objeto, porque expresan algo que un valor simple no puede.

## Valores simples

| Escribes | La celda contiene |
| --- | --- |
| `'Laptop'` | Texto. Una cadena siempre es texto, incluso si empieza con `=`. |
| `999.99` | Un número. Debe ser finito. |
| `true` | Un booleano. |
| `new Date(2024, 0, 15)` | Una fecha, guardada como número de serie de Excel. |
| `null` o `undefined` | Una celda vacía. |

```ts title="plain-values.ts"
import { ExcelWriter } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Item', 'Price', 'In stock', 'Added', 'Note'],
      ['Laptop', 999.99, true, new Date(2024, 0, 15), null],
      ['=not a formula', 29.99, false, new Date(2024, 1, 2), 'plain text'],
    ],
  },
]);
```

Como las cadenas siempre son texto, exportar datos de usuarios no puede crear una fórmula por accidente. Un nombre como `=HYPERLINK("https://example.com")` se escribe como el texto que es.

## Fórmulas

Escribe una fórmula con un objeto. Omite el `=` inicial.

```ts title="formulas.ts"
import { ExcelWriter } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Product', 'Price', 'Quantity', 'Total'],
      ['Laptop', 999.99, 5, { formula: 'B2*C2', result: 4999.95 }],
      ['Mouse', 29.99, 20, { formula: 'B3*C3' }],
      ['', '', 'Sum', { formula: 'SUM(D2:D3)' }],
    ],
  },
]);
```

- **`{ formula }`** escribe una fórmula sin valor almacenado. El libro le pide a Excel que recalcule todas las fórmulas al abrir el archivo.
- **`{ formula, result }`** también guarda el último valor que tuvo la fórmula. `result` es una cadena, un número, un booleano, un `Date` o un `{ error }`.

La biblioteca nunca calcula una fórmula. Un lector que no calcula, como un visor de vista previa o un script, muestra `result` cuando lo proporcionas y una celda vacía cuando no.

> [!WARNING]
> Nunca pases entrada no confiable a `{ formula }`. Una fórmula se ejecuta cuando se abre el archivo. El texto del usuario va en una cadena simple.

## Texto literal y errores

`{ text }` escribe texto, exactamente igual que una cadena simple. Úsalo cuando el mismo código también deba ejecutarse en excel-bridge 1.6, donde una cadena simple que empieza con `=` es una fórmula. Antes de 1.6 una celda con objeto se escribe como `[object Object]`.

`{ error }` escribe un valor de error. Los siete errores clásicos son `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!` y `#N/A`. Cualquier otro texto de error lanza un error.

```ts title="text-and-errors.ts"
import type { CellValue } from 'excel-bridge';

const row: CellValue[] = [{ text: '=SUM(A1:A3)' }, { error: '#N/A' }, { formula: '1/0', result: { error: '#DIV/0!' } }];
```

## Fechas

Un `Date` se escribe como hora local de reloj, así que `new Date(2024, 0, 15)` es el 15 de enero en cualquier zona horaria. Dos funciones auxiliares convierten entre fechas y números de serie de Excel: `dateToExcelSerial(date)` y `excelSerialToDate(serial)`.

> [!LIMIT]
> Una fecha a medianoche UTC, como `new Date('2024-01-15')`, cae en el día anterior con desfases negativos. Una hora local que no existe, dentro de un salto por horario de verano, se adelanta.

Para mostrar la hora del día u otro formato, asigna a la celda un `numberFormat` de fecha. Consulta [Da estilo a las celdas](../styling/).

## Qué comprueba el escritor

Los escritores lanzan un [`ExcelBridgeError`](../errors/) con el código `INVALID_INPUT` para las entradas siguientes.

| Entrada | Resultado |
| --- | --- |
| `NaN`, `Infinity` o un `Date` no válido | Lanza un error y nombra la celda, como `Cell B2 holds NaN, which a worksheet cannot store`. |
| Una cadena de más de 32,767 caracteres | Lanza un error. |
| `{ formula: '' }` | Lanza un error. |
| Un valor de error fuera de los siete | Lanza un error. |
| Caracteres de control distintos de tabulación, salto de línea y retorno de carro | Se eliminan sin avisar. |

## Cadenas y tamaño del archivo

Las cadenas se escriben en línea de forma predeterminada, lo cual es simple y confiable. Para un libro con muchas cadenas repetidas, activa una tabla de cadenas compartidas para reducir el tamaño del archivo.

```ts title="shared-strings.ts"
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ sharedStrings: true });
const bytes = writer.createWorkbookBuffer([
  { data: [['Status'], ['Open'], ['Open'], ['Done'], ['Open']] },
]);
```

El [escritor de streaming](../streaming/) ignora esta opción y siempre escribe cadenas en línea.
