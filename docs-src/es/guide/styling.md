---
title: Da estilo a las celdas
description: Color, fuentes, formatos de número, alineación y bordes por lado para cualquier celda, con la fila y la columna como clave.
group: Guía
groupOrder: 2
order: 2
---

# Da estilo a las celdas

Dale a cualquier celda un relleno, una fuente, una alineación, un formato de número o un borde. Los estilos viven junto a los datos, con la clave `"<row>-<col>"` y índices que empiezan en cero, así que puedes construirlos a partir de los propios datos.

```ts title="header-row.ts"
import { ExcelWriter } from 'excel-bridge';
import type { ExcelData } from 'excel-bridge';

const header = { background: '#4472C4', bold: true, color: '#FFFFFF' };

const sheet: ExcelData = {
  data: [
    ['Product', 'Price', 'Stock', 'Status'],
    ['Laptop', 999.99, 15, 'Available'],
    ['Mouse', 29.99, 5, 'Low stock'],
    ['Keyboard', 79.99, 0, 'Out of stock'],
  ],
  styles: {
    '0-0': header,
    '0-1': header,
    '0-2': header,
    '0-3': header,
    '1-2': { background: '#E2EFDA', color: '#006100' },
    '2-2': { background: '#FFC7CE', color: '#9C0006' },
    '3-3': { background: '#FFE6E6', color: '#C00000' },
  },
  options: { name: 'Inventory', freezePane: { row: 1 }, autoWidth: true },
};

const bytes = new ExcelWriter().createWorkbookBuffer([sheet]);
```

> [!TIP]
> Anota la hoja como `ExcelData`. Sin eso, TypeScript amplía literales como `align: 'center'` a `string` y rechaza el objeto.

## Qué puede contener un estilo

| Propiedad | Valor | Efecto |
| --- | --- | --- |
| `background` | color | Relleno sólido. |
| `color` | color | Color de la fuente. |
| `bold`, `italic`, `underline` | `boolean` | Peso, inclinación y subrayado de la fuente. |
| `fontSize` | `number` | Puntos. Predeterminado: 11. |
| `fontName` | `string` | Predeterminado: `Calibri`. |
| `align` | `'left' \| 'center' \| 'right'` | Alineación horizontal. |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | Alineación vertical. |
| `wrapText` | `boolean` | Ajusta el texto largo dentro de la celda. |
| `numberFormat` | código de formato | Por ejemplo `'#,##0.00'` o `'0%'`. |
| `border` | `boolean`, un estilo de línea o entradas por lado | Consulta [Bordes](#bordes). |

Los colores son hexadecimales: `#RGB`, `#RRGGBB` o `#AARRGGBB`, con o sin el `#`. Cualquier otro valor, como `red`, lanza un error en lugar de escribir un valor no válido.

```ts title="invoice.ts"
import type { ExcelData } from 'excel-bridge';

const invoice: ExcelData = {
  data: [
    ['Invoice', 1250.5],
    ['Tax', 237.6],
  ],
  styles: {
    '0-0': { bold: true, italic: true, fontSize: 14, fontName: 'Arial', align: 'center', wrapText: true },
    '0-1': { numberFormat: '#,##0.00' },
    '1-1': { numberFormat: '#,##0.00', verticalAlign: 'middle' },
  },
};
```

## Las fechas también admiten estilos

Una celda `Date` usa el formato de fecha corta integrado. Dale a su estilo un `numberFormat` de fecha para mostrar la hora del día u otro formato. Un `numberFormat` que no es de fecha se ignora en una celda `Date`.

```ts title="dates.ts"
import type { ExcelData } from 'excel-bridge';

const events: ExcelData = {
  data: [['Launch', new Date(2024, 0, 15, 13, 45)]],
  styles: { '0-1': { bold: true, numberFormat: 'yyyy-mm-dd hh:mm' } },
};
```

> [!LIMIT]
> Un `Date` se escribe como hora local de reloj: `new Date(2024, 0, 15)` es el 15 de enero en cualquier zona horaria. Una fecha a medianoche UTC, como `new Date('2024-01-15')`, cae en el día anterior con desfases negativos, y una hora local que no existe (dentro de un salto por horario de verano) se adelanta.

## Bordes

`border: true` dibuja un recuadro negro delgado en los cuatro lados. Para cualquier otra cosa, pasa un estilo de línea, un par `{ style, color? }` o un objeto con una entrada por lado.

```ts title="borders.ts"
import type { ExcelData } from 'excel-bridge';

const report: ExcelData = {
  data: [
    ['Region', 'Revenue'],
    ['North', 1200],
    ['South', 480],
  ],
  styles: {
    '0-0': { bold: true, border: { bottom: 'medium' } },
    '0-1': { bold: true, border: { bottom: 'medium' } },
    '1-1': { border: { left: { style: 'thin', color: '#CCCCCC' } } },
    '2-0': { border: 'thin' },
    '2-1': { border: { top: 'thin', bottom: { style: 'double', color: '#CC0000' } } },
  },
};
```

- **Valores aceptados:** `true` o `false`; un estilo de línea para los cuatro lados; `{ style, color? }` para los cuatro lados; o `{ left?, right?, top?, bottom? }`, cada uno con un estilo de línea o `{ style, color? }`.
- **Estilos de línea (13):** `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`.
- **Colores:** `#RGB`, `#RRGGBB` o `#AARRGGBB`. Sin color, la línea es negra.
- **Errores:** un estilo de línea desconocido o un color como `red` lanza un error y nombra el valor.
- **Celdas combinadas:** el archivo guarda los bordes por celda, y el escritor no copia un borde a lo largo de un rango combinado. Define el lado en cada celda del borde y rellena los datos con `null` para que esas celdas existan. Un estilo en una celda más allá del final de su fila se ignora.

> [!LIMIT]
> No hay bordes diagonales ni bordes dentro de los formatos condicionales.

## Formato condicional sin un estilo por celda

Resalta celdas por valor o por fórmula, o con una escala de color, usando [reglas](../rules/) en lugar de un estilo por celda.
