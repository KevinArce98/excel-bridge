---
title: Define el diseño de las hojas
description: Varias hojas, celdas combinadas, paneles inmovilizados, anchos de columna, alturas de fila, filas y columnas ocultas, filtros e hipervínculos.
group: Guía
groupOrder: 2
order: 3
---

# Define el diseño de las hojas

Todo lo relativo a dónde se ubica cada elemento en una hoja vive en las `options` de la hoja y en algunos campos hermanos. Esta página cubre el conjunto completo, en el orden en que sueles necesitarlo.

## Varias hojas

Pasa un objeto por hoja. El nombre y la visibilidad de la hoja van en `options`.

```ts title="two-sheets.ts"
import { ExcelWriter } from 'excel-bridge';
import type { ExcelData } from 'excel-bridge';

const header = { background: '#4472C4', bold: true, color: '#FFFFFF' };

const sales: ExcelData = {
  data: [
    ['Product', 'Price', 'Quantity', 'Total'],
    ['Laptop', 999.99, 5, { formula: 'B2*C2' }],
    ['Mouse', 29.99, 20, { formula: 'B3*C3' }],
  ],
  styles: { '0-0': header, '0-1': header, '0-2': header, '0-3': header },
  options: { name: 'Sales Report', freezePane: { row: 1 }, autoWidth: true },
};

const timeline: ExcelData = {
  data: [
    ['Event', 'Date'],
    ['Launch', new Date(2024, 6, 15)],
  ],
  options: { name: 'Timeline', autoWidth: true },
};

const bytes = new ExcelWriter({ creator: 'My App' }).createWorkbookBuffer([sales, timeline]);
```

El nombre de una hoja debe tener de 1 a 31 caracteres, sin ninguno de `\ / ? * [ ] :`, sin caracteres de control, sin apóstrofo en ninguno de los extremos, no ser `History` (Excel lo reserva) y ser único sin distinguir mayúsculas de minúsculas. Un libro necesita al menos una hoja y al menos una hoja visible. De lo contrario, los escritores lanzan un error.

Oculta una hoja con `state: 'hidden'` o `state: 'veryHidden'`.

## Combinar celdas

Enumera los rangos en `mergeCells`. Pon el valor en la celda superior izquierda.

```ts title="merged-title.ts"
import type { ExcelData } from 'excel-bridge';

const quarterly: ExcelData = {
  data: [
    ['Q1 2024 Sales Report', '', '', ''],
    ['Product', 'January', 'February', 'March'],
    ['Laptops', 45000, 52000, 48000],
    ['Accessories', 12000, 15000, 13500],
  ],
  styles: { '0-0': { background: '#5B9BD5', bold: true, color: '#FFFFFF' } },
  mergeCells: ['A1:D1'],
  options: { name: 'Quarterly Report', freezePane: { row: 2 } },
};
```

## Paneles inmovilizados y anchos

`freezePane: { row: 2 }` mantiene visibles las dos primeras filas, y `col` inmoviliza columnas de la misma forma. Para los anchos, elige una de dos opciones:

- `autoWidth: true` mide el texto y ajusta el tamaño de cada columna, dentro de límites razonables.
- `columnWidths: [24, 12, 12, 12]` define anchos exactos en unidades de carácter. Tiene prioridad sobre `autoWidth`.

## Alturas de fila, filas ocultas y columnas ocultas

```ts title="row-layout.ts"
import type { ExcelData } from 'excel-bridge';

const layout: ExcelData = {
  data: [
    ['Title'],
    ['Row 2'],
    ['Row 3'],
    ['Row 4'],
  ],
  options: {
    rowHeights: { 0: 28 },
    hiddenRows: [2],
    hiddenColumns: [3],
  },
};
```

- `rowHeights` asocia índices de fila base cero con alturas en puntos, mayores que 0 y de 409.5 como máximo. Una fila sin altura conserva el valor predeterminado de Excel.
- `hiddenRows` y `hiddenColumns` enumeran índices base cero. Una columna oculta sin ancho se escribe con ancho `9.140625`.
- Las tres funcionan en el [escritor de streaming](../streaming/) y en un [`Workbook`](../workbook/) mediante `setRowHeight`, `setRowHidden` y `setColumnHidden`.

> [!LIMIT]
> Los valores predeterminados de la hoja, como la altura de fila predeterminada, el ancho de columna predeterminado y los niveles de esquema, no se escriben ni se conservan.

## Listas desplegables de filtro

`options.autoFilter` agrega listas desplegables de filtro a una fila de encabezado. Incluye también las filas de datos, tal como Excel guarda un filtro.

```ts title="filter.ts"
import type { ExcelData } from 'excel-bridge';

const rows = [
  ['Region', 'Rep', 'Revenue'],
  ['North', 'Ann', 1200],
  ['South', 'Bob', 480],
];

const filtered: ExcelData = {
  data: rows,
  options: { freezePane: { row: 1 }, autoFilter: { range: `A1:C${rows.length}` } },
};
```

Una hoja admite un solo filtro. El escritor también agrega el nombre oculto `_xlnm._FilterDatabase` que escribe Excel para un filtro. Los criterios de filtro y el estado de ordenación no se escriben ni se leen, así que un archivo escrito por `ExcelWriter` se abre con todas las filas visibles, a menos que pases `hiddenRows`.

## Hipervínculos

Adjunta vínculos mediante el arreglo `hyperlinks` de una hoja. La celda conserva el texto de `data`; el vínculo define a dónde lleva un clic. Los constructores `hyperlink` cubren los tres tipos.

```ts title="links.ts"
import { ExcelWriter, hyperlink } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Resource', 'Contact', 'Details'],
      ['Docs', 'Email the team', 'See Q1'],
    ],
    hyperlinks: [
      hyperlink.url('A2', 'https://example.com/docs', { tooltip: 'Open the docs' }),
      hyperlink.email('B2', 'team@example.com', { subject: 'Report question' }),
      hyperlink.internal('C2', 'Q1 Sales', 'A1'),
    ],
    options: { name: 'Links' },
  },
  { data: [['Q1']], options: { name: 'Q1 Sales' } },
]);
```

Un vínculo también puede ser un objeto simple: `{ range: 'A2', url: 'https://…' }` o `{ range: 'C2', location: "'Q1 Sales'!A1" }`, con `tooltip` y `display` opcionales. Se descarta un `#` inicial en `location`.

- **Aspecto:** la primera celda de cada vínculo recibe el color de hipervínculo de Excel (`#0563C1`) y un subrayado. Un `color` o un `underline` en el estilo propio de esa celda prevalece, cada uno por separado.
- **URL permitidas:** `http:`, `https:` y `mailto:`, con los espacios y las comillas codificados en porcentaje. Cualquier otra cosa lanza un error, porque las exportaciones suelen llevar datos de usuarios y los vínculos `file:` o los controladores de protocolo personalizados no son seguros de abrir con un clic.
- **Límites:** un vínculo por rango, 65,530 vínculos por hoja y 2,079 caracteres por dirección.
