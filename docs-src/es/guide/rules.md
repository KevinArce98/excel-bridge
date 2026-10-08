---
title: Agrega reglas
description: Formato condicional que reacciona a valores y fórmulas, y validación de datos que limita lo que las personas pueden escribir.
group: Guía
groupOrder: 2
order: 4
---

# Agrega reglas

Dos funciones cambian el comportamiento de una hoja después de escribirla: el formato condicional cambia el estilo de las celdas según lo que contienen, y la validación de datos limita lo que se puede escribir en ellas.

## Formato condicional

Agrega `conditionalFormats` a una hoja para resaltar celdas por valor, por fórmula o con una escala de color.

```ts title="conditional-formats.ts"
import { ExcelWriter } from 'excel-bridge';
import type { ConditionalFormat } from 'excel-bridge';

const conditionalFormats: ConditionalFormat[] = [
  {
    type: 'cellValue',
    range: 'C2:C4',
    operator: 'lessThan',
    value: 10,
    style: { background: '#FFC7CE', color: '#9C0006' },
  },
  {
    type: 'expression',
    range: 'A2:D4',
    formula: '$D2="Out of Stock"',
    style: { background: '#FFE6E6', bold: true },
  },
  {
    type: 'colorScale',
    range: 'B2:B4',
    colors: ['#F8696B', '#FFEB84', '#63BE7B'],
  },
];

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Product', 'Price', 'Stock', 'Status'],
      ['Laptop', 999.99, 15, 'Available'],
      ['Mouse', 29.99, 5, 'Low Stock'],
      ['Keyboard', 79.99, 0, 'Out of Stock'],
    ],
    conditionalFormats,
  },
]);
```

| Tipo | Campos | Qué hace |
| --- | --- | --- |
| `cellValue` | `operator`, `value`, `value2?`, `style` | Compara cada celda con un valor. `between` y `notBetween` usan `value2`. |
| `expression` | `formula`, `style` | Aplica estilo al rango donde la fórmula es verdadera. |
| `colorScale` | `colors` (dos o tres) | Sombrea los números desde el valor más bajo hasta el más alto. |

Los operadores son `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `between` y `notBetween`. El `style` de una regla acepta `background`, `color`, `bold` e `italic`.

> [!LIMIT]
> Un `Workbook` conserva las reglas `cellValue`, `expression` y `colorScale` cuando carga un archivo y descarta todos los demás tipos de regla. El [escritor de streaming](../streaming/) no escribe formatos condicionales.

## Validación de datos

Usa los constructores con tipos de `dataValidation` en lugar de escribir cadenas de reglas sin procesar. Cada uno devuelve un `CellValidation` para el arreglo `validations` de una hoja, o para `Workbook.addValidation`.

```ts title="validation.ts"
import { ExcelWriter, dataValidation } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [['Status', 'Priority', 'Score', 'Due']],
    validations: [
      dataValidation.list('A2:A100', ['Open', 'In Progress', 'Done']),
      dataValidation.wholeNumber('B2:B100', 'between', 1, 5),
      dataValidation.decimal('C2:C100', 'greaterThanOrEqual', 0),
      dataValidation.dateBetween('D2:D100', new Date(2024, 0, 1), new Date(2024, 11, 31)),
    ],
  },
]);
```

Los constructores son `list(range, values)`, `wholeNumber`, `decimal` y `textLength` (cada uno `(range, operator, value, value2?)`), y `dateBetween(range, start, end)`. Los operadores son los mismos ocho de arriba.

Las reglas de tipo `time` y `custom` se escriben a partir de objetos simples: `{ range, type: 'custom', formula1: 'ISNUMBER(A1)' }`.

> [!NOTE]
> Toda regla guardada se escribe para mostrar tanto el mensaje de entrada como el de error y para rechazar las entradas incorrectas. El texto de los mensajes y el estilo de error (detener, advertencia, información) no se escriben ni se conservan.
