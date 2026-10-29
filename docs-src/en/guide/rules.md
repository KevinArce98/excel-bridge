---
title: Add rules
description: Conditional formatting that reacts to values and formulas, and data validation that limits what people can type.
group: Guide
groupOrder: 2
order: 4
---

# Add rules

Two features change how a sheet behaves after it is written: conditional formatting restyles cells by what they hold, and data validation limits what can be typed into them.

## Conditional formatting

Attach `conditionalFormats` to a sheet to highlight cells by value, by formula, or with a colour scale.

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

| Type | Fields | What it does |
| --- | --- | --- |
| `cellValue` | `operator`, `value`, `value2?`, `style` | Compares each cell with a value. `between` and `notBetween` use `value2`. |
| `expression` | `formula`, `style` | Styles the range where the formula is true. |
| `colorScale` | `colors` (two or three) | Shades numbers from the lowest to the highest value. |

Operators are `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `between` and `notBetween`. A rule's `style` takes `background`, `color`, `bold` and `italic`.

> [!LIMIT]
> A `Workbook` keeps `cellValue`, `expression` and `colorScale` rules when it loads a file and drops every other rule type. The [streaming writer](../streaming/) does not write conditional formats.

## Data validation

Use the typed `dataValidation` builders instead of writing raw rule strings. Each returns a `CellValidation` for a sheet's `validations` array, or for `Workbook.addValidation`.

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

The builders are `list(range, values)`, `wholeNumber`, `decimal` and `textLength` (each `(range, operator, value, value2?)`), and `dateBetween(range, start, end)`. The operators are the same eight as above.

Rules of type `time` and `custom` are written from plain objects: `{ range, type: 'custom', formula1: 'ISNUMBER(A1)' }`.

> [!NOTE]
> Every saved rule is written to show both the input and the error message and to reject bad entries. The message text and the error style (stop, warning, information) are not written or kept.
