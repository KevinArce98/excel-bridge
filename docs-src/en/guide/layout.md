---
title: Lay out sheets
description: Several sheets, merged cells, frozen panes, column widths, row heights, hidden rows and columns, filters and links.
group: Guide
groupOrder: 2
order: 3
---

# Lay out sheets

Everything about where things sit on a sheet lives in the sheet's `options` and a few sibling fields. This page covers the whole set, in the order you usually need it.

## Several sheets

Pass one object per sheet. The sheet's name and visibility go in `options`.

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

A sheet name must be 1 to 31 characters, with none of `\ / ? * [ ] :` and no control characters, no apostrophe at either end, not `History` (Excel reserves it) and unique when case is ignored. A workbook needs at least one sheet and at least one visible sheet. The writers throw otherwise.

Hide a sheet with `state: 'hidden'` or `state: 'veryHidden'`.

## Merge cells

List ranges in `mergeCells`. Put the value in the top-left cell.

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

## Freeze panes and widths

`freezePane: { row: 2 }` keeps the first two rows in view, and `col` freezes columns the same way. For widths, choose one of two options:

- `autoWidth: true` measures the text and sizes each column, within sensible bounds.
- `columnWidths: [24, 12, 12, 12]` sets exact widths in character units. It takes precedence over `autoWidth`.

## Row heights, hidden rows and hidden columns

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

- `rowHeights` maps zero-based row indexes to heights in points, above 0 and at most 409.5. A row without a height keeps Excel's default.
- `hiddenRows` and `hiddenColumns` list zero-based indexes. A hidden column without a width is written with width `9.140625`.
- All three work in the [streaming writer](../streaming/) and on a [`Workbook`](../workbook/) through `setRowHeight`, `setRowHidden` and `setColumnHidden`.

> [!LIMIT]
> Sheet defaults, such as the default row height, the default column width and outline levels, are not written or kept.

## Filter dropdowns

`options.autoFilter` adds filter dropdowns to a header row. Cover the data rows too, the way Excel saves a filter.

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

A sheet holds one filter. The writer also adds the hidden `_xlnm._FilterDatabase` name that Excel writes for a filter. Filter criteria and sort state are not written or read, so a file written by `ExcelWriter` opens with every row visible unless you pass `hiddenRows`.

## Hyperlinks

Attach links through a sheet's `hyperlinks` array. The cell keeps the text from `data`; the link sets where a click goes. The `hyperlink` builders cover the three kinds.

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

A link is also a plain object: `{ range: 'A2', url: 'https://…' }` or `{ range: 'C2', location: "'Q1 Sales'!A1" }`, with optional `tooltip` and `display`. A leading `#` in `location` is dropped.

- **Look:** the first cell of each link gets Excel's hyperlink colour (`#0563C1`) and an underline. A `color` or `underline` in that cell's own style wins, each on its own.
- **Allowed URLs:** `http:`, `https:` and `mailto:`, with spaces and quotes percent-encoded. Anything else throws, because exports often carry user data and `file:` links or custom protocol handlers are not safe to click.
- **Limits:** one link per range, 65,530 links per sheet and 2,079 characters per address.
