import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { ExcelBridge } from '../../src';
import type { ParsedSheet, ParsedWorkbook } from '../../src/reader';

export const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/${name}`, import.meta.url)));

export const cellAt = (sheet: ParsedSheet, coordinate: string) =>
  sheet.data.flat().find(cell => cell?.coordinate === coordinate);

export const calendarDay = (date: Date) => [
  date.getFullYear(),
  date.getMonth() + 1,
  date.getDate(),
];

export const part = (bytes: Uint8Array, path: string): string => strFromU8(unzipSync(bytes)[path]);

export const readFirstSheet = (bytes: Uint8Array): ParsedSheet => ExcelBridge.read(bytes).sheets[0];

export const normalise = (workbook: ParsedWorkbook) => ({
  metadata: {
    creator: workbook.metadata.creator,
    title: workbook.metadata.title,
    subject: workbook.metadata.subject,
  },
  sheets: workbook.sheets.map(sheet => ({
    ...sheet,
    data: sheet.data.map(row =>
      row.map(cell => ({
        ...cell,
        value: cell.value instanceof Date ? calendarDay(cell.value) : cell.value,
      }))
    ),
  })),
});
