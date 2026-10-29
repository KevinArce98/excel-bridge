import { isExcelErrorValue } from '../core/cells';
import { ExcelBridgeError, invalidInput } from '../core/errors';
import type { ErrorCell } from '../core/types';
import type { ParsedSheet } from '../reader';

export type ObjectCellValue = string | number | boolean | Date | ErrorCell | null;

export type ReadColumn<Row extends object> = {
  [Key in Extract<keyof Row, string>]: {
    key: Key;
    header?: string;
    column?: number;
    optional?: boolean;
    parse?: (value: ObjectCellValue, row: number, column: number) => Row[Key];
  };
}[Extract<keyof Row, string>];

export interface SheetToObjectsOptions<Row extends object> {
  headerRow?: number | null;
  columns?: ReadColumn<Row>[];
}

interface ResolvedColumn {
  key: string;
  index: number;
  parse?: (value: ObjectCellValue, row: number, column: number) => unknown;
}

type SheetRow = ParsedSheet['data'][number];

const rowNumber = (row: SheetRow): number => row.find(cell => cell)?.rowIndex ?? -1;

const headerText = (cell: SheetRow[number] | undefined): string =>
  cell && cell.type !== 'empty' ? String(cell.value).trim() : '';

const uniqueKeys = (headers: string[]): ResolvedColumn[] => {
  const taken = new Set<string>();
  const columns: ResolvedColumn[] = [];
  headers.forEach((header, index) => {
    if (!header || header === '__proto__') return;
    let key = header;
    for (let copy = 1; taken.has(key); copy++) key = `${header}_${copy}`;
    taken.add(key);
    columns.push({ key, index });
  });
  return columns;
};

const explicitColumns = <Row extends object>(
  definitions: ReadColumn<Row>[],
  headers: string[] | undefined,
  headerRow: number | null
): ResolvedColumn[] =>
  definitions.map(({ key, header = key, column, optional, parse }) => {
    const index = column ?? headers?.indexOf(header) ?? -1;
    if (index < 0 && !optional) {
      throw headerRow === null
        ? invalidInput(`Column "${key}" needs a column index because there is no header row`)
        : new ExcelBridgeError(
            'INVALID_FILE',
            `Column "${header}" not found in row ${headerRow + 1}`
          );
    }
    return { key, index, parse };
  });

export function sheetToObjects<Row extends object = Record<string, ObjectCellValue>>(
  sheet: Pick<ParsedSheet, 'data'>,
  { headerRow = 0, columns }: SheetToObjectsOptions<Row> = {}
): Row[] {
  const rows = sheet.data.filter(row => (row?.length ?? 0) > 0);
  const header = headerRow === null ? undefined : rows.find(row => rowNumber(row) === headerRow);
  const headers = header ? header.map(headerText) : undefined;
  const resolved = columns
    ? explicitColumns(columns, headers, headerRow)
    : uniqueKeys(headers ?? []);

  const objects: Row[] = [];
  for (const row of rows) {
    const rowIndex = rowNumber(row);
    if (headerRow !== null && rowIndex <= headerRow) continue;
    let filled = false;
    const object: Record<string, unknown> = {};
    for (const { key, index, parse } of resolved) {
      const cell = row[index];
      const value: ObjectCellValue =
        !cell || cell.type === 'empty'
          ? null
          : cell.type === 'error' && isExcelErrorValue(cell.value)
            ? { error: cell.value }
            : cell.value;
      if (value !== null && value !== '') filled = true;
      object[key] = parse ? parse(value, rowIndex, index) : value;
    }
    if (filled) objects.push(object as Row);
  }
  return objects;
}
