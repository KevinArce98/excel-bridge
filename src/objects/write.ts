import type { ExcelData, SheetOptions } from '../writer';
import type { StreamingSheetInput } from '../writer/stream';
import { invalidInput } from '../core/errors';
import type { CellStyle, CellValue } from '../core/types';

export interface WriteColumn<Row extends object = Record<string, CellValue>> {
  key: Extract<keyof Row, string>;
  header?: string;
  width?: number;
  style?: CellStyle;
  numberFormat?: string;
}

export type StreamingColumn<Row extends object = Record<string, CellValue>> = Omit<
  WriteColumn<Row>,
  'style' | 'numberFormat'
>;

export interface ObjectsToSheetOptions extends SheetOptions {
  headerStyle?: CellStyle;
}

export interface ObjectsToStreamingSheetOptions extends Omit<
  StreamingSheetInput,
  'rows' | 'columnWidths'
> {
  headerStyle?: CellStyle;
}

interface ColumnShape {
  key: string;
  header?: string;
  width?: number;
  style?: CellStyle;
  numberFormat?: string;
}

const bodyStyle = ({ style, numberFormat }: ColumnShape): CellStyle | undefined =>
  numberFormat === undefined ? style : { ...style, numberFormat };

const widthsOf = (columns: ColumnShape[]): number[] | undefined => {
  if (columns.every(column => column.width === undefined)) return undefined;
  const widths: number[] = [];
  columns.forEach(({ width }, index) => {
    if (width !== undefined) widths[index] = width;
  });
  return widths;
};

const headerOf = (columns: ColumnShape[]): CellValue[] =>
  columns.map(column => column.header ?? column.key);

const valueAt = (row: object, key: string): CellValue =>
  key in Object.prototype && !Object.prototype.hasOwnProperty.call(row, key)
    ? undefined
    : ((row as Record<string, unknown>)[key] as CellValue);

const cellsOf = <Row extends object>(row: Row, columns: WriteColumn<Row>[]): CellValue[] =>
  columns.map(column => valueAt(row, column.key));

const headerStyles = (
  columns: ColumnShape[],
  headerStyle?: CellStyle
): Record<string, CellStyle> => {
  const styles: Record<string, CellStyle> = {};
  if (headerStyle) columns.forEach((_, col) => (styles[`0-${col}`] = headerStyle));
  return styles;
};

export function objectsToSheet<Row extends object = Record<string, CellValue>>(
  rows: Iterable<Row>,
  columns: WriteColumn<Row>[],
  { headerStyle, ...options }: ObjectsToSheetOptions = {}
): ExcelData {
  const data: CellValue[][] = [headerOf(columns)];
  for (const row of rows) data.push(cellsOf(row, columns));

  const styles = headerStyles(columns, headerStyle);
  columns.forEach((column, col) => {
    const style = bodyStyle(column);
    if (style) for (let row = 1; row < data.length; row++) styles[`${row}-${col}`] = style;
  });

  const columnWidths = widthsOf(columns);
  return {
    data,
    ...(Object.keys(styles).length > 0 ? { styles } : {}),
    options: { ...options, ...(columnWidths ? { columnWidths } : {}) },
  };
}

export function objectsToStreamingSheet<Row extends object = Record<string, CellValue>>(
  rows: Iterable<Row> | AsyncIterable<Row>,
  columns: StreamingColumn<Row>[],
  { headerStyle, ...options }: ObjectsToStreamingSheetOptions = {}
): StreamingSheetInput {
  const styled = columns.find(column => bodyStyle(column as ColumnShape));
  if (styled) {
    throw invalidInput(`Streaming column "${styled.key}" cannot have a style`);
  }

  async function* cells(): AsyncGenerator<CellValue[]> {
    yield headerOf(columns);
    for await (const row of rows) yield cellsOf(row, columns);
  }

  const styles = { ...headerStyles(columns, headerStyle), ...options.styles };
  const columnWidths = widthsOf(columns);
  return {
    ...options,
    rows: { [Symbol.asyncIterator]: cells },
    styles,
    ...(columnWidths ? { columnWidths } : {}),
  };
}
