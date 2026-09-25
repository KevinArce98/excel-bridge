import { validateRowIndex, validateColIndex } from './date-utils';

export interface CellCoord {
  row: number;
  col: number;
}

export interface CellRange {
  start: CellCoord;
  end: CellCoord;
}

const CELL_REF = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i;

export const indexToColumnLetter = (index: number): string => {
  let letter = '';
  let num = index + 1;

  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }

  return letter;
};

export const columnLetterToIndex = (letters: string): number => {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
};

export const coordinateToIndex = (coordinate: string): CellCoord => {
  const match = coordinate.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid coordinate format: ${coordinate}`);
  }

  return { row: parseInt(match[2]) - 1, col: columnLetterToIndex(match[1]) };
};

export const indexToCoordinate = (row: number, col: number): string =>
  `${indexToColumnLetter(col)}${row + 1}`;

const parseCellRef = (ref: string, range: string): CellCoord => {
  const match = CELL_REF.exec(ref);
  if (!match) {
    throw new Error(`Invalid range format: ${range}`);
  }

  const coord = { row: Number(match[2]) - 1, col: columnLetterToIndex(match[1].toUpperCase()) };
  validateRowIndex(coord.row);
  validateColIndex(coord.col);
  return coord;
};

export const parseRange = (range: string): CellRange => {
  const refs = String(range).split(':');
  if (refs.length > 2) {
    throw new Error(`Invalid range format: ${range}`);
  }

  const first = parseCellRef(refs[0], range);
  const second = refs.length === 2 ? parseCellRef(refs[1], range) : first;

  return {
    start: { row: Math.min(first.row, second.row), col: Math.min(first.col, second.col) },
    end: { row: Math.max(first.row, second.row), col: Math.max(first.col, second.col) },
  };
};

export const formatRange = ({ start, end }: CellRange, absolute = false): string => {
  const cell = ({ row, col }: CellCoord) =>
    absolute ? `$${indexToColumnLetter(col)}$${row + 1}` : `${indexToColumnLetter(col)}${row + 1}`;

  return start.row === end.row && start.col === end.col
    ? cell(start)
    : `${cell(start)}:${cell(end)}`;
};

export const quoteSheetName = (name: string): string => `'${name.replace(/'/g, "''")}'`;
