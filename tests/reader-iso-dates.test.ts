import { describe, expect, it } from 'vitest';
import { parseIsoDate } from '../src/core/date-utils';
import { calendarDay, cellAt, readFirstSheet } from './helpers/read';
import { buildXlsx } from './helpers/xlsx';

const cellWith = (value: string) =>
  cellAt(
    readFirstSheet(
      buildXlsx(`<sheetData><row r="1"><c r="A1" t="d"><v>${value}</v></c></row></sheetData>`)
    ),
    'A1'
  );

const wallClock = (date: Date) => [
  ...calendarDay(date),
  date.getHours(),
  date.getMinutes(),
  date.getSeconds(),
  date.getMilliseconds(),
];

describe('ISO 8601 date cells (t="d")', () => {
  it.each([
    ['2024-02-29', [2024, 2, 29, 0, 0, 0, 0]],
    ['2024-02-29T00:00:00.000Z', [2024, 2, 29, 0, 0, 0, 0]],
    ['2024-02-29T13:45:30', [2024, 2, 29, 13, 45, 30, 0]],
    ['2024-02-29T13:45', [2024, 2, 29, 13, 45, 0, 0]],
    ['2024-02-29T13:45:30.5', [2024, 2, 29, 13, 45, 30, 500]],
    ['2024-02-29T13:45:30.123456', [2024, 2, 29, 13, 45, 30, 123]],
    ['2024-02-29T23:59:59Z', [2024, 2, 29, 23, 59, 59, 0]],
    ['0099-01-01', [99, 1, 1, 0, 0, 0, 0]],
  ])(
    'reads %j as the wall-clock time %j, whatever the time zone of the machine',
    (text, expected) => {
      const cell = cellWith(text);
      expect(cell?.type).toBe('date');
      expect(wallClock(cell?.value as Date)).toEqual(expected);
    }
  );

  it.each([
    '',
    'not a date',
    '2024-13-01',
    '2024-02-30',
    '2023-02-29',
    '2024-02-29T24:00:00',
    '2024-02-29T12:60:00',
    '2024-02-29T12:00:60',
    '29/02/2024',
    '45000',
    '2024-2-9',
    '2024-02-29 13:45:30',
    '2024-02-29T13:45:30+05:30',
    '2024-02-29T13:45:30-0800',
    '2024-02-29T13:45:30+00:00',
    '  2024-02-29  ',
    '2024-02-29Z',
    '2024-02-29t13:45:30',
    '2024-02-29T13:45:30z',
  ])('keeps %j as text when it is not a valid ISO date', text => {
    const cell = cellWith(text);
    if (text === '') {
      expect(cell?.type).toBe('empty');
    } else {
      expect(cell).toMatchObject({ type: 'string', value: text });
    }
  });

  it('keeps the formula of a t="d" cell', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" t="d"><f>TODAY()</f><v>2024-02-29</v></c></row></sheetData>'
    );
    expect(cellAt(readFirstSheet(bytes), 'A1')).toMatchObject({
      type: 'date',
      formula: 'TODAY()',
    });
  });

  it('returns undefined for a value that is not an ISO date', () => {
    expect(parseIsoDate('tomorrow')).toBeUndefined();
    expect(parseIsoDate('2024-02-29')).toBeInstanceOf(Date);
  });
});
