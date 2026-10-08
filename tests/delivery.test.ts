import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  ExcelReader,
  ExcelWriter,
  XLSX_CONTENT_TYPE,
  createExcelWorkbookStream,
  downloadXlsx,
  toReadableStream,
  xlsxResponse,
} from '../src';
import type { CellValue } from '../src';

const sheets = [{ name: 'Data', rows: [['id'], [1], [2]] as CellValue[][] }];

const bytesOf = async (response: Response) => new Uint8Array(await response.arrayBuffer());

const chunksOf = async function* (...chunks: number[][]) {
  for (const chunk of chunks) yield new Uint8Array(chunk);
};

describe('importing the library', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.resetModules();
  });

  it.each(['document', 'Response', 'Headers', 'ReadableStream', 'Blob', 'File'])(
    'does not need %s at import time',
    async name => {
      vi.stubGlobal(name, undefined);
      vi.resetModules();
      await expect(import('../src')).resolves.toHaveProperty('xlsxResponse');
    }
  );
});

describe('downloadXlsx', () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  const stubBrowser = () => {
    const link = { href: '', download: '', style: { display: '' }, click: vi.fn(), remove: vi.fn() };
    const body = { appendChild: vi.fn() };
    vi.stubGlobal('document', { createElement: vi.fn(() => link), body });
    const createObjectURL = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:fake');
    const revokeObjectURL = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
    return { link, body, createObjectURL, revokeObjectURL };
  };

  it('throws UNSUPPORTED without a document, and importing never touches one', () => {
    expect(typeof document).toBe('undefined');
    expect(() => downloadXlsx(new Uint8Array(1))).toThrow(
      expect.objectContaining({ code: 'UNSUPPORTED' })
    );
  });

  it('clicks a hidden link to an object URL, then revokes the URL later', () => {
    vi.useFakeTimers();
    const { link, body, createObjectURL, revokeObjectURL } = stubBrowser();

    downloadXlsx(new Uint8Array([1, 2, 3]), 'Sales 2024');

    const blob = createObjectURL.mock.calls[0][0] as Blob;
    expect(blob.type).toBe(XLSX_CONTENT_TYPE);
    expect(blob.size).toBe(3);
    expect(link.href).toBe('blob:fake');
    expect(link.download).toBe('Sales 2024.xlsx');
    expect(link.style.display).toBe('none');
    expect(body.appendChild).toHaveBeenCalledWith(link);
    expect(link.click).toHaveBeenCalledTimes(1);
    expect(link.remove).toHaveBeenCalledTimes(1);

    vi.advanceTimersByTime(39_999);
    expect(revokeObjectURL).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:fake');
  });

  it('passes a Blob through untouched', () => {
    vi.useFakeTimers();
    const { createObjectURL } = stubBrowser();
    const blob = new ExcelWriter().createWorkbook([{ data: [['a']] }]);

    downloadXlsx(blob);

    expect(createObjectURL.mock.calls[0][0]).toBe(blob);
  });

  it.each([
    [undefined, 'workbook.xlsx'],
    ['', 'workbook.xlsx'],
    ['   ', 'workbook.xlsx'],
    ['.xlsx', 'workbook.xlsx'],
    ['report', 'report.xlsx'],
    ['report.xlsx', 'report.xlsx'],
    ['REPORT.XLSX', 'REPORT.xlsx'],
    ['a/b\\c:d*e?f"g<h>i|j', 'a_b_c_d_e_f_g_h_i_j.xlsx'],
    ['line\r\nbreak', 'line_break.xlsx'],
    ['../../etc/passwd', '.._.._etc_passwd.xlsx'],
    ['Año 2024', 'Año 2024.xlsx'],
  ])('names the file for %j as %j', (given, expected) => {
    vi.useFakeTimers();
    const { link } = stubBrowser();
    downloadXlsx(new Uint8Array(1), given);
    expect(link.download).toBe(expected);
  });
});

describe('toReadableStream', () => {
  it('delivers the chunks in order and closes', async () => {
    const stream = toReadableStream(chunksOf([1, 2], [3]));
    expect(await bytesOf(new Response(stream))).toEqual(new Uint8Array([1, 2, 3]));
  });

  it('pulls from the source only as fast as the reader asks', async () => {
    let produced = 0;
    async function* source() {
      for (let i = 0; i < 100; i++) {
        produced++;
        yield new Uint8Array([i]);
      }
    }
    const reader = toReadableStream(source()).getReader();
    await reader.read();
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(produced).toBeLessThanOrEqual(3);
    await reader.cancel();
  });

  it('stops the source when the reader cancels', async () => {
    let finished = false;
    async function* source() {
      try {
        for (;;) yield new Uint8Array([1]);
      } finally {
        finished = true;
      }
    }
    const reader = toReadableStream(source()).getReader();
    await reader.read();
    await reader.cancel();
    expect(finished).toBe(true);
  });

  it('errors the stream when the source throws', async () => {
    async function* source() {
      yield new Uint8Array([1]);
      throw new Error('boom');
    }
    await expect(new Response(toReadableStream(source())).arrayBuffer()).rejects.toThrow('boom');
  });
});

describe('xlsxResponse', () => {
  it('sets the content type and an attachment header', async () => {
    const response = await xlsxResponse(new Uint8Array([1]), 'Report');
    expect(response.status).toBe(200);
    expect(response.headers.get('Content-Type')).toBe(XLSX_CONTENT_TYPE);
    expect(response.headers.get('Content-Disposition')).toBe(
      `attachment; filename="Report.xlsx"; filename*=UTF-8''Report.xlsx`
    );
  });

  it.each([
    ['Año 2024 (final)', `attachment; filename="A_o 2024 (final).xlsx"; filename*=UTF-8''A%C3%B1o%202024%20%28final%29.xlsx`],
    ["it's", `attachment; filename="it's.xlsx"; filename*=UTF-8''it%27s.xlsx`],
    ['a\r\nSet-Cookie: x=1', `attachment; filename="a_Set-Cookie_ x=1.xlsx"; filename*=UTF-8''a_Set-Cookie_%20x%3D1.xlsx`],
    ['日本語', `attachment; filename="___.xlsx"; filename*=UTF-8''%E6%97%A5%E6%9C%AC%E8%AA%9E.xlsx`],
    [undefined, `attachment; filename="workbook.xlsx"; filename*=UTF-8''workbook.xlsx`],
  ])('encodes the filename %j', async (filename, header) => {
    expect((await xlsxResponse(new Uint8Array(1), filename)).headers.get('Content-Disposition')).toBe(header);
  });

  it('keeps init and its other headers, and replaces the two it owns', async () => {
    const response = await xlsxResponse(new Uint8Array([1]), 'a', {
      status: 201,
      headers: { 'Cache-Control': 'no-store', 'Content-Type': 'text/plain' },
    });
    expect(response.status).toBe(201);
    expect(response.headers.get('Cache-Control')).toBe('no-store');
    expect(response.headers.get('Content-Type')).toBe(XLSX_CONTENT_TYPE);
  });

  it.each<[string, () => Parameters<typeof xlsxResponse>[0]]>([
    ['a Uint8Array', () => new Uint8Array([1, 2, 3])],
    ['a Blob', () => new Blob([new Uint8Array([1, 2, 3])])],
    ['a ReadableStream', () => toReadableStream(chunksOf([1, 2], [3]))],
    ['an async generator', () => chunksOf([1], [2, 3])],
  ])('sends %s as the body', async (_, make) => {
    expect(await bytesOf(await xlsxResponse(make()))).toEqual(new Uint8Array([1, 2, 3]));
  });

  it('streams a workbook that the reader can open', async () => {
    const response = await xlsxResponse(createExcelWorkbookStream(sheets), 'data');
    const parsed = new ExcelReader().parseFromBuffer(await bytesOf(response));
    expect(parsed.sheets[0].data.map(row => row[0].value)).toEqual(['id', 1, 2]);
  });

  it('rejects before a response exists when the workbook is invalid up front', async () => {
    const stream = createExcelWorkbookStream([{ name: 'a/b', rows: [['x']] }]);
    await expect(xlsxResponse(stream)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('closes the rows iterable of a streamed sheet when the client goes away', async () => {
    let started = false;
    let closed = false;
    async function* rows(): AsyncGenerator<CellValue[]> {
      started = true;
      try {
        for (let i = 0; ; i++) yield [i];
      } finally {
        closed = true;
      }
    }
    const response = await xlsxResponse(createExcelWorkbookStream([{ rows: rows() }]));
    const reader = response.body!.getReader();
    while (!started) await reader.read();
    await reader.cancel();
    expect(closed).toBe(true);
  });

  it('cannot undo the response when a later row is invalid: the body errors instead', async () => {
    const stream = createExcelWorkbookStream([{ rows: [['ok'], [NaN]] }]);
    const response = await xlsxResponse(stream);
    expect(response.status).toBe(200);
    await expect(response.arrayBuffer()).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
});
