import { XLSX_CONTENT_TYPE } from '../core/constants';
import { attachmentHeader, xlsxFilename } from './filename';

const readableFrom = (
  iterator: AsyncIterator<Uint8Array>,
  first?: IteratorResult<Uint8Array>
): ReadableStream<Uint8Array> => {
  let pending = first;
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      const result = pending ?? (await iterator.next());
      pending = undefined;
      if (result.done) controller.close();
      else controller.enqueue(result.value);
    },
    async cancel(reason) {
      await iterator.return?.(reason);
    },
  });
};

export const toReadableStream = (chunks: AsyncIterable<Uint8Array>): ReadableStream<Uint8Array> =>
  readableFrom(chunks[Symbol.asyncIterator]());

export async function xlsxResponse(
  body: Blob | Uint8Array | ReadableStream<Uint8Array> | AsyncIterable<Uint8Array>,
  filename?: string,
  init: ResponseInit = {}
): Promise<Response> {
  const headers = new Headers(init.headers);
  headers.set('Content-Type', XLSX_CONTENT_TYPE);
  headers.set('Content-Disposition', attachmentHeader(xlsxFilename(filename)));

  let payload: BodyInit;
  let iterator: AsyncIterator<Uint8Array> | undefined;
  if (body instanceof Blob || ArrayBuffer.isView(body)) {
    payload = body as BodyInit;
  } else if (Symbol.asyncIterator in body && !(body instanceof ReadableStream)) {
    iterator = body[Symbol.asyncIterator]();
    payload = readableFrom(iterator, await iterator.next());
  } else {
    payload = body as BodyInit;
  }
  try {
    return new Response(payload, { ...init, headers });
  } catch (error) {
    await iterator?.return?.().catch(() => undefined);
    throw error;
  }
}
