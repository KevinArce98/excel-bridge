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
  if (body instanceof Blob || body instanceof Uint8Array || body instanceof ReadableStream) {
    payload = body as BodyInit;
  } else {
    const iterator = body[Symbol.asyncIterator]();
    payload = readableFrom(iterator, await iterator.next());
  }
  return new Response(payload, { ...init, headers });
}
