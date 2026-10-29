import { strToU8 } from 'fflate';

export const declareUncompressedSize = (zip: Uint8Array, entryName: string, size: number): Uint8Array => {
  const copy = new Uint8Array(zip);
  const view = new DataView(copy.buffer);
  const wanted = strToU8(entryName);
  let patched = 0;
  for (let offset = 0; offset < copy.length - 46; offset++) {
    if (view.getUint32(offset, true) !== 0x02014b50) continue;
    const nameLength = view.getUint16(offset + 28, true);
    const name = copy.subarray(offset + 46, offset + 46 + nameLength);
    if (name.length === wanted.length && name.every((byte, i) => byte === wanted[i])) {
      view.setUint32(offset + 24, size, true);
      patched++;
    }
  }
  if (patched === 0) throw new Error(`central directory entry ${entryName} not found`);
  return copy;
};
