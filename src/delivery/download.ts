import { XLSX_CONTENT_TYPE } from '../core/constants';
import { ExcelBridgeError } from '../core/errors';
import { xlsxFilename } from './filename';

const REVOKE_DELAY_MS = 40_000;

export function downloadXlsx(data: Blob | Uint8Array, filename?: string): void {
  if (typeof document === 'undefined' || typeof URL.createObjectURL !== 'function') {
    throw new ExcelBridgeError(
      'UNSUPPORTED',
      'downloadXlsx needs a browser: document and URL.createObjectURL are not available here'
    );
  }

  const blob =
    data instanceof Blob ? data : new Blob([data as BlobPart], { type: XLSX_CONTENT_TYPE });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = xlsxFilename(filename);
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), REVOKE_DELAY_MS);
}
