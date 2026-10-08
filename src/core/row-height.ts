export const MAX_ROW_HEIGHT = 409.5;

export const isRowHeight = (height: unknown): height is number =>
  typeof height === 'number' && height > 0 && height <= MAX_ROW_HEIGHT;
