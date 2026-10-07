const VALID_SHEET_NAME = /^[^\\/?*[\]:']([^\\/?*[\]:]{0,29}[^\\/?*[\]:'])?$/;

export const validateSheetName = (name: string): void => {
  if (!VALID_SHEET_NAME.test(name)) {
    throw new Error(
      `Invalid sheet name "${name}": use 1 to 31 characters, none of \\ / ? * [ ] :, and no apostrophe at either end`
    );
  }
};

export const validateSheetNames = (names: string[]): void => {
  if (names.length === 0) {
    throw new Error('A workbook needs at least one sheet');
  }

  const seen = new Set<string>();
  for (const name of names) {
    validateSheetName(name);
    const key = name.toLowerCase();
    if (seen.has(key)) {
      throw new Error(`Sheet name "${name}" is used twice (names are not case-sensitive)`);
    }
    seen.add(key);
  }
};
