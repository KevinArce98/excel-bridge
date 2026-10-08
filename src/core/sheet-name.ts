import { invalidInput } from './errors';
const VALID_SHEET_NAME =
  /^[^\\/?*[\]:\x00-\x1F']([^\\/?*[\]:\x00-\x1F]{0,29}[^\\/?*[\]:\x00-\x1F'])?$/;

export const validateSheetName = (name: string): void => {
  if (!VALID_SHEET_NAME.test(name)) {
    throw invalidInput(
      `Invalid sheet name "${name}": use 1 to 31 characters, none of \\ / ? * [ ] : or control characters, and no apostrophe at either end`
    );
  }
};

export const validateSheetNames = (names: string[]): void => {
  if (names.length === 0) {
    throw invalidInput('A workbook needs at least one sheet');
  }

  const seen = new Set<string>();
  for (const name of names) {
    validateSheetName(name);
    const key = name.toLowerCase();
    if (seen.has(key)) {
      throw invalidInput(`Sheet name "${name}" is used twice (names are not case-sensitive)`);
    }
    seen.add(key);
  }
};
