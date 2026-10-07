import { it } from 'vitest';

interface Expectation {
  message: RegExp;
  error?: string;
}

export const knownDefect = (
  name: string,
  run: () => unknown,
  { message, error = 'AssertionError' }: Expectation
): void => {
  it(`known defect: ${name}`, async () => {
    const failure = await Promise.resolve()
      .then(run)
      .then(
        () => undefined,
        (reason: unknown) => reason ?? new Error('rejected without a value')
      );

    if (failure === undefined) {
      throw new Error(`"${name}" no longer fails. Replace knownDefect() with a regular test.`);
    }

    const { name: errorName, message: errorMessage } = failure as Error;
    if (errorName !== error || !message.test(errorMessage)) {
      throw new Error(`"${name}" fails for a different reason: ${errorName}: ${errorMessage}`, {
        cause: failure,
      });
    }
  });
};
