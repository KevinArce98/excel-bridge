import { expect, it } from 'vitest';

export const knownDefect = (
  name: string,
  run: () => unknown,
  failureName = 'AssertionError'
): void => {
  it(`known defect: ${name}`, async () => {
    const failure = await Promise.resolve()
      .then(run)
      .then(
        () => undefined,
        (error: unknown) => error ?? new Error('rejected without a value')
      );

    if (failure === undefined) {
      throw new Error(`"${name}" no longer fails. Replace knownDefect() with a regular test.`);
    }
    expect((failure as Error).name).toBe(failureName);
  });
};
