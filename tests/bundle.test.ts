import { describe, it, expect } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname } from 'node:path';
import { build } from 'esbuild';

const entry = fileURLToPath(new URL('../src/index.ts', import.meta.url));

const bundledInputs = async (names: string): Promise<string[]> => {
  const result = await build({
    stdin: {
      contents: `export { ${names} } from ${JSON.stringify(entry)};`,
      loader: 'ts',
      resolveDir: dirname(entry),
    },
    bundle: true,
    minify: true,
    platform: 'browser',
    format: 'esm',
    write: false,
    metafile: true,
    logLevel: 'silent',
  });
  const [output] = Object.values(result.metafile.outputs);

  return Object.entries(output.inputs)
    .filter(([, input]) => input.bytesInOutput > 0)
    .map(([path]) => path);
};

describe('Tree-shaking', () => {
  it.each(['ExcelWriter', 'createExcelWorkbookStream', 'hyperlink'])(
    '%s leaves the reader out of the bundle',
    async name => {
      const inputs = await bundledInputs(name);

      expect(inputs.length).toBeGreaterThan(0);
      expect(inputs.filter(path => /src\/reader\/|fast-xml-parser/.test(path))).toEqual([]);
    }
  );

  it('ExcelReader leaves the writer side out of the bundle', async () => {
    const inputs = await bundledInputs('ExcelReader');

    expect(inputs.some(path => path.includes('src/reader/'))).toBe(true);
    expect(
      inputs.filter(path => /src\/(writer\/|core\/(hyperlinks|cell-ref|xml-templates))/.test(path))
    ).toEqual([]);
  });
});
