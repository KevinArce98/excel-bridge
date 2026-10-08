export interface Rng {
  next: () => number;
  int: (low: number, high: number) => number;
  pick: <T>(items: readonly T[]) => T;
  chance: (probability: number) => boolean;
}

export const makeRng = (seed: number): Rng => {
  let state = seed >>> 0;
  const next = () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  return {
    next,
    int: (low, high) => low + Math.floor(next() * (high - low + 1)),
    pick: items => items[Math.floor(next() * items.length)],
    chance: probability => next() < probability,
  };
};

const NAMES = [
  'a',
  'b',
  'c',
  'r',
  't',
  'v',
  'f',
  'is',
  'si',
  'row',
  'x:y',
  'ns0:elem',
  'dc:title',
  'x14:cfRule',
  'n-1',
  'n.2',
  'n_3',
  'é',
  '日本',
  'A1',
  'xml',
  'mc:AlternateContent',
  'sst',
  'extLst',
  'ext',
  'val',
  'id',
  'name',
  'ref',
] as const;
const ATTRIBUTES = [
  'r',
  's',
  't',
  'x',
  'id',
  'name',
  'xml:space',
  'r:id',
  'xmlns',
  'xmlns:r',
  'ref',
  'val',
  'a1',
  'é',
  'mc:Ignorable',
  'x14ac:dyDescent',
  'ht',
  'n-m',
  'n.m',
] as const;
const TEXTS = [
  '',
  'x',
  ' ',
  '  ',
  '\n',
  '\n  ',
  '\t',
  'hello world',
  'a&amp;b',
  '&lt;tag&gt;',
  '&quot;q&quot;',
  '&apos;',
  '&amp;amp;',
  '&amp;#233;',
  'a > b',
  'a ] b',
  ']]',
  ']>',
  '1e5',
  '日本語 😀',
  ' ',
  'a\r\nb',
  'a\rb',
  'a\nb',
  ' ',
  '&#65;&#x42;&#233;',
  '&#10;',
  '&#9;',
  '&#128512;',
  '&#38;',
  '&#60;x&#62;',
] as const;
const ATTRIBUTE_VALUES = [
  '',
  'v',
  ' ',
  'a b',
  '1',
  '1.5',
  '&lt;&amp;&gt;&quot;&apos;',
  "it's",
  'say "x"',
  'a\nb',
  'a\tb',
  'a\r\nb',
  'a/b',
  'a>b',
  '=',
  'a=b',
  '日本語',
  '&amp;#65;',
  '_x000D_',
  '/',
  '//',
  '&#65;',
  '&#38;amp;',
  '&#10;',
] as const;
const CDATA = [
  '',
  'x',
  '<a>',
  '&amp;',
  ']]',
  ']',
  '>',
  'a\r\nb',
  ' ',
  '\n',
  '<!-- c -->',
  '<![CDATA',
] as const;

const element = (rng: Rng, depth: number): string => {
  const name = rng.pick(NAMES);
  let open = name;
  const used = new Set<string>();
  for (let i = rng.int(0, 4); i > 0; i--) {
    const attribute = rng.pick(ATTRIBUTES);
    if (used.has(attribute)) continue;
    used.add(attribute);
    const value = rng.pick(ATTRIBUTE_VALUES);
    const quote = value.includes('"') ? "'" : rng.pick(['"', "'"] as const);
    const escaped = quote === "'" ? value.replace(/'/g, '&apos;') : value;
    open += `${rng.pick([' ', '  ', '\n', '\t', ' \r\n'] as const)}${attribute}${rng.pick(['=', ' = ', '=\n'] as const)}${quote}${escaped}${quote}`;
  }
  const close = `</${name}${rng.pick(['', '', ' ', '\n'] as const)}>`;
  if (depth > 6 || rng.chance(0.25)) {
    if (rng.chance(0.5)) return `<${open}${rng.pick(['', ' ', '\n'] as const)}/>`;
    return `<${open}>${rng.chance(0.7) ? rng.pick(TEXTS) : ''}${close}`;
  }
  let content = '';
  for (let i = rng.int(0, 6); i > 0; i--) {
    const kind = rng.next();
    if (kind < 0.45) content += element(rng, depth + 1);
    else if (kind < 0.75) content += rng.pick(TEXTS);
    else if (kind < 0.85) content += `<![CDATA[${rng.pick(CDATA)}]]>`;
    else if (kind < 0.93) content += `<!--${rng.pick(['', ' c ', ' a > b '] as const)}-->`;
    else content += rng.pick(['<?pi?>', '<?pi x="1"?>', '<?php echo 1 ?>'] as const);
  }
  return `<${open}>${content}${close}`;
};

export const randomXml = (seed: number): string => {
  const rng = makeRng(seed * 40503);
  const declaration = rng.pick([
    '',
    '',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
    "<?xml version='1.0'?>\n",
    '  \n',
    '<!-- lead -->',
  ] as const);
  return (
    declaration + element(rng, 0) + rng.pick(['', '\n', '\r\n', '  ', '<!-- trail -->'] as const)
  );
};
