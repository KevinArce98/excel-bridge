import { ExcelBridgeError } from '../core/errors';

export type XmlTree = Record<string, any>;

type XmlValue = string | XmlTree;

interface OpenElement {
  name: string;
  attributes: Record<string, string> | null;
  children: XmlTree | null;
  text: string;
}

const MAX_DEPTH = 101;

const NAMED_REFERENCES: Record<string, string> = {
  lt: '<',
  gt: '>',
  amp: '&',
  quot: '"',
  apos: "'",
};

const REFERENCE = /&(?:#(\d+)|#x([\da-fA-F]+)|(lt|gt|amp|quot|apos));/g;

const isXmlChar = (code: number): boolean =>
  code === 9 ||
  code === 10 ||
  code === 13 ||
  (code >= 32 && code <= 0xd7ff) ||
  (code >= 0xe000 && code <= 0x10ffff && code !== 0xfffe && code !== 0xffff);

const resolveReference = (
  reference: string,
  decimal?: string,
  hex?: string,
  named?: string
): string => {
  if (named) return NAMED_REFERENCES[named];
  const code = decimal ? parseInt(decimal, 10) : parseInt(hex as string, 16);
  return isXmlChar(code) ? String.fromCodePoint(code) : reference;
};

const normalizeLineEnds = (text: string): string =>
  text.indexOf('\r') < 0 ? text : text.replace(/\r\n?/g, '\n');

const decodeText = (text: string): string => {
  const normalized = normalizeLineEnds(text);
  return normalized.indexOf('&') < 0 ? normalized : normalized.replace(REFERENCE, resolveReference);
};

const isNameStart = (code: number): boolean =>
  (code >= 97 && code <= 122) ||
  (code >= 65 && code <= 90) ||
  code === 95 ||
  code === 58 ||
  code > 127;

const isNameChar = (code: number): boolean =>
  isNameStart(code) || (code >= 48 && code <= 57) || code === 45 || code === 46;

const finish = ({ attributes, children, text }: OpenElement): XmlValue => {
  if (!children) {
    if (!attributes) return text;
    if (text) attributes['#text'] = text;
    return attributes;
  }
  if (attributes) Object.assign(children, attributes);
  if (text) children['#text'] = text;
  return children;
};

export const parseXml = (xml: string): XmlTree => {
  const length = xml.length;
  const stack: OpenElement[] = [{ name: '', attributes: null, children: null, text: '' }];
  let current = stack[0];
  let prefix = '';
  let position = xml.charCodeAt(0) === 0xfeff ? 1 : 0;

  const fail = (message: string): never => {
    throw new ExcelBridgeError('INVALID_FILE', `Malformed XML: ${message} at position ${position}`);
  };

  const scanName = (from: number): number => {
    let end = from;
    while (isNameChar(xml.charCodeAt(end))) end++;
    return end;
  };

  const skipSpace = (from: number): number => {
    let end = from;
    while (xml.charCodeAt(end) <= 32) end++;
    return end;
  };

  const attach = (parent: OpenElement, name: string, value: XmlValue): void => {
    const key = prefix && name.startsWith(prefix) ? name.slice(prefix.length) : name;
    const children = (parent.children ??= {});
    if (!Object.hasOwn(children, key)) {
      children[key] = value;
    } else if (Array.isArray(children[key])) {
      children[key].push(value);
    } else {
      children[key] = [children[key], value];
    }
  };

  while (position < length) {
    const lessThan = xml.indexOf('<', position);
    const textEnd = lessThan < 0 ? length : lessThan;

    if (textEnd > position) {
      const text = xml.slice(position, textEnd);
      if (stack.length > 1) current.text += decodeText(text);
      else if (/\S/.test(text)) fail('text outside the root element');
    }
    if (lessThan < 0) break;
    position = lessThan;

    const marker = xml.charCodeAt(position + 1);

    if (marker === 47) {
      const end = xml.indexOf('>', position + 2);
      if (end < 0) fail('closing tag is not terminated');
      const nameEnd = position + 2 + current.name.length;
      if (
        stack.length === 1 ||
        !xml.startsWith(current.name, position + 2) ||
        skipSpace(nameEnd) !== end
      ) {
        fail(`unexpected closing tag, expected </${current.name}>`);
      }
      const closed = stack.pop() as OpenElement;
      current = stack[stack.length - 1];
      attach(current, closed.name, finish(closed));
      position = end + 1;
    } else if (marker === 63) {
      const end = xml.indexOf('?>', position + 2);
      if (end < 0) fail('processing instruction is not terminated');
      position = end + 2;
    } else if (marker === 33) {
      if (xml.startsWith('<!--', position)) {
        const end = xml.indexOf('-->', position + 4);
        if (end < 0) fail('comment is not terminated');
        position = end + 3;
      } else if (xml.startsWith('<![CDATA[', position)) {
        const end = xml.indexOf(']]>', position + 9);
        if (end < 0) fail('CDATA section is not terminated');
        if (stack.length === 1) fail('CDATA outside the root element');
        current.text += normalizeLineEnds(xml.slice(position + 9, end));
        position = end + 3;
      } else {
        fail('DOCTYPE and other declarations are not supported');
      }
    } else {
      if (!isNameStart(marker)) fail('invalid tag name');
      let cursor = scanName(position + 1);
      const name = xml.slice(position + 1, cursor);
      if (name === '__proto__') fail('reserved name');
      let attributes: Record<string, string> | null = null;
      let selfClosing = false;

      for (;;) {
        const afterValue = cursor;
        cursor = skipSpace(cursor);
        position = cursor;
        const code = xml.charCodeAt(cursor);
        if (code === 62) {
          cursor++;
          break;
        }
        if (code === 47) {
          if (xml.charCodeAt(cursor + 1) !== 62) fail('stray "/" in tag');
          selfClosing = true;
          cursor += 2;
          break;
        }
        if (!isNameStart(code) || (attributes && cursor === afterValue)) {
          fail('invalid attribute name');
        }
        const nameEnd = scanName(cursor);
        const attributeName = xml.slice(cursor, nameEnd);
        if (attributeName === '__proto__') fail('reserved name');
        cursor = skipSpace(nameEnd);
        if (xml.charCodeAt(cursor) !== 61) fail(`attribute ${attributeName} has no value`);
        cursor = skipSpace(cursor + 1);
        const quote = xml.charCodeAt(cursor);
        if (quote !== 34 && quote !== 39) fail(`attribute ${attributeName} is not quoted`);
        const valueEnd = xml.indexOf(quote === 34 ? '"' : "'", cursor + 1);
        if (valueEnd < 0) fail(`attribute ${attributeName} is not terminated`);
        (attributes ??= {})[attributeName] = decodeText(xml.slice(cursor + 1, valueEnd));
        cursor = valueEnd + 1;
      }

      if (stack.length === 1) {
        const colon = name.indexOf(':');
        prefix = colon > 0 ? name.slice(0, colon + 1) : '';
      }
      if (selfClosing) {
        attach(current, name, attributes ?? '');
      } else {
        if (stack.length > MAX_DEPTH) fail('elements are nested too deeply');
        current = { name, attributes, children: null, text: '' };
        stack.push(current);
      }
      position = cursor;
    }
  }

  if (stack.length > 1) fail(`<${current.name}> is not closed`);
  return current.children ?? {};
};
