import { readdirSync, readFileSync, writeFileSync, mkdirSync, existsSync, statSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Marked } from 'marked';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const sourceDir = join(root, 'docs-src');
const outputDir = join(root, 'docs');
const SITE = 'https://kevinarce98.github.io/excel-bridge';
const REPO = 'https://github.com/KevinArce98/excel-bridge';
const LANGUAGES = ['en', 'es'];

const COPY = {
  en: {
    skip: 'Skip to content',
    docs: 'Docs',
    site: 'Site',
    pager: 'Previous and next',
    tryIt: 'Try it',
    contents: 'Contents',
    onThisPage: 'On this page',
    previous: 'Previous',
    next: 'Next',
    copy: 'Copy',
    copied: 'Copied',
    editPage: 'Edit this page',
    language: 'Language',
    langName: 'English',
    notFound: 'Page not found',
    notFoundBody: 'That page does not exist. Start from the contents.',
    callouts: { NOTE: 'Note', TIP: 'Tip', LIMIT: 'Limit', BREAKING: 'Breaking change', WARNING: 'Warning' },
    footerLicense: 'MIT licence · Microsoft Excel is a trademark of Microsoft; this project is not affiliated with Microsoft.',
  },
  es: {
    skip: 'Saltar al contenido',
    docs: 'Docs',
    site: 'Sitio',
    pager: 'Anterior y siguiente',
    tryIt: 'Probar',
    contents: 'Contenido',
    onThisPage: 'En esta página',
    previous: 'Anterior',
    next: 'Siguiente',
    copy: 'Copiar',
    copied: 'Copiado',
    editPage: 'Editar esta página',
    language: 'Idioma',
    langName: 'Español',
    notFound: 'Página no encontrada',
    notFoundBody: 'Esa página no existe. Empieza por el contenido.',
    callouts: { NOTE: 'Nota', TIP: 'Consejo', LIMIT: 'Límite', BREAKING: 'Cambio incompatible', WARNING: 'Aviso' },
    footerLicense: 'Licencia MIT · Microsoft Excel es una marca de Microsoft; este proyecto no está afiliado a Microsoft.',
  },
};

const escapeHtml = text =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const slugify = text =>
  text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/<[^>]+>/g, '')
    .replace(/[`*_]/g, '')
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-');

const TS_KEYWORDS = new Set(
  'import export from as const let var function return if else for of in while await async new class extends type interface typeof instanceof null undefined true false throw try catch default yield void readonly keyof satisfies enum declare'.split(
    ' '
  )
);

const tokenize = (code, lang) => {
  const scripts = ['ts', 'typescript', 'js', 'javascript', 'mjs', 'tsx'];
  const pattern = scripts.includes(lang)
    ? /(\/\/[^\n]*|\/\*[\s\S]*?\*\/)|(`(?:\\[\s\S]|[^`\\])*`|'(?:\\.|[^'\\\n])*'|"(?:\\.|[^"\\\n])*")|(\b\d[\d_]*(?:\.\d+)?\b)|([A-Za-z_$][\w$]*)|([^\sA-Za-z0-9_$]+|\s+)/g
    : lang === 'json'
      ? /()("(?:\\.|[^"\\\n])*")(?=\s*:)|()("(?:\\.|[^"\\\n])*")|(-?\b\d+(?:\.\d+)?\b)|(true|false|null)|([^\s"\d]+|\s+)/g
      : lang === 'bash' || lang === 'sh' || lang === 'shell'
        ? /(#[^\n]*)|('(?:[^'\n])*'|"(?:\\.|[^"\\\n])*")|(\s-{1,2}[A-Za-z][\w-]*)|([^\s'"#]+|\s+)/g
        : null;
  if (!pattern) return escapeHtml(code);

  let html = '';
  if (scripts.includes(lang)) {
    let match;
    while ((match = pattern.exec(code))) {
      const [text, comment, string, number, word] = match;
      const safe = escapeHtml(text);
      if (comment) html += `<span class="tk-c">${safe}</span>`;
      else if (string) html += `<span class="tk-s">${safe}</span>`;
      else if (number) html += `<span class="tk-n">${safe}</span>`;
      else if (word && TS_KEYWORDS.has(word)) html += `<span class="tk-k">${safe}</span>`;
      else if (word && /^[A-Z]/.test(word)) html += `<span class="tk-t">${safe}</span>`;
      else html += safe;
    }
    return html;
  }
  if (lang === 'json') {
    let match;
    while ((match = pattern.exec(code))) {
      const text = match[0];
      const safe = escapeHtml(text);
      if (match[2]) html += `<span class="tk-t">${safe}</span>`;
      else if (match[4]) html += `<span class="tk-s">${safe}</span>`;
      else if (match[5]) html += `<span class="tk-n">${safe}</span>`;
      else if (match[6]) html += `<span class="tk-k">${safe}</span>`;
      else html += safe;
    }
    return html;
  }
  let match;
  let first = true;
  while ((match = pattern.exec(code))) {
    const [text, comment, string, flag] = match;
    const safe = escapeHtml(text);
    if (comment) html += `<span class="tk-c">${safe}</span>`;
    else if (string) html += `<span class="tk-s">${safe}</span>`;
    else if (flag) html += `<span class="tk-n">${safe}</span>`;
    else if (first && text.trim()) {
      html += `<span class="tk-k">${safe}</span>`;
      first = false;
    } else html += safe;
    if (text.includes('\n')) first = true;
  }
  return html;
};

const parseFrontMatter = (source, file) => {
  const match = /^---\n([\s\S]*?)\n---\n/.exec(source);
  if (!match) throw new Error(`${file}: missing front matter`);
  const meta = {};
  for (const line of match[1].split('\n')) {
    const entry = /^([A-Za-z]+):\s*(.*)$/.exec(line);
    if (!entry) throw new Error(`${file}: cannot read front matter line "${line}"`);
    meta[entry[1]] = entry[2].replace(/^"(.*)"$/, '$1');
  }
  for (const key of ['title', 'description', 'group', 'groupOrder', 'order']) {
    if (meta[key] === undefined) throw new Error(`${file}: front matter needs "${key}"`);
  }
  meta.groupOrder = Number(meta.groupOrder);
  meta.order = Number(meta.order);
  return { meta, body: source.slice(match[0].length) };
};

const listMarkdown = dir =>
  readdirSync(dir).flatMap(name => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? listMarkdown(path) : name.endsWith('.md') ? [path] : [];
  });

const loadPages = () =>
  LANGUAGES.flatMap(lang => {
    const dir = join(sourceDir, lang);
    if (!existsSync(dir)) return [];
    return listMarkdown(dir).map(file => {
      const source = readFileSync(file, 'utf8');
      const { meta, body } = parseFrontMatter(source, relative(root, file));
      const slug = relative(dir, file).split(sep).join('/').replace(/\.md$/, '');
      const included = meta.include ? readFileSync(join(root, meta.include), 'utf8') : body;
      return { lang, slug, file: relative(root, file).split(sep).join('/'), body: included, ...meta };
    });
  });

const pagePath = (lang, slug) => `${lang === 'en' ? '' : 'es/'}${slug === 'index' ? '' : `${slug}/`}`;
const pageUrl = (lang, slug) => `${SITE}/${pagePath(lang, slug)}`;
const hrefFrom = (from, to) => {
  const fromDepth = from.split('/').filter(Boolean).length;
  const up = '../'.repeat(fromDepth) || './';
  return `${up}${to}` || './';
};

const calloutKinds = ['NOTE', 'TIP', 'LIMIT', 'BREAKING', 'WARNING'];

const createRenderer = (lang, headings) => {
  const copy = COPY[lang];
  const usedIds = new Map();
  let lastHeading = copy.docs;
  const marked = new Marked({
    gfm: true,
    renderer: {
      heading({ tokens, depth }) {
        const text = this.parser.parseInline(tokens);
        const raw = tokens.map(token => token.raw ?? token.text ?? '').join('');
        const base = slugify(raw);
        const seen = usedIds.get(base) ?? 0;
        usedIds.set(base, seen + 1);
        const id = seen === 0 ? base : `${base}-${seen + 1}`;
        lastHeading = raw.replace(/[`*]/g, '');
        if (depth === 2 || depth === 3) headings.push({ depth, id, text: lastHeading });
        const anchor = depth > 1 ? `<a class="anchor" href="#${id}" aria-label="${escapeHtml(raw.replace(/[`*]/g, ''))}">#</a>` : '';
        return `<h${depth} id="${id}">${text}${anchor}</h${depth}>\n`;
      },
      code({ text, lang: info }) {
        const [language = ''] = (info ?? '').split(/\s+/);
        const titleMatch = /title="([^"]*)"/.exec((info ?? '').replace(language, ''));
        const tabMatch = /tab="([^"]*)"/.exec(info ?? '');
        const title = titleMatch ? titleMatch[1] : '';
        const label = title || (language ? language.toUpperCase() : '');
        const body = tokenize(text.replace(/\n$/, ''), language);
        const tab = tabMatch ? ` data-tab="${escapeHtml(tabMatch[1])}"` : '';
        return `<figure class="code"${tab}><figcaption><span class="code-name">${escapeHtml(label)}</span><button type="button" class="copy" data-copy="${escapeHtml(copy.copy)}" data-copied="${escapeHtml(copy.copied)}">${escapeHtml(copy.copy)}</button></figcaption><pre tabindex="0"><code>${body}</code></pre></figure>\n`;
      },
      codespan({ text }) {
        return `<code>${escapeHtml(text)}</code>`;
      },
      table(token) {
        const alignments = token.align.map(align => (align ? ` class="a-${align}"` : ''));
        const head = token.header
          .map((cell, index) => `<th scope="col"${alignments[index]}>${this.parser.parseInline(cell.tokens)}</th>`)
          .join('');
        const rows = token.rows
          .map(
            row =>
              `<tr>${row
                .map((cell, index) => `<td${alignments[index]} data-label="${escapeHtml(token.header[index].text)}">${this.parser.parseInline(cell.tokens)}</td>`)
                .join('')}</tr>`
          )
          .join('\n');
        return `<div class="table-scroll" role="region" tabindex="0" aria-label="${escapeHtml(lastHeading)}"><table><thead><tr>${head}</tr></thead><tbody>\n${rows}\n</tbody></table></div>\n`;
      },
      blockquote({ tokens }) {
        const first = tokens[0];
        const marker = first?.type === 'paragraph' ? /^\[!([A-Z]+)\]\s*/.exec(first.text) : null;
        if (marker && calloutKinds.includes(marker[1])) {
          const inline = first.tokens[0];
          if (inline && typeof inline.text === 'string') inline.text = inline.text.replace(/^\[!([A-Z]+)\]\s*/, '');
          const kind = marker[1].toLowerCase();
          return `<aside class="callout callout-${kind}"><p class="callout-name">${escapeHtml(copy.callouts[marker[1]])}</p>${this.parser.parse(tokens)}</aside>\n`;
        }
        return `<blockquote>${this.parser.parse(tokens)}</blockquote>\n`;
      },
      link({ href, title, tokens }) {
        const text = this.parser.parseInline(tokens);
        const external = /^https?:\/\//.test(href);
        const rel = external ? ' rel="noopener"' : '';
        const titleAttribute = title ? ` title="${escapeHtml(title)}"` : '';
        return `<a href="${escapeHtml(href)}"${titleAttribute}${rel}>${text}</a>`;
      },
    },
  });
  return marked;
};

const logoMark = `<svg class="mark" viewBox="0 0 32 32" aria-hidden="true" focusable="false"><rect width="32" height="32" rx="7" fill="#1B2336"/><rect x="5" y="6" width="22" height="6" rx="2" fill="#22C55E"/><rect x="5" y="15" width="9" height="4" rx="1.5" fill="#334155"/><rect x="18" y="15" width="9" height="4" rx="1.5" fill="#334155"/><rect x="5" y="22" width="9" height="4" rx="1.5" fill="#334155"/><rect x="18" y="22" width="9" height="4" rx="1.5" fill="#22C55E" fill-opacity="0.5"/></svg>`;

const renderNav = (pages, current, lang) => {
  const groups = [...new Map(pages.map(page => [page.groupOrder, page.group])).entries()].sort((a, b) => a[0] - b[0]);
  return groups
    .map(([groupOrder, group]) => {
      const items = pages
        .filter(page => page.groupOrder === groupOrder)
        .sort((a, b) => a.order - b.order)
        .map(page => {
          const active = page.slug === current.slug;
          const href = hrefFrom(pagePath(current.lang, current.slug), pagePath(lang, page.slug));
          return `<li><a href="${href}"${active ? ' aria-current="page"' : ''}>${escapeHtml(page.navTitle ?? page.title)}</a></li>`;
        })
        .join('');
      return `<section><p class="rail-group">${escapeHtml(group)}</p><ul>${items}</ul></section>`;
    })
    .join('\n');
};

const renderPage = (page, pages, allPages) => {
  const copy = COPY[page.lang];
  const headings = [];
  const marked = createRenderer(page.lang, headings);
  const bodyHtml = marked.parse(page.body);
  const here = pagePath(page.lang, page.slug);
  const toRoot = hrefFrom(here, '');
  const ordered = [...pages].sort((a, b) => a.groupOrder - b.groupOrder || a.order - b.order);
  const index = ordered.findIndex(entry => entry.slug === page.slug);
  const previous = ordered[index - 1];
  const next = ordered[index + 1];
  const alternates = LANGUAGES.map(lang => ({ lang, page: allPages.find(entry => entry.lang === lang && entry.slug === page.slug) })).filter(entry => entry.page);
  const homeHref = toRoot;
  const tocItems = headings
    .filter(heading => heading.depth === 2)
    .map(heading => `<li><a href="#${heading.id}">${escapeHtml(heading.text)}</a></li>`)
    .join('');
  const stub = (entry, caption, direction) =>
    entry
      ? `<a class="stub stub-${direction}" href="${hrefFrom(here, pagePath(page.lang, entry.slug))}"><span class="field-caption">${escapeHtml(caption)}</span><span class="stub-title">${escapeHtml(entry.title)}</span></a>`
      : '<span></span>';
  const languageLinks = alternates
    .map(({ lang, page: alternate }) => {
      const href = hrefFrom(here, pagePath(lang, alternate.slug));
      return lang === page.lang
        ? `<span class="lang-current" aria-current="true" lang="${lang}">${lang.toUpperCase()}</span>`
        : `<a href="${href}" hreflang="${lang}" lang="${lang}" aria-label="${escapeHtml(COPY[lang].langName)}">${lang.toUpperCase()}</a>`;
    })
    .join('');
  const canonical = pageUrl(page.lang, page.slug);
  const hreflangs = alternates
    .map(({ lang, page: alternate }) => `<link rel="alternate" hreflang="${lang}" href="${pageUrl(lang, alternate.slug)}" />`)
    .join('\n');
  const title = page.slug === 'index' ? `${page.title} | excel-bridge` : `${page.title} | excel-bridge docs`;

  return `<!doctype html>
<html lang="${page.lang}">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHtml(title)}</title>
<meta name="description" content="${escapeHtml(page.description)}" />
<link rel="canonical" href="${canonical}" />
${hreflangs}
<meta property="og:type" content="article" />
<meta property="og:title" content="${escapeHtml(title)}" />
<meta property="og:description" content="${escapeHtml(page.description)}" />
<meta property="og:url" content="${canonical}" />
<meta name="theme-color" content="#0f172a" />
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='7' fill='%231B2336'/%3E%3Crect x='5' y='6' width='22' height='6' rx='2' fill='%2322C55E'/%3E%3Crect x='5' y='15' width='9' height='4' rx='1.5' fill='%23334155'/%3E%3Crect x='18' y='15' width='9' height='4' rx='1.5' fill='%23334155'/%3E%3Crect x='5' y='22' width='9' height='4' rx='1.5' fill='%23334155'/%3E%3Crect x='18' y='22' width='9' height='4' rx='1.5' fill='%2322C55E' fill-opacity='0.5'/%3E%3C/svg%3E" />
<link rel="preconnect" href="https://fonts.googleapis.com" />
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
<link rel="stylesheet" href="${toRoot}assets/docs.css" />
</head>
<body>
<a class="skip" href="#content">${escapeHtml(copy.skip)}</a>
<header class="waybill">
  <div class="waybill-in">
    <a class="wordmark" href="${homeHref}" aria-label="excel-bridge">${logoMark}<span><span class="w-excel">excel</span><span class="w-dash">-</span><span class="w-bridge">bridge</span></span></a>
    <span class="field"><span class="field-caption">${escapeHtml(copy.docs)}</span><span class="field-value">2.x</span></span>
    <nav class="top-links" aria-label="${escapeHtml(copy.site)}">
      <a href="${homeHref}#pack">${escapeHtml(copy.tryIt)}</a>
      <a href="${REPO}" rel="noopener">GitHub</a>
      <a href="https://www.npmjs.com/package/excel-bridge" rel="noopener">npm</a>
    </nav>
    <div class="lang" role="group" aria-label="${escapeHtml(copy.language)}">${languageLinks}</div>
  </div>
</header>
<div class="shell">
  <nav class="rail" aria-label="${escapeHtml(copy.docs)}">
    <details class="rail-fold" open>
      <summary>${escapeHtml(copy.contents)}</summary>
      ${renderNav(pages, page, page.lang)}
    </details>
  </nav>
  <main id="content" class="sheet">
    <div class="sheet-rule" aria-hidden="true"></div>
    <article class="prose">
${bodyHtml}
    </article>
    <nav class="stubs" aria-label="${escapeHtml(copy.pager)}">${stub(previous, copy.previous, 'prev')}${stub(next, copy.next, 'next')}</nav>
    <p class="doc-foot"><a href="${REPO}/edit/main/${page.file}" rel="noopener">${escapeHtml(copy.editPage)}</a> · ${escapeHtml(copy.footerLicense)}</p>
  </main>
  <aside class="toc" aria-label="${escapeHtml(copy.onThisPage)}">
    <p class="field-caption">${escapeHtml(copy.onThisPage)}</p>
    <ul>${tocItems}</ul>
  </aside>
</div>
<script src="${toRoot}assets/docs.js" defer></script>
</body>
</html>
`;
};

const renderNotFound = lang => {
  const copy = COPY[lang];
  return `<!doctype html>
<html lang="${lang}">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHtml(copy.notFound)} | excel-bridge</title>
<meta name="robots" content="noindex" />
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@500;800&display=swap" rel="stylesheet" />
<link rel="stylesheet" href="${SITE}/assets/docs.css" />
</head>
<body class="lost">
<main class="lost-card">
  <h1>${escapeHtml(copy.notFound)}</h1>
  <p>${escapeHtml(copy.notFoundBody)}</p>
  <p><a class="button" href="${SITE}/${lang === 'en' ? '' : 'es/'}">${escapeHtml(copy.docs)}</a></p>
</main>
</body>
</html>
`;
};

const renderSitemap = pages =>
  `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:xhtml="http://www.w3.org/1999/xhtml">
<url><loc>${SITE}/</loc></url>
${pages
  .map(page => {
    const alternates = LANGUAGES.map(lang => pages.find(entry => entry.lang === lang && entry.slug === page.slug))
      .filter(Boolean)
      .map(entry => `<xhtml:link rel="alternate" hreflang="${entry.lang}" href="${pageUrl(entry.lang, entry.slug)}"/>`)
      .join('');
    return `<url><loc>${pageUrl(page.lang, page.slug)}</loc>${alternates}</url>`;
  })
  .join('\n')}
</urlset>
`;

const buildAll = () => {
  const allPages = loadPages();
  const files = new Map();
  for (const lang of LANGUAGES) {
    const pages = allPages.filter(page => page.lang === lang);
    for (const page of pages) {
      files.set(`${pagePath(lang, page.slug)}index.html`, renderPage(page, pages, allPages));
    }
    if (pages.length) files.set(lang === 'en' ? '404.html' : 'es/404.html', renderNotFound(lang));
  }
  files.set('sitemap.xml', renderSitemap(allPages));
  files.set('robots.txt', `User-agent: *\nAllow: /\nSitemap: ${SITE}/sitemap.xml\n`);
  return files;
};

const main = () => {
  const files = buildAll();
  if (process.argv.includes('--check')) {
    const stale = [...files].filter(([path, content]) => {
      const file = join(outputDir, path);
      return !existsSync(file) || readFileSync(file, 'utf8') !== content;
    });
    if (stale.length) {
      console.error(`Generated docs are out of date:\n  ${stale.map(([path]) => `docs/${path}`).join('\n  ')}\nRun: pnpm run docs:build`);
      process.exit(1);
    }
    console.log(`docs ok: ${files.size} generated files match docs-src`);
    return;
  }
  for (const [path, content] of files) {
    const file = join(outputDir, path);
    mkdirSync(dirname(file), { recursive: true });
    writeFileSync(file, content);
  }
  console.log(`docs built: ${files.size} files`);
};

main();
