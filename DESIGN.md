# Design

<!-- impeccable:design-schema 1 -->

Visual system for the excel-bridge demo/landing page (`docs/index.html`). Recorded from the built page. The organizing idea: **the page is built on a spreadsheet** — real cells, column letters, row numbers, and A1 coordinates are the layout and wayfinding, and the library's two directions (write / read) are shown as authentic light `.xlsx` documents living inside a dark workspace.

## Mode

Persuade, with an embedded Operate playground. A first-time visitor edits a live sheet, downloads a real `.xlsx`, and reads one back — proof over claims.

## Color

Committed strategy: a warm-navy ground owns the surface; green is the "active / live" signal (selected cells, primary actions, section addresses, rules), carrying ~30%. The interactive demo artifacts break to authentic **light** spreadsheet documents so styling and conditional formatting read truthfully.

Ground & surface (dark workspace)
- `--navy-0 #0B1120` deepest ground · `--navy-1 #0F172A` base · `--navy-2 #1E293B`
- `--surface #1b2336` (dark cells) · `--surface-2 #212c44` · `--surface-3 #273049`
- `--line #334155` (hairline) · `--line-soft #24304a`

Accent (green — the live signal)
- `--green #22c55e` (primary) · `--green-hi #34d399` (links/highlight) · `--green-deep #15803d` (green text on light surfaces, AA)
- `--green-dim rgba(34,197,94,.13)` fills · `--green-glow rgba(34,197,94,.28)` selection only

Text (on dark) — all ≥ WCAG AA
- `--text-hi #f8fafc` · `--text #cbd5e1` · `--text-mid #a3b0c2` · `--text-dim #8d9bb0`
- `--amber #fbbf24` (limited/warning marks)

Light `.xlsx` document scope (the two demo sheets)
- `--sheet-bg #ffffff` · `--sheet-alt #f8fafc` · `--sheet-line #e2e8f0` · `--sheet-head #f1f5f9` (col/row chrome)
- `--sheet-tx #0f172a` · `--sheet-tx-dim #64748b`
- Header row: green `#22c55e` fill with dark `#08260F` bold text (legible, on-brand — not white-on-green)
- Conditional-format color scale (matches the file written): `#f8696b → #ffeb84 → #63be7b`, always with dark cell text

## Type

- Data face — **JetBrains Mono** (`--mono`): cell values, A1 coordinates, code, chips, nav, install command, formula bar. Monospace is used for genuine tabular data / code / measurement, never as decoration.
- Prose face — **Inter** (`--sans`): headings and body. (Inherited brand face from `assets/banner.svg`; deliberate for a developer audience.)
- Display: H1 `clamp(2.5rem, 5.4vw, 4.15rem)`, tracking `-0.032em`, weight 800, `text-wrap: balance`. H2 `clamp(1.8rem, 3.4vw, 2.7rem)`, `-0.026em`. Body 1.6 line-height; lede/prose measures kept ≤ ~60ch.
- Numerals use `font-variant-numeric: tabular-nums` wherever figures align (sheet cells, perf, metadata).

## Space & shape

- Radii: `--r-sm 8` · `--r-md 12` · `--r-lg 16` · `--r-xl 22`. Card radii stay 12–16; pills for small controls only.
- Section rhythm: `clamp(58px, 8vw, 104px)` vertical band padding; more space above a heading than below.
- Content width `--maxw 1200px`, gutter 24px (18px ≤560).
- Elevation is declared **once** per element — border **or** shadow, never a 1px border under a wide soft shadow. Floating things (toast) use shadow only; contained things (install, code window, comparison) use a hairline border only. The sheet uses an outline + a real offset/blur drop shadow (`--shadow-sheet`).

## Components

- **Formula-bar nav** — sticky top bar (wordmark, version pill, Docs/npm/GitHub) over a thin monospace `fx` strip that echoes the current live call (`new ExcelWriter().createWorkbook([sheet]) → Blob`, and updates on download/read). Hidden ≤560.
- **Sheet** (`.sheet-shell` + `table.xl`) — light `.xlsx` document with a titlebar (traffic-light dots, filename, static LIVE mark), A/B/C column headers and 1..n row numbers, green header row, editable number/text cells, a live color-scale conditional-format column, and formula cells tagged `fx`. Scrolls horizontally inside its own container on narrow screens; never widens the page.
- **Cell region** (`.feat-region`/`.feat`) — capabilities laid out as a contiguous sheet region of coordinate-addressed cells (varied column spans), 1px gridline gaps, line-icon + heading + terse copy. Not free-floating cards.
- **Comparison sheet** (`table.cmp`) — the README comparison as a styled table; the excel-bridge column tinted with the green live signal; authored SVG check / warn marks (never emoji), `Pro` / `CJS-first` qualifiers, `None`/`Several`.
- **Buttons** — primary = solid green with dark text + inset highlight and a neutral offset shadow (no neon glow); ghost = surface + hairline; mini/outline variants for the light sheet.
- **Code window** — tabbed (`write/read/edit/stream`), traffic-light chrome, hand-rolled tokenizer (keywords, strings, comments, numbers, class names), copy button.
- **Chips / badges** — mono hairline chips; live shields.io badges (npm version, downloads) plus a text link to Bundlephobia (no fragile size badge on the hero).
- **Browser surfaces themed**: green text selection, green caret in editable cells, slate custom scrollbars, green `:focus-visible` rings, tabular numerals.

## Motion

One authored moment: **the download build**. On "Download .xlsx" the visible grid flashes green column-by-column and a confirmation chip (filename · size · dims) rises and settles (exponential ease-out). All entrance motion is gated behind `prefers-reduced-motion`. No scattered per-section reveals; content is visible by default.

## Accessibility

WCAG-AA contrast throughout (verified by the Impeccable detector). Full keyboard operation: editable cells are real inputs; the dropzone is a keyboard-activatable control with a standard file picker fallback; visible focus rings; `aria-live` on demo results/status; honored reduced-motion.

## Runtime

Static, no build step. The library is loaded as browser ESM from jsDelivr (`excel-bridge@1.3.0/+esm`), so the demo runs the real published package — generate a styled workbook and download it, or read an uploaded/just-generated file and render its parsed cells, types, formulas and styles (lossless round-trip). Fonts from Google Fonts with system fallbacks. Degrades to static content if the CDN module fails to load.
