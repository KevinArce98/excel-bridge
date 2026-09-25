---
name: excel-bridge
description: "The .xlsx toolkit that ships at 12.0 KB, shown as a shipping desk where every library carries its real weight on its label."
colors:
  dock-0: "#0b1120"
  dock: "#0f172a"
  dock-2: "#1e293b"
  plate: "#1b2336"
  plate-2: "#212c44"
  rule: "#334155"
  rule-soft: "#24304a"
  paper: "#f8fafc"
  paper-2: "#eef2f7"
  paper-3: "#e3e9f0"
  paper-rule: "#cbd5e1"
  ink: "#0f172a"
  ink-2: "#334155"
  ink-3: "#475569"
  green: "#22c55e"
  green-hi: "#34d399"
  green-deep: "#15803d"
  green-ink: "#08260f"
  green-wash: "#ecfdf5"
  amber: "#fbbf24"
  amber-deep: "#b45309"
  text-hi: "#f8fafc"
  text: "#cbd5e1"
  text-mid: "#a3b0c2"
  text-dim: "#8d9bb0"
  sheet-bg: "#ffffff"
  sheet-alt: "#f8fafc"
  sheet-line: "#e2e8f0"
  sheet-head: "#f1f5f9"
  sheet-head-tx: "#475569"
  sheet-tx: "#0f172a"
  sheet-tx-dim: "#64748b"
typography:
  weight:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "clamp(4.4rem, 9.2vw, 8.6rem)"
    fontWeight: 800
    lineHeight: 0.86
    letterSpacing: "-0.055em"
    fontFeature: "tnum"
  display:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "clamp(2.5rem, 4.5vw, 4.1rem)"
    fontWeight: 800
    lineHeight: 1.02
    letterSpacing: "-0.05em"
  display-close:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "clamp(2.6rem, 5.2vw, 4.4rem)"
    fontWeight: 800
    lineHeight: 1.08
    letterSpacing: "-0.045em"
  headline:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "clamp(1.85rem, 3.1vw, 2.7rem)"
    fontWeight: 800
    lineHeight: 1.08
    letterSpacing: "-0.045em"
  headline-doc:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "clamp(1.7rem, 2.6vw, 2.35rem)"
    fontWeight: 800
    lineHeight: 1.08
    letterSpacing: "-0.045em"
  title:
    fontFamily: "Inter, Segoe UI, system-ui, -apple-system, sans-serif"
    fontSize: "1.4rem"
    fontWeight: 700
    lineHeight: 1.25
    letterSpacing: "-0.012em"
  lede:
    fontFamily: "Inter, Segoe UI, system-ui, -apple-system, sans-serif"
    fontSize: "clamp(1.06rem, 1.45vw, 1.2rem)"
    fontWeight: 400
    lineHeight: 1.6
  body:
    fontFamily: "Inter, Segoe UI, system-ui, -apple-system, sans-serif"
    fontSize: "16px"
    fontWeight: 400
    lineHeight: 1.6
  button:
    fontFamily: "Inter, Segoe UI, system-ui, -apple-system, sans-serif"
    fontSize: "15px"
    fontWeight: 600
    lineHeight: 1
  field-caption:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "11px"
    fontWeight: 700
    lineHeight: 1.3
    letterSpacing: "0.16em"
  field-value:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "13.5px"
    fontWeight: 600
    lineHeight: 1.35
  method:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "12.5px"
    fontWeight: 500
    lineHeight: 1.6
  code:
    fontFamily: "JetBrains Mono, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace"
    fontSize: "13.5px"
    fontWeight: 400
    lineHeight: 1.75
rounded:
  zone: "3px"
  label: "6px"
  doc: "8px"
  control: "10px"
  printer: "12px"
spacing:
  gutter: "24px"
  gutter-mobile: "18px"
  field-x: "16px"
  field-y: "10px"
  hero-top: "clamp(44px, 6.5vw, 92px)"
  band-top: "clamp(72px, 9vw, 128px)"
  band-bottom: "clamp(64px, 8vw, 110px)"
components:
  button-primary:
    backgroundColor: "{colors.green}"
    textColor: "{colors.green-ink}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "0 20px"
    height: "48px"
  button-primary-hover:
    backgroundColor: "#2fd26b"
    textColor: "{colors.green-ink}"
  button-ghost:
    backgroundColor: "transparent"
    textColor: "{colors.text-hi}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "0 20px"
    height: "48px"
  button-ghost-hover:
    backgroundColor: "{colors.plate}"
    textColor: "{colors.text-hi}"
  button-quiet:
    backgroundColor: "transparent"
    textColor: "{colors.text}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "0 20px"
    height: "48px"
  tracking-field:
    backgroundColor: "{colors.dock-0}"
    textColor: "{colors.text-hi}"
    rounded: "{rounded.control}"
    padding: "14px"
  carrier-tab:
    backgroundColor: "{colors.dock}"
    textColor: "{colors.text-dim}"
    padding: "9px 15px 8px"
  carrier-tab-active:
    backgroundColor: "{colors.dock-0}"
    textColor: "{colors.text-hi}"
  shipping-label:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    typography: "{typography.field-value}"
    rounded: "{rounded.label}"
  label-routing-bar:
    backgroundColor: "{colors.green}"
    textColor: "{colors.green-ink}"
    padding: "8px 16px"
  label-zone:
    backgroundColor: "{colors.ink}"
    textColor: "{colors.paper}"
    rounded: "{rounded.zone}"
    padding: "5px 9px 4px"
  label-stub:
    backgroundColor: "{colors.paper-2}"
    textColor: "{colors.ink}"
    padding: "9px 16px 10px"
  paper-doc:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    rounded: "{rounded.doc}"
  doc-foot:
    backgroundColor: "{colors.paper-2}"
    textColor: "{colors.ink-3}"
    padding: "12px 18px 14px"
  printer-head:
    backgroundColor: "{colors.dock-2}"
    textColor: "{colors.text-mid}"
    rounded: "{rounded.printer}"
    padding: "16px 18px 30px"
  drop-dock:
    backgroundColor: "{colors.dock-2}"
    textColor: "{colors.text-hi}"
    rounded: "{rounded.control}"
    padding: "36px 44px 30px"
    height: "236px"
  drop-dock-over:
    backgroundColor: "{colors.plate-2}"
  code-window:
    backgroundColor: "{colors.dock-0}"
    textColor: "{colors.text}"
    typography: "{typography.code}"
    rounded: "{rounded.control}"
    padding: "22px 24px 26px"
  code-tab-active:
    backgroundColor: "{colors.dock-0}"
    textColor: "{colors.text-hi}"
    padding: "15px 18px 14px"
  toast:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    rounded: "{rounded.control}"
    padding: "12px 16px"
---

# Design System: excel-bridge

## Overview

**Creative North Star: "The Shipping Desk"**

excel-bridge is freight you can weigh. The page is a carrier's desk on a navy loading dock: labels, waybills, manifests and handling stamps, where every library carries its real weight printed on its label. The dock is the ground and the equipment (printer, receiving dock, code window, tracking field); anything that has been printed (labels, manifests, rate sheets, the demo spreadsheet) is light thermal paper in navy ink. Each object belongs to one of those two stocks.

The page is dense like paperwork: mono caps field captions over values, hairline rules between fields, heavy 3px ink rules under headers, dashed perforations wherever a part could tear off. It persuades with measurement, never with atmosphere. The largest type on the page is the weight on the excel-bridge label, not the headline. Motion is limited to one physical act, the label printer feeding a printout.

The world rejects the dark-SaaS hero: headline plus gradient plus glow, glass panels, floating feature-card grids, and stats as chips. It also rejects eyebrows above headings; captions live inside fields, beside values, the way a label prints them.

**Key Characteristics:**
- Two stocks: navy dock (ground and equipment) and thermal paper (anything printed).
- JetBrains Mono for display, fields, captions, numbers and code; Inter for prose, buttons and nav.
- Green is cleared-stamp and routing ink; amber is handling-and-fault ink.
- Real artifacts: Code 39 barcodes that encode real strings, labels that print the real downloaded file.
- Every number carries its method beside it.
- Flat paper with one elevation per object; no gradients, glass or glow.

## Colors

A navy dock under thermal paper, printed in navy ink, with green reserved for clearance and routing and amber reserved for handling and faults.

### Primary
- **Cleared Green** (`green`): the carrier's cleared-stamp and routing ink on the dock. The label's routing bar, the primary button, the active carrier-tab and code-tab underline, the status LED when the library is ready, the dashed lane lines of the receiving dock, the "12.0 KB." in the headline, excel-bridge's own row in the transit table, the strip-head trailing value, focus rings and text selection.
- **Deep Green** (`green-deep`): the same ink on paper, where plain green fails AA. The CLEARED stamp, excel-bridge's rate bars and figures on labels and manifests, the rate-sheet column head, check marks, the active sheet-tab underline, toast success icon.
- **Routing Highlight** (`green-hi`): links and the package name in the tracking field on the dock; keywords in code.
- **Green Ink** (`green-ink`): text set on solid green (routing bar, primary button, sheet header row, selection).
- **Cleared Wash** (`green-wash`): the excel-bridge column in the rate sheet and the total row in the demo sheet.

### Secondary
- **Handling Amber** (`amber-deep`): the HEAVY handling stamp on the ExcelJS and SheetJS back labels; also the error-toast icon and boolean cell-type ink in the read-back view.
- **Fault Amber** (`amber`): the status LED when the live library is offline; a translucent tint of it (with literal `#fde68a` / `#fef3c7` text) is the read-error panel.

### Neutral
- **Dock Floor** (`dock-0`): the deepest ground: receiving-dock band, footer, tracking field, code window body, active tab fill.
- **Loading Dock** (`dock`): page ground and sticky waybill header.
- **Dock Steel** (`dock-2`): equipment bodies: label printer head, receiving-dock drop zone.
- **Plate** (`plate`, `plate-2`): hover and drag-over fills on dock equipment and nav links.
- **Dock Rule** (`rule`, `rule-soft`): hairlines, dashed perforations and equipment borders on the dock; `rule-soft` for quieter separators (status strip, band edges, tab dividers).
- **Thermal Stock** (`paper`, `paper-2`, `paper-3`): label and document stock; `paper-2` for stubs, doc-feet, sheet tab rails, the blank label and the nearer back label; `paper-3` for the farthest back label.
- **Paper Rule** (`paper-rule`): 1px field dividers on paper.
- **Navy Ink** (`ink`, `ink-2`, `ink-3`): printed values, secondary printed text, and field captions on paper; `ink` also draws the 3px header rules, barcodes and the XLSX zone block.
- **Dock Text** (`text-hi`, `text`, `text-mid`, `text-dim`): headings, body, secondary, and captions/method lines on the dock.
- **Spreadsheet Scope** (`sheet-*`): the editable demo sheet and read-back grid are authentic light `.xlsx` documents; their white cells, gridlines, headers and text use this scoped set. Data-type inks inside the sheet are literal values in the build, not tokens: formula and number teal `#0f766e`, date violet `#6d28d9`, Excel link blue `#0563c1`.

### Named Rules
**The Two Stocks Rule.** Every surface is either dock (navy, equipment, flat with hairlines) or paper (light, printed, shadowed). Nothing sits in between; no translucent or glass layers.

**The Cleared-Ink Rule.** Green means cleared, routed, live or "this is excel-bridge". On paper it is always `green-deep`; bright `green` never carries text on paper.

**The Handling-Stamp Rule.** Amber marks handling and faults only: the HEAVY stamp, the offline LED, errors. It is never emphasis, decoration or a second brand color.

## Typography

**Display Font:** JetBrains Mono (with ui-monospace, SFMono-Regular, Menlo, Consolas, monospace)
**Body Font:** Inter (with Segoe UI, system-ui, -apple-system, sans-serif)
**Label/Mono Font:** JetBrains Mono

**Character:** The mono face is the printer: headlines, field captions, values, weights, barcodes' human-readable lines and code are all stamped in it at heavy weights (700–800) with tight negative tracking for display and wide positive tracking for caps. Inter is the clerk's handwriting: prose, ledes, buttons, nav and the few Inter h3s.

### Hierarchy
- **Weight** (800, `clamp(4.4rem, 9.2vw, 8.6rem)`, 0.86, -0.055em, tabular): the WEIGHT figure on the hero label, about 132px at 1440 and 70px at 390. The largest type on the page. The printed label's weight is `clamp(2.6rem, 4.4vw, 3.3rem)`.
- **Display** (800, `clamp(2.5rem, 4.5vw, 4.1rem)`, 1.02, -0.05em, balanced): the hero headline "Ships at 12.0 KB.", about 64px at 1440. The footer "Ship it." uses the larger close variant `clamp(2.6rem, 5.2vw, 4.4rem)`.
- **Headline** (800, `clamp(1.85rem, 3.1vw, 2.7rem)`, 1.08, -0.045em, balanced): section h2s on the dock. On paper docs, `clamp(1.7rem, 2.6vw, 2.35rem)` in navy ink.
- **Title** (Inter 700, 1.4rem, 1.25, -0.012em): h3 sub-heads ("Transit times"); 16px inside the drop zone.
- **Body** (Inter 400, 16px, 1.6): prose. Lede `clamp(1.06rem, 1.45vw, 1.2rem)` capped at 44ch; section subtitles `clamp(1rem, 1.35vw, 1.12rem)` capped at 60ch; doc-title prose 14.5px at 52ch; notes 13–13.5px.
- **Field caption** (Mono 700, 11px, 0.16em, uppercase): every field label on labels, docs, tables, strip-heads, the tracking field, status strip and footer column heads. Table header captions share it.
- **Field value** (Mono 600, 13.5px, 1.35): printed values under captions; table cells 12.5–13.5px with 700 and tabular numerals for figures.
- **Method** (Mono 500, 12.5px, 1.6): the method line under the lede, meta lines, footer bottom.
- **Code** (Mono 400, 13.5px, 1.75): the code window.

### Named Rules
**The Heaviest Number Rule.** The excel-bridge weight figure is the largest type on the page at every width. No headline or stat may outsize it.

**The Printed Field Rule.** A caption lives inside its field, directly above or beside its value, in 11px mono caps. Captions never float above a section heading as an eyebrow.

**The Tabular Figures Rule.** Every weight, time, size and count that sits in a column uses `font-variant-numeric: tabular-nums` and right alignment.

## Layout

Content width is 1200px with a 24px gutter (18px at 560 and below). The waybill header is sticky; sections are full-bleed bands with `band-top` / `band-bottom` padding, separated on the dock by a dashed perforation inside the content width; the receiving dock band and the footer drop to `dock-0` with `rule-soft` edges.

Grids, in page order:
- **Hero:** 42fr / 58fr, gap `clamp(32px, 4.5vw, 64px)`, centered. The label stack is right-aligned, `min(100%, 600px)` wide, with 80px top room for the back labels.
- **Strip-head:** 0.95fr / 1.05fr / 150px (heading, subtitle, trailing field), bottom-aligned, 3px `text-hi` rule above and 1px `rule` below.
- **Split-head:** 1fr / 1fr, bottom-aligned.
- **Pack:** 1.6fr sheet / 1fr printer.
- **Read:** 380px dock / 1fr read-out.
- **Manifest:** 1fr itemized / 1.1fr declaration.
- **Code:** 0.78fr sticky side (top 132px) / 2fr window.
- **Close:** 1.4fr / 1fr, bottom-aligned; footer link columns 1fr / 1fr.

Breakpoints (as built: 1020 / 900 / 640 / 560):
- **≤1020px:** pack and manifest grids stack; printer caps at 520px.
- **≤900px:** hero, read, close, code and split-head stack; the label stack centers; the strip-head stacks and its trailing field turns into a row under a dashed top rule; the code side un-sticks.
- **≤640px:** the rate sheet drops its header and becomes a 3-column grid per capability, the capability name spanning the row and each cell prefixed by its column name as a mono caps caption from `data-label`. The declaration of contents does the same on a 30px / 1fr / 1fr grid, the item number spanning two rows, Qty and Net weight as captioned cells.
- **≤560px:** gutter 18px; the manifest number, status strip, "Try it"/"Code" nav links and "Via" service line hide; back labels hide and the CLEARED stamp shrinks to 84px; carrier tabs become four equal-width tabs; CTAs go full width; the itemized manifest drops its min-width and tightens to 10px/12px cells; the transit table hides its bar column.

Wide tables otherwise scroll inside their own focusable region and never widen the page.

## Elevation & Depth

Paper is flat stock lifted off the dock by one shadow; dock equipment is flat and carries only hairline borders. Each object gets one elevation: a border or a shadow, never both. There are no gradients, blurs, glass or colored glows anywhere.

### Shadow Vocabulary
- **Paper** (`box-shadow: 0 2px 3px -1px rgba(2, 6, 23, 0.35), 0 22px 44px -22px rgba(2, 6, 23, 0.85)`): documents and back labels lying on the desk.
- **Lift** (`box-shadow: 0 3px 5px -2px rgba(2, 6, 23, 0.4), 0 34px 60px -28px rgba(2, 6, 23, 0.9)`): the front hero label and the demo spreadsheet, the objects in hand.
- **Float** (`box-shadow: 0 8px 14px -6px rgba(2, 6, 23, 0.55), 0 26px 48px -20px rgba(2, 6, 23, 0.7)`): the toast only.
- **Printout** (`filter: drop-shadow(0 18px 18px rgba(2, 6, 23, 0.55))`): the printed label, so the shadow follows its torn edge.
- **Primary button** (`box-shadow: 0 10px 22px -12px rgba(2, 6, 23, 0.9), inset 0 1px 0 rgba(255, 255, 255, 0.35)`): a neutral navy drop and a 1px top highlight; never a green glow.

### Named Rules
**The One Elevation Rule.** An object is bordered or shadowed, not both. Paper and the toast take shadows; dock equipment takes a hairline.

**The No-Glow Rule.** Shadows are navy (`rgba(2, 6, 23, …)`). No colored shadow, blur backdrop or gradient fill exists in the system.

## Shapes

Corners are small and practical: 3px on the XLSX zone block and LEDs (2px), 6px on labels, 8px on paper docs and the demo sheet, 10px on controls (buttons, tracking field, drop zone, code window, toast), 12px on the printer head's top corners. The printed label has no radius; its bottom edge is a torn sawtooth (12×8px triangles) made with a CSS mask.

Lines carry the form language. A heavy 3px ink rule closes every label carrier strip, weight block, doc-head and rate-sheet header (2px on the printed label). Fields are divided by 1px `paper-rule` hairlines. Dashed 1px lines are perforations: the band separators, the label stub, doc-feet, the tracking field's caption divider and the strip-head trailing field. Stamps are rotated a few degrees (CLEARED -14deg, HEAVY -5deg) and roughened by the shared SVG `#ink` filter (turbulence speckle plus slight displacement), so they read as inked, not rendered. Back labels are rotated ±1deg.

### Named Rules
**The Perforation Rule.** A dashed line means "this part tears off" or "this is a separate slip". Solid hairlines divide fields; dashes separate parts.

## Components

### Buttons
Tactile desk controls in Inter 600 15px, 48px tall, 10px radius, 20px side padding, 18px leading icon.
- **Primary:** solid Cleared Green with Green Ink text and the primary-button shadow; hover `#2fd26b`; active nudges 1px down. Used once per region ("Pack a workbook", "Print label & download .xlsx").
- **Ghost:** transparent with a `rule` hairline and `text-hi`; hover fills `plate` and lightens the border.
- **Quiet:** transparent, `rule-soft` hairline, `text` color, for secondary resets.
- **Receive:** full-width `plate` with a `rule` border; hover turns the border green. Disabled until a file has been packed; disabled buttons drop to 45% opacity.

### Waybill header and status strip
The sticky navy bar (62px) carries the mono wordmark with green "bridge", a "MANIFEST v1.4.0" field divided by a hairline, and Inter nav links (Try it, Code, Docs, npm, GitHub) with `plate` hover. Below a dashed perforation, the status strip reads STATUS with a square LED: `rule` while loading, green when the live package loads, amber when offline, plus the current call in mono 12px.

### Hero shipping label (signature)
A 4×6 thermal label on paper stock with the Lift shadow, set in mono throughout:
- **Carrier strip:** logo mark, "excel-bridge", "VIA <pm>" service line, and the inverted XLSX zone block, closed by a 3px ink rule.
- **Routing bar:** solid green band, green-ink caps: "Only what you import / tree-shaken".
- **FROM / TO** and **CONTENTS / PIECES** rows: two-column fields split by hairlines.
- **WEIGHT block:** caption "WEIGHT · MIN+GZIP", the 12.0 KB figure at the Weight role, and the circular CLEARED stamp (SVG, Deep Green, "PROVENANCE SIGNED · NPM" on its arc, `#ink` filter, -14deg), closed by a 3px rule.
- **Rates:** "Same job, other carriers" table: excel-bridge, SheetJS, ExcelJS with right-aligned weights and proportional rate bars (outlined navy for others, solid Deep Green for excel-bridge).
- **Barcode:** a real Code 39 SVG of the current install command (`NPM I EXCEL-BRIDGE` by default) with its human-readable line.
- **Stub:** dashed tear line over `paper-2` with VERSION, LICENSE and HEAVY DEPS fields.

### Back labels
Two partial labels behind the hero label (`paper-3` farther, `paper-2` nearer, Paper shadow, ±1deg) showing library name, version, weight and an amber HEAVY stamp: a double-ruled (2px border plus 1px outline) Handling Amber box, 11px mono caps at 0.2em, -5deg, inked. Decorative to assistive tech; hidden at 560 and below.

### Install tracking field with carrier tabs
A package-manager selector above a tracking field, used in the hero and the footer.
- **Carrier tabs:** a native radio group styled as mono 12.5px tabs (npm, yarn, pnpm, bun) joined to the top of the field; active tab fills `dock-0` with a 2px green underline; focus draws a green inset outline.
- **Tracking field:** `dock-0` field with a 10px radius (flat top-left under the tabs), an INSTALL caption behind a dashed divider, the command in mono 15px with the package name in `green-hi`, and a 48px copy button behind a hairline.
- **Behavior:** choosing a carrier updates both install blocks, re-draws the hero barcode to encode the new command, rewrites the label's service line to "Via <pm>", updates the barcode's accessible name, and persists the choice in `localStorage` (`excel-bridge:pm`).

### Strip-head
Section header on the dock: heading, subtitle and a trailing field (e.g. OUTBOUND / step 1 of 2 in green mono 14px) behind a dashed left rule, with a 3px `text-hi` rule above and a hairline below.

### Packing station: sheet and label printer
- **Sheet:** the editable demo spreadsheet as a white `.xlsx` document (8px, Lift shadow): a tab bar with filename and a LIVE mark, column letters and row numbers on `sheet-head`, a green header row with green-ink text, a `green-wash` total row, fx-tagged formula cells, a hyperlink cell in Excel blue, and filter buttons. Focused cells draw a 2px green inset.
- **Label printer:** a `dock-2` printer head (12px top corners, hairline) with a caps name, a square LED that turns green once fed, a meta line, and an inset slot. Below it hangs the printout: a blank `paper-2` label until the first download, then a real label for the file just built (file name, print time, sheet dimensions, filter range, contents, weight on disk, and a Code 39 barcode of the file name), with a torn sawtooth bottom edge.

### Receiving dock
A `dock-2` drop zone (min 236px, 10px radius, `rule-soft` border) with two 4px dashed green lane lines 16px from each side, an outlined "INBOUND" stencil (mono 800, 30px, 0.34em tracking, 1px stroke, no fill), a box icon and an Inter title. Hover and focus lift the fill to `plate-2`; drag-over turns the border green and pulls the lanes 6px inward. Read results render as a paper meta strip (captioned fields, 3px ink rule) over paper sheet tabs and a read-only sheet.

### Paper documents
Paper stock (8px, Paper shadow, mono) built from:
- **Doc-title:** a navy-ink h2 and Inter prose above a hairline.
- **Doc-head:** document name in mono 13px 800 caps at 0.12em on the left, method or scope on the right, closed by a 3px ink rule.
- **Itemized manifest:** a table of imports with right-aligned tabular weights and Deep Green rate bars; competitors sit in a `paper-2` footer with outlined bars.
- **Declaration of contents:** a customs form with No., description (bold mono name over Inter note), Qty and Net weight columns.
- **Rate sheet:** the capability comparison with a green-wash excel-bridge column, SVG check and warning marks, and mono sizes.
- **Doc-foot:** `paper-2` slip behind a dashed perforation, carrying the measurement method.

The **transit table** is printed directly on the dock instead: mono 13.5px rows over `rule` hairlines, slate bars with excel-bridge's in green.

### Code window
`dock-0` window with a hairline and 10px radius; a `dock` tab bar of mono 13px file-name tabs (active: `dock-0` fill with a 2px green underline) and a copy button; code in mono 13.5px/1.75 with a small token palette (keywords `green-hi`, strings `#86efac`, numbers and class names `text-hi`, comments italic `text-dim`).

### Toast
A paper slip fixed bottom-center (10px radius, Float shadow, Inter 600 14px) with a Deep Green icon, or Handling Amber on errors. Rises 16px and fades in.

### Footer "Ship it."
A `dock-0` close band: "Ship it." at the close display size, the second install block, a one-line summary, two link columns under mono caps heads, and a perforation above the bottom line ("runs excel-bridge@1.4.0 live from jsDelivr").

### Motion
One authored moment: printing a label. The printout feeds from the slot in 1.15s on `cubic-bezier(0.16, 1, 0.3, 1)`, revealed top-down with `clip-path` while settling 14px. Under `prefers-reduced-motion: reduce` there is no feed: the label appears complete, smooth scrolling is off, and scroll-into-view jumps. Elsewhere, transitions are 0.18–0.35s state changes on the same ease (hover, tab color, toast, lane shift).

### Accessibility
- WCAG-AA on both stocks: green text on paper is always `green-deep`; text on green is `green-ink`.
- Visible focus everywhere: 2px green outline, 3px offset.
- Carrier tabs are native radios in a labelled radiogroup; the drop zone is a keyboard button (Enter or Space opens the file picker) with a real file input.
- Code and sheet tab lists use roving focus: only the active tab is in the tab order, and arrow keys, Home and End move between tabs.
- The printout, read-out and toast are live regions; tables carry captions or labelled scroll regions that take focus.
- Barcodes are `role="img"` with the decoded string as their name; decorative stamps, back labels, stencil and lanes are hidden from assistive tech.
- The demo degrades to static content if the CDN module fails, and says so.

## Do's and Don'ts

### Do:
- **Do** put every object on one of the two stocks: navy dock (`dock-0`/`dock`/`dock-2`) or thermal paper (`paper`) in navy ink.
- **Do** set captions as 11px JetBrains Mono caps at 0.16em inside the field they label.
- **Do** keep the excel-bridge weight figure the largest type on the page.
- **Do** print every number with its method beside it: min+gzip, esbuild 0.27.3, measured 2026-09-24, reproducible with `pnpm run size` (benchmarks with `pnpm run bench`).
- **Do** make barcodes real Code 39 of real strings (the install command, the downloaded file name), with the human-readable line underneath.
- **Do** use `green-deep` for green on paper and reserve amber for HEAVY stamps and faults.
- **Do** give each object one elevation: a hairline on the dock or a navy shadow on paper.
- **Do** show the printed label complete when reduced motion is requested.

### Don't:
- **Don't** use the dark-SaaS look: gradients, glow, glass, or floating feature-card grids.
- **Don't** put eyebrows or kickers above headings; captions belong inside fields.
- **Don't** present stats as chips or pills; weights live on labels, manifests and rate sheets.
- **Don't** invent testimonials, customer logos, download counts or any unmeasured number; live third-party numbers come from dynamic badges only.
- **Don't** encode fake data in a barcode or draw decorative barcode stripes.
- **Don't** use amber for emphasis or decoration.
- **Don't** add colored shadows, backdrop blur, or a second elevation to an already bordered object.
