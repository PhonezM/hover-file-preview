# Changelog

## Unreleased

- Dev dependencies refreshed by Dependabot (#1): ECharts/ZRender 6.1, uuid 14,
  esbuild 0.28, tslib 2.9.
- Renderer rebuilt with ECharts 6.1: 741,377 bytes minified and a 280,973-byte
  Chrome package. The demo deck's chart is unchanged apart from antialiasing;
  unit, size-budget and 6/6 end-to-end checks pass.
- `build.sh` pins esbuild 0.28.2 so packaging matches the lockfile.

## 1.0.1 - 2026-11-03

Performance, safety limits, localisation and an end-to-end test. Previews are
unchanged in every kind; the renderer is 42.8% smaller, only loaded when a deck
is actually previewed, and hostile Office files can no longer hang the page.

**Renderer**

- echarts is imported modularly (bar, line, pie, SVG); the minified renderer is
  700,976 bytes instead of 1,203,512, and the bundle no longer contains any
  `new Function`.
- PowerPoint uses the library's list mode: slides are stacked in place instead
  of being cloned one by one, so the 120 ms per-chart wait and the whole clone
  loop are gone.
- A deck stops at 100 parsed and rendered slides; the header says
  "first 100 of N slides" instead of silently cutting it.
- The 685 KB renderer is no longer a static content script. A background script
  injects it into the tab's isolated world only when a .pptx link is hovered, so
  other pages run the ~100 KB of readers and the card instead.

**Limits**

- The zip reader enforces 64 MiB per entry and 256 MiB per package for the
  actual inflated bytes, validates the archive offsets, and refuses unknown
  compression methods.
- A merged cell whose range is far outside the displayed sheet is clamped
  before iterating (A1:XFD1048576 used to ask for 17.2 billion cells).
- A row past the sheet's row limit is reported as truncated instead of being
  drawn at the wrong index.

**Safety**

- Word output validates colours, fonts, style values, hyperlink schemes and
  list markers, so a crafted .docx can no longer inject HTML through a style
  attribute or a `javascript:` link.
- Markdown links only allow http(s), mailto and fragments; markdown images
  become clickable links instead of automatically requesting third-party URLs.

**Card**

- A replaced or closed card aborts its in-flight download and every superseded
  parse/render is discarded; the preview cache keeps a 64 MiB budget, not only
  30 entries.
- Two redundant copies of the deck buffer per render are gone.
- The interface ships in English and Traditional Chinese (`_locales`), chosen by
  the browser locale with English as the fallback.
- Accessibility: the card is a labelled dialog, status and metadata are live
  regions, sheet tabs are a tablist/tabpanel, Tabbing onto a previewable link
  opens the card and moves focus into it, and `Esc` closes it.

**Delivery**

- CI on Node 22 runs the unit tests, the build, `web-ext lint`, a bundle-size
  budget and the Chromium end-to-end suite, then uploads both store packages.
- A `v*` tag publishes the release; addons.mozilla.org signing runs when
  `WEB_EXT_API_KEY`/`WEB_EXT_API_SECRET` are configured, and the unsigned build
  otherwise says so.
- Added `SECURITY.md`, `CONTRIBUTING.md`, `PRIVACY.md`, `RELEASING.md`, issue and
  pull-request templates, and weekly Dependabot updates.

**Tests**

- 24 unit tests, including truncated-XML termination and the injection cases.
- 6 Chromium end-to-end cases (`npm run test:e2e`) against the packaged
  extension, including keyboard open/close; before/after screenshots differ only
  in the build stamp.

## 1.0.0 - 2026-10-30

First public release. Hover an image, PDF, text/markdown, spreadsheet, Word or
PowerPoint link and the file is previewed in a card without being downloaded.

**Card**

- One card path: placed once from the room and width it will have, never moving
  afterwards, and kept inside the window for every kind of preview.
- Built in a closed shadow root; closes 200 ms after the pointer leaves it and
  never while the pointer is on it. `Esc` closes it, scrolling the page closes
  it, scrolling the deck inside it does not.
- An LRU of 30 entries keeps the fetched bytes; a rendered document keeps its
  pictures alive for as long as its HTML stays cached, so a second hover shows
  the pictures instead of blank boxes.
- A link on another origin whose file sends no CORS header shows no card at all
  rather than a card that cannot be filled.

**Word (.docx)**

- Our own reader: paragraphs and runs (bold, italic, underline, strike, size,
  colour, highlight, font), heading styles, numbered and bulleted lists, tables
  (grid widths, horizontal and vertical merges, shading, borders from the table
  style), pictures, page breaks, headers and footers, section columns and
  footnotes.
- Numbering follows the file's `w:lvlText`, so a heading under "2." reads "2.1"
  and under that "3.1.1"; the label sits at the level's own indent, which for a
  numbered heading is the page margin.
- Paragraph spacing starts from `w:docDefaults` and is overridden by the style
  chain and then by the paragraph itself, so a document that relies on the
  defaults keeps Word's line pitch.
- The page is the document's own sheet: width and margins come from
  `w:pgSz`/`w:pgMar` (A4 = 794 px wide), so tables and centred pictures sit
  where Word puts them.
- Pictures stay inline (`w:drawing` is inline in Word): two in one paragraph
  share a line and a centred paragraph centres them.
- Table borders are drawn in the collapsing model: the frame sits on the outer
  cells and `w:tcBorders` still wins per side.

**Excel (.xlsx)**

- Our own reader: every sheet, values, booleans, gaps and merged ranges.
- Every cell is shown through its own number format: `General`, decimals,
  thousands separators, percentages, scientific/engineering notation, fractions,
  currency and colour sections (`[$$-409]`, `[Red]`), the positive/negative/zero
  sections, and date/time codes including the 1900 leap-year quirk.
- Column widths (`<cols>`) and row heights (`<row ht>`) are kept; a cell wraps
  instead of running off the card.
- A sheet longer than 2000 rows says so in the card header ("first 2000 rows")
  instead of being cut silently.

**PowerPoint (.pptx)**

- Slides render as real slides with the theme colours, tables using the file's
  own columns, hyperlinks, callout geometry and the font's line pitch.
- Charts are drawn without the entry animation, with a palette only when the
  file names colours, data labels when asked for, and a legend only when the
  file has one.
- The deck is built out of sight and stacked in the card, capped at 100 slides.

**Packaging**

- The zip layer is ours (`DecompressionStream`); there is no third-party zip
  library in the package.
- One build produces both store packages: `hover-file-preview-firefox.zip`
  (addons.mozilla.org) and `hover-file-preview-chrome.zip` (Chrome Web Store),
  each a plain zip with `manifest.json` at the root and PNG icons at 16/32/48/128.
- `npm test` covers link detection, markdown, the reader helpers, number formats
  and the docx table/marker rules (21 `node:test` cases).
