# Changelog

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
