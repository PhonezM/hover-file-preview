import { createRequire } from "node:module";
import assert from "node:assert/strict";
import test from "node:test";

const require = createRequire(import.meta.url);
const { detectKind, fileName, formatSize, decodeText } = require("../src/lib.js");

const JIRA = "https://hatool.home/jira/secure/attachment/";

test("detects Jira attachment types", () => {
  assert.equal(detectKind(JIRA + "1304210/Data+patching_WT_6002429553+%26+6009027901.pdf"), "pdf");
  assert.equal(detectKind(JIRA + "1304208/CheckTable_byEncounter_ENC20191107002164.txt"), "text");
  assert.equal(detectKind(JIRA + "1304207/INC000001767347.xlsx"), "sheet");
  assert.equal(detectKind(JIRA + "1304208/spec.docx"), "docx");
  assert.equal(detectKind(JIRA + "1304208/deck.pptx"), "pptx");
  assert.equal(detectKind(JIRA + "1304208/migration.sql"), "text");
  assert.equal(detectKind(JIRA + "1304208/legacy.doc"), null);
});

test("ignores query string and hash", () => {
  assert.equal(detectKind("https://x.test/a/b/scan.PNG?w=200#frag"), "image");
  assert.equal(detectKind("https://x.test/dl?id=1"), null);
});

test("rejects non-http schemes and extensionless paths", () => {
  assert.equal(detectKind("javascript:void(0)"), null);
  assert.equal(detectKind("data:text/plain,hi"), null);
  assert.equal(detectKind("https://x.test/dir/"), null);
});

test("resolves relative hrefs against a base", () => {
  assert.equal(detectKind("../files/report.pdf", "https://x.test/a/b/page.html"), "pdf");
  assert.equal(detectKind("/files/notes.md", "https://x.test/"), "text");
});

test("preserves '+' in the displayed filename", () => {
  assert.equal(fileName(JIRA + "1304210/Data+patching_%26_x.pdf"), "Data+patching_&_x.pdf");
});

test("formats sizes", () => {
  assert.equal(formatSize(512), "512 B");
  assert.equal(formatSize(2048), "2.0 KB");
  assert.equal(formatSize(3 * 1024 * 1024), "3.0 MB");
});

test("decodes UTF-8 and falls back to Big5", () => {
  assert.equal(decodeText(Buffer.from("hello 世界", "utf8")), "hello 世界");
  // 測試 in Big5 = B4 FA B8 D5 (invalid as UTF-8)
  assert.equal(decodeText(Buffer.from([0xb4, 0xfa, 0xb8, 0xd5])), "測試");
});


// ---------------------------------------------------------------- markdown

const { markdown } = require("../src/lib.js");

test("markdown renders headings, paragraphs and emphasis", () => {
  assert.equal(
    markdown("# Title\n\nSome **bold** and *soft* text."),
    "<h1>Title</h1>\n<p>Some <strong>bold</strong> and <em>soft</em> text.</p>",
  );
});

test("markdown escapes raw html and keeps code spans literal", () => {
  const html = markdown("<script>alert(1)</script>\n\nuse `a < b` here");
  assert.equal(html.includes("<script>"), false);
  assert.equal(html.includes("&lt;script&gt;"), true);
  assert.equal(html.includes("<code>a &lt; b</code>"), true);
});

test("markdown renders lists, fences and links", () => {
  const html = markdown("- one\n- two\n\n```js\nif (a < b) go();\n```\n\n[docs](https://example.test/x)");
  assert.equal(html.includes("<ul><li>one</li><li>two</li></ul>"), true);
  assert.equal(html.includes('<code class="lang-js">if (a &lt; b) go();</code>'), true);
  assert.equal(html.includes('href="https://example.test/x"'), true);
});

test("markdown renders pipe tables and blockquotes", () => {
  const html = markdown("| a | b |\n| --- | --- |\n| 1 | 2 |\n\n> quoted");
  assert.equal(html.includes("<th>a</th><th>b</th>"), true);
  assert.equal(html.includes("<td>1</td><td>2</td>"), true);
  assert.equal(html.includes("<blockquote><p>quoted</p></blockquote>"), true);
});

test("markdown renders task list checkboxes", () => {
  const html = markdown("- [x] done\n- [ ] todo");
  assert.equal(html.includes("\u2611 done"), true);
  assert.equal(html.includes("\u2610 todo"), true);
});

// ------------------------------------------------------------ github raw

const { rawUrl } = require("../src/lib.js");
const fs = require("node:fs");
const path = require("node:path");

test("github blob links resolve to the file on raw.githubusercontent.com", () => {
  assert.equal(
    rawUrl("https://github.com/PhonezM/hover-file-preview/blob/main/demo/demo-photo.png"),
    "https://github.com/PhonezM/hover-file-preview/raw/main/demo/demo-photo.png",
  );
  assert.equal(
    rawUrl("https://github.com/o/r/raw/v2/a/b/c.pptx"),
    "https://github.com/o/r/raw/v2/a/b/c.pptx",
  );
});

test("other links are left alone", () => {
  assert.equal(rawUrl("https://example.test/a/b.png"), "https://example.test/a/b.png");
  assert.equal(
    rawUrl("https://github.com/PhonezM/hover-file-preview/tree/main/demo"),
    "https://github.com/PhonezM/hover-file-preview/tree/main/demo",
  );
});

// ------------------------------------------------------------- own xlsx

require("../src/zip.js");
require("../src/xlsx.js");
const xlsx = globalThis.HFPxlsx;

test("xlsx: sheets, values, dates, booleans, gaps and merges", async () => {
  const file = path.join(import.meta.dirname, "fixtures", "sheet.xlsx");
  const sheets = await xlsx.read(fs.readFileSync(file));

  assert.equal(sheets.length, 2);
  assert.equal(sheets[0].name, "Q3");
  assert.equal(sheets[1].name, "Notes");

  const html = sheets[0].html;
  assert.match(html, /<td>Region<\/td>/);
  assert.match(html, /<td>1284<\/td>/);
  assert.match(html, /<td>942.5<\/td>/);
  assert.match(html, /<td>2026-01-31<\/td>/, "a date number format becomes a date");
  assert.match(html, /<td>TRUE<\/td>/, "a boolean cell reads TRUE");
  assert.match(html, /<td colspan="2" rowspan="2">merged<\/td>/, "a merged range is one cell");
  assert.match(html, /<td><\/td>/, "an empty cell in the middle keeps its column");

  assert.match(sheets[1].html, /hello &lt;world&gt; &amp; co/);
});

// Every expectation below was checked against SheetJS SSF.format (and the zoo
// workbook renders identically through both readers).
test("xlsx: a cell reads through its own number format code", () => {
  const f = xlsx._fmt;
  const cases = [
    ["General", 12345.678, "12345.678"],
    ["0.00", 12345.678, "12345.68"],
    ["#,##0.00", 1234567.891, "1,234,567.89"],
    ["0.00%", 0.256, "25.60%"],
    ["0.00E+00", 12345, "1.23E+04"],
    ["##0.0E+0", 0.000123, "123.0E-6"],
    ["00000", 42, "00042"],
    ["#.##", 0.5, ".5"],
    ['"USD "0.00', -1, "-USD 1.00"],
    ["[$€-407]#,##0.00", 1234.5, "€1,234.50"],
    ["#,##0.00;[Red](#,##0.00)", -1234.5, "(1,234.50)"],
    ['0;-0;"zero"', 0, "zero"],
    ["#,##0,", 12345678, "12,346"],
    ["@", 2.5, "2.5"],
    ["m/d/yy", 45567, "10/2/24"],
    ["d-mmm-yy", 45567, "2-Oct-24"],
    ["dd/mm/yyyy hh:mm", 45567.5, "02/10/2024 12:00"],
    ["h:mm AM/PM", 0.75, "6:00 PM"],
  ];
  for (const [code, v, want] of cases) assert.equal(f(code, v), want, code + " " + v);
  assert.equal(f("", 42), "42", "an unstyled cell reads as General");
});

test("xlsx: a cut sheet reports the truncation, and column widths are kept", async () => {
  const file = path.join(import.meta.dirname, "fixtures", "sheet.xlsx");
  const full = await xlsx.read(fs.readFileSync(file));
  assert.equal(full[0].truncated, false, "a short sheet is complete");

  const cut = await xlsx.read(fs.readFileSync(file), { maxRows: 2 });
  assert.equal(cut[0].truncated, true, "a sheet cut at maxRows says so");
  assert.match(cut[0].html, /^<table>/, "still a table");
});

// ------------------------------------------------------------- own docx

globalThis.HFPdocx = undefined;
require("../src/docx.js");
const docx = globalThis.HFPdocx;

test("docx: headings, run styles, lists, tables, merges, images, header/footer", async () => {
  const file = path.join(import.meta.dirname, "fixtures", "doc.docx");
  const html = await docx.render(fs.readFileSync(file));

  assert.match(html, /<h1[^>]*>[\s\S]*?Northwind Coffee/, "a Heading1 becomes <h1>");
  assert.match(html, /<h2[^>]*>[\s\S]*?Q3 2026 review/, "a Heading2 becomes <h2>");
  assert.match(html, /font-weight:bold/, "bold run");
  assert.match(html, /font-style:italic/, "italic run");
  assert.match(html, /text-decoration:underline/, "underline run");
  assert.match(html, /font-size:16pt/, "run size");
  assert.match(html, /color:#C00000/i, "run colour");
  assert.match(html, /class="docx-marker">•/, "a bullet list marker");
  assert.match(html, /class="docx-marker">1\./, "a numbered list marker");
  assert.match(html, /<table class="docx-table"/, "a table");
  assert.match(html, /colspan="2"/, "a horizontal merge");
  assert.match(html, /class="docx-img"/, "a picture");
  assert.match(html, /class="docx-header"/, "the header");
  assert.match(html, /class="docx-footer"/, "the footer");
  assert.match(html, /column-count:2/, "the two column section");
  // the second section is continuous, so it carries on on page 1; only the
  // explicit page break starts page 2
  const pages = html.split(/<div class="docx-page"/).slice(1);
  assert.equal(pages.length, 2, "a page break is the only thing that starts a sheet");
  assert.equal(pages[0].includes("Two column text"), true, "the continuous section continues on page 1");
  assert.equal(pages[1].includes("After the page break"), true, "the text after the break is on page 2");
  assert.equal(pages[0].includes("After the page break"), false, "and not on page 1");
  assert.match(html, /2 \/ 2/, "pages are numbered");
});

// A paragraph-level <w:tabs><w:tab/></w:tabs> is a tab *stop*, not a tab in the
// text. It used to be picked up by the run scanner, which then skipped the
// rest of the paragraph: Word's TOC entries and table labels lost their text.
test("docx: a tab stop in w:pPr keeps the paragraph text", () => {
  const ctx = { styles: {}, numbering: { nums: {}, abstract: {} }, rels: {}, media: {}, counters: {}, para: () => ({}), run: docx._props };

  const withStop = docx._para(
    '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="2473"/></w:tabs></w:pPr>' +
      "<w:r><w:t>Feature File Path</w:t></w:r><w:r><w:tab/></w:r></w:p>",
    ctx
  );
  assert.match(withStop, /Feature File Path/, "the text before the tab stop survives");

  const plainTab = docx._para('<w:p><w:r><w:tab/><w:t>after inline tab</w:t></w:r></w:p>', ctx);
  assert.match(plainTab, /after inline tab/, "a real run-level tab still renders its text");
});

// A w:drawing is inline in Word: two pictures in one paragraph share a line and
// a centred paragraph centres them. The card stylesheet makes .docx-img a
// block, so the renderer has to set the inline flow itself.
test("docx: pictures stay inline", () => {
  const ctx = {
    styles: {},
    numbering: { nums: {}, abstract: {} },
    rels: { rId9: { target: "word/media/a.png", external: false } },
    media: { "word/media/a.png": "blob:pic" },
    counters: {},
    para: docx._pProps,
    run: docx._props,
  };
  const html = docx._para(
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:drawing><wp:extent cx="9525" cy="9525"/><a:blip r:embed="rId9"/></w:drawing></w:r>' +
      '<w:r><w:drawing><wp:extent cx="9525" cy="9525"/><a:blip r:embed="rId9"/></w:drawing></w:r></w:p>',
    ctx
  );
  assert.match(html, /display:inline-block/, "the picture flows inline");
  assert.equal((html.match(/<img /g) || []).length, 2, "both pictures are emitted");
  assert.match(html, /text-align:center/, "the paragraph still centres them");
});

// insideH/insideV are the lines *between* rows and columns, so they belong on
// the cells. Putting them on <table> leaves every inner row without a line.
test("docx: table inside borders land on the cells", () => {
  const ctx = { styles: {}, numbering: { nums: {}, abstract: {} }, rels: {}, media: {}, counters: {}, para: () => ({}), run: docx._props };
  const html = docx._table(
    "<w:tbl><w:tblPr><w:tblBorders>" +
      '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>' +
      '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
      '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>' +
      "</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>" +
      "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>" +
      "<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    ctx
  );
  assert.equal((html.match(/<td[^>]*border-bottom/g) || []).length, 4, "every cell carries a bottom line");
  assert.equal((html.match(/<td[^>]*border-right/g) || []).length, 4, "every cell carries a right line");
  assert.equal((html.match(/<td[^>]*border-top/g) || []).length, 2, "the frame's top sits on the first row");
  assert.equal((html.match(/<td[^>]*border-left/g) || []).length, 2, "the frame's left sits on the first column");
  // the table element's own borders are ignored by the collapsing model, so the
  // frame must not be left there alone
  const corner = html.match(/<td style="([^"]*)"/)[1];
  assert.equal(corner.includes("border-top") && corner.includes("border-left"), true, "the top-left corner cell carries both");
});
