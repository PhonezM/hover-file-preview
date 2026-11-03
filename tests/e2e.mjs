#!/usr/bin/env node
/* End-to-end check of the packaged extension in Chromium.
 *
 *   node tests/e2e.mjs [--root DIR] [--out DIR] [--label NAME]
 *
 * Boots the unpacked Chrome build in <root>/build/chrome, serves <root>/demo
 * over http and hovers each demo link. A case passes when the card appears
 * with the expected renderer, no error attribute, and (for a deck) a slide
 * count. Every case also writes a screenshot, so the same script produces the
 * before/after pairs.
 *
 * Needs `playwright-core` and a Chromium in the Playwright cache; the build
 * itself is `bash build.sh` (or `npm run build`).
 */
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { chromium } from "playwright-core";

const args = Object.fromEntries(
  process.argv.slice(2).reduce((pairs, arg, i, all) => {
    if (arg.startsWith("--")) pairs.push([arg.slice(2), all[i + 1]]);
    return pairs;
  }, [])
);

const ROOT = path.resolve(args.root || path.join(import.meta.dirname, ".."));
const OUT = path.resolve(args.out || path.join(ROOT, "compare"));
const LABEL = args.label || "run";
const EXT = path.join(ROOT, "build", "chrome");
const DEMO = path.join(ROOT, "demo");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".md": "text/markdown; charset=utf-8",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

const CASES = [
  { link: "demo-deck.pptx", kind: "pptx", renderer: "pptx", settle: 1500 },
  { link: "demo-sheet.xlsx", kind: "sheet", renderer: "direct", settle: 500 },
  { link: "demo-doc.docx", kind: "docx", renderer: "direct", settle: 800 },
  { link: "notes.md", kind: "markdown", renderer: "direct", settle: 400 },
  { link: "demo-photo.png", kind: "image", renderer: "direct", settle: 400 },
  { link: "notes.md", kind: "keyboard", renderer: "direct", settle: 300, focus: true },
];

function fail(msg) {
  console.error("FAIL " + msg);
  process.exitCode = 1;
}

function serve(dir) {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      const rel = decodeURIComponent(req.url.split("?")[0]);
      const file = path.join(dir, rel === "/" ? "index.html" : rel);
      if (!file.startsWith(dir) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
        res.writeHead(404).end("not found");
        return;
      }
      res.writeHead(200, { "content-type": MIME[path.extname(file)] || "application/octet-stream" });
      fs.createReadStream(file).pipe(res);
    });
    server.listen(0, "127.0.0.1", () => resolve(server));
  });
}

async function main() {
  for (const p of [EXT, DEMO]) {
    if (!fs.existsSync(p)) throw new Error("missing " + p + " (run npm run build first)");
  }
  const manifest = JSON.parse(fs.readFileSync(path.join(EXT, "manifest.json"), "utf8"));
  const staticScripts = manifest.content_scripts.flatMap((c) => c.js);
  if (staticScripts.some((f) => f.includes("pptx"))) fail("renderer is still a static content script");
  if (!fs.existsSync(path.join(EXT, "dist", "pptx.js"))) fail("renderer missing from the package");
  fs.mkdirSync(OUT, { recursive: true });
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "hfp-e2e-"));
  const server = await serve(DEMO);
  const url = "http://127.0.0.1:" + server.address().port + "/index.html";
  const ctx = await chromium.launchPersistentContext(profile, {
    headless: true,
    channel: "chromium",
    locale: args.locale || "en",
    viewport: { width: 1200, height: 800 },
    args: ["--disable-extensions-except=" + EXT, "--load-extension=" + EXT, "--lang=" + (args.locale || "en")],
  });

  const results = [];
  try {
    const page = await ctx.newPage();
    await page.goto(url);
    await page.waitForFunction(() => document.documentElement.hasAttribute("data-hfp-cs"), null, { timeout: 10000 });

    for (const c of CASES) {
      await page.goto(url);               // a fresh page drops the previous card
      if (c.focus) await page.focus('a[href="' + c.link + '"]');
      else await page.hover('a[href="' + c.link + '"]');
      let card = null;
      try {
        await page.waitForFunction(
          () => {
            const h = document.querySelector("[data-hover-file-preview]");
            return !!h && (h.hasAttribute("data-hfp-renderer") || h.hasAttribute("data-hfp-error"));
          },
          null,
          { timeout: 20000 }
        );
        card = await page.evaluate(() => {
          const h = document.querySelector("[data-hover-file-preview]");
          return {
            renderer: h.getAttribute("data-hfp-renderer"),
            error: h.getAttribute("data-hfp-error"),
            slides: h.getAttribute("data-hfp-slides"),
            pictures: h.getAttribute("data-hfp-pictures"),
          };
        });
      } catch (e) {
        fail(c.kind + ": no card within 20s");
      }
      await page.waitForTimeout(c.settle);
      const shot = path.join(OUT, LABEL + "-" + c.kind + ".png");
      await page.screenshot({ path: shot });

      if (card) {
        if (card.error) fail(c.kind + ": data-hfp-error=" + card.error);
        if (card.renderer !== c.renderer) fail(c.kind + ": renderer " + card.renderer + " != " + c.renderer);
        if (c.kind === "pptx" && !(Number(card.slides) > 0)) fail("pptx: no slide count");
      }
      results.push({ case: c.kind, ...card, screenshot: shot });
      console.log("PASS " + c.kind + " " + JSON.stringify(card));

      if (c.focus) {
        await page.keyboard.press("Escape");
        const gone = await page.waitForFunction(
          () => !document.querySelector("[data-hover-file-preview]"),
          null,
          { timeout: 3000 }
        ).then(() => true, () => false);
        if (!gone) fail("keyboard: Escape did not close the card");
      }
    }
  } finally {
    await ctx.close();
    server.close();
  }
  console.log(JSON.stringify({ label: LABEL, root: ROOT, results }, null, 2));
}

main().catch((e) => {
  fail(e.stack || String(e));
  process.exit(1);
});
