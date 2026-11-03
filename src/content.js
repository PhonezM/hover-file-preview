(() => {
  "use strict";

  const HFP = globalThis.HFP;
  if (!HFP) return;
  const { detectKind, fileName, formatSize, mimeFor, decodeText, markdown, rawUrl } = HFP;

  const BUILD = "";        // filled in by build.sh, shows in the card
  const INTENT_MS = 300;   // hover this long before fetching
  const LEAVE_MS = 200;    // grace period to cross the gap from link to card
  const CACHE_MAX = 30;    // blobs kept in memory (LRU)
  const CACHE_BYTES = 64 * 1024 * 1024; // ponytail: approximate retained size, not exact heap
  const LIMITS = {
    image: 8 * 1024 * 1024,
    pdf: 25 * 1024 * 1024,
    text: 2 * 1024 * 1024,
    sheet: 10 * 1024 * 1024,
    docx: 20 * 1024 * 1024,
    pptx: 30 * 1024 * 1024
  };
  const MAX_SHEET_ROWS = 2000;  // ponytail: truncate, real pagination if anyone cares
  const MAX_SLIDES = 100;       // parsed and rendered per deck; the card says so

  /* UI copy comes from _locales/<lang>/messages.json; the fallback keeps the
     source usable when the i18n API is absent (e.g. a bare test page). */
  function t(key, fallback, subs) {
    try {
      const msg = typeof chrome !== "undefined" && chrome.i18n && chrome.i18n.getMessage(key, subs);
      return msg || fallback;
    } catch (e) {
      return fallback;
    }
  }

  /* The 685 KB renderer is injected on demand into this tab's isolated world.
     A dynamic import is not available to a Firefox MV3 content script
     (Bugzilla 1803950) and a page-world script would obey the page's CSP, so
     the background script uses scripting.executeScript; see src/background.js. */
  let rendererPromise = null;
  function loadRenderer() {
    if (globalThis.pptxPreview) return Promise.resolve(globalThis.pptxPreview);
    if (rendererPromise) return rendererPromise;
    rendererPromise = new Promise((resolve) => {
      let settled = false;
      const done = () => {
        if (settled) return;
        settled = true;
        resolve(globalThis.pptxPreview || null);
      };
      try {
        chrome.runtime.sendMessage({ type: "hfp:load-renderer" }, () => {
          void chrome.runtime.lastError;   // no receiver: treat as unavailable
          done();
        });
      } catch (e) {
        done();
        return;
      }
      setTimeout(done, 3000);   // a missing background must not freeze the card
    });
    rendererPromise.then((lib) => { if (!lib) rendererPromise = null; });
    return rendererPromise;
  }

  const CSS = `
    :host { all: initial; }
    .card {
      position: fixed;
      z-index: 2147483647;
      display: flex;
      flex-direction: column;
      min-width: 240px;
      max-width: 900px;
      max-height: 70vh;
      background: #fff;
      color: #172b4d;
      border: 1px solid rgba(9,30,66,.25);
      border-radius: 6px;
      box-shadow: 0 8px 24px rgba(9,30,66,.28);
      font: 13px/1.45 -apple-system, "Segoe UI", Roboto, sans-serif;
      overflow: hidden;
    }
    .head {
      display: flex; align-items: center; gap: 8px;
      padding: 6px 10px;
      background: #fafbfc;
      border-bottom: 1px solid #ebecf0;
      flex: 0 0 auto;
    }
    .name { flex: 1 1 auto; min-width: 0; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .meta { flex: 0 0 auto; color: #6b778c; font-size: 11px; white-space: nowrap; }
    .open { flex: 0 0 auto; color: #0052cc; font-size: 11px; text-decoration: none; white-space: nowrap; }
    .open:hover { text-decoration: underline; }
    .body { flex: 1 1 auto; overflow: auto; min-height: 0; overflow-anchor: none; }
    .body img { display: block; max-width: 100%; height: auto; }
    .body embed { display: block; width: 100%; height: 60vh; min-height: 320px; }
    .body pre { margin: 0; padding: 10px; white-space: pre-wrap; word-break: break-word;
                font: 12px/1.5 ui-monospace, Consolas, "Courier New", monospace; }
    .tabs { position: sticky; top: 0; z-index: 1; display: flex; flex-wrap: wrap; gap: 4px;
            padding: 5px 8px; background: #fafbfc; border-bottom: 1px solid #ebecf0; }
    .tab { flex: 0 0 auto; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
           font: inherit; font-size: 12px; color: #42526e; background: none; cursor: pointer;
           border: 1px solid transparent; border-radius: 4px; padding: 3px 8px; }
    .tab:hover { background: #ebecf0; }
    .tab[aria-selected="true"] { background: #fff; border-color: #dfe1e6; color: #0052cc; font-weight: 600; }
    .pane { padding: 0; }
    .pane table { border-collapse: collapse; font-size: 12px; }
    .pane td { border: 1px solid #e4e6ea; padding: 2px 7px; vertical-align: top;
               white-space: pre-wrap; overflow-wrap: break-word; }
    .prose { padding: 10px 12px; }
    .docx { padding: 18px 26px; font: 11pt/1.35 Calibri, Carlito, "Segoe UI", sans-serif; color: #000; }
    .docx p { margin: 0 0 6px; }
    .docx h1, .docx h2, .docx h3, .docx h4, .docx h5, .docx h6 { margin: 12px 0 6px; line-height: 1.2; }
    .docx .docx-marker { display: inline-block; min-width: 1.6em; color: #333; }
    .docx .docx-img { max-width: 100%; height: auto; display: block; margin: 4px 0; }
    .docx .docx-page { position: relative; background: #fff; border: 1px solid #e4e6ea;
                       box-shadow: 0 1px 3px rgba(9,30,66,.08); padding: 18px 26px; margin: 0 0 12px; }
    .docx .docx-page:last-child { margin-bottom: 0; }
    .docx .docx-page-no { position: absolute; top: 6px; right: 10px; font: 9pt/1 system-ui, sans-serif;
                          color: #97a0af; }
    /* Word fills a column to the bottom of the page before starting the next
       one, so two short paragraphs stay in column one */
    .docx .docx-cols { column-gap: 26px; column-fill: auto; min-height: 420px; }
    .docx .docx-header, .docx .docx-footer { color: #5e6c84; font-size: 9pt; }
    .docx .docx-header { border-bottom: 1px solid #ebecf0; padding-bottom: 4px; margin-bottom: 10px; }
    .docx .docx-footer { border-top: 1px solid #ebecf0; padding-top: 4px; margin-top: 12px; }
    .docx .docx-footnotes { border-top: 1px solid #ebecf0; margin-top: 14px; padding-top: 6px; font-size: 9pt; color: #42526e; }
    .docx .docx-footnote { margin: 0 0 4px; }
    .docx table.docx-table { margin: 6px 0 10px; font-size: 10pt; }
    .docx table.docx-table td { padding: 2px 5px; vertical-align: top; }
    .docx a { color: #0052cc; }
    .md { padding: 10px 14px; line-height: 1.5; }
    .md h1, .md h2, .md h3, .md h4, .md h5, .md h6 { margin: 14px 0 6px; line-height: 1.25; }
    .md h1 { font-size: 20px; border-bottom: 1px solid #ebecf0; padding-bottom: 4px; }
    .md h2 { font-size: 17px; }
    .md h3 { font-size: 15px; }
    .md h4, .md h5, .md h6 { font-size: 13px; }
    .md > :first-child { margin-top: 0; }
    .md p { margin: 0 0 8px; }
    .md ul, .md ol { margin: 0 0 8px; padding-left: 22px; }
    .md li { margin: 2px 0; }
    .md li > p { margin: 0; }
    .md code { background: #f4f5f7; padding: 1px 4px; border-radius: 3px;
               font: 12px/1.45 ui-monospace, Consolas, "Courier New", monospace; }
    .md pre { background: #f7f8f9; border: 1px solid #ebecf0; border-radius: 4px;
              padding: 8px 10px; margin: 0 0 10px; overflow: auto; }
    .md pre code { background: none; padding: 0; white-space: pre; }
    .md blockquote { margin: 0 0 10px; padding: 2px 10px; border-left: 3px solid #dfe1e6; color: #505f79; }
    .md hr { border: 0; border-top: 1px solid #ebecf0; margin: 12px 0; }
    .md a { color: #0052cc; }
    .md img { max-width: 100%; height: auto; }
    .md table { border-collapse: collapse; margin: 6px 0 10px; font-size: 12px; }
    .md th, .md td { border: 1px solid #e4e6ea; padding: 3px 8px; text-align: left; vertical-align: top; }
    .md th { background: #fafbfc; }
    .prose p { margin: 0 0 6px; }
    .prose table { border-collapse: collapse; }
    .prose td, .prose th { border: 1px solid #e4e6ea; padding: 2px 6px; vertical-align: top; }
    .prose img { max-width: 100%; height: auto; }
    .slides { padding: 6px 12px 10px; }
    .ppt-scroll { width: 100%; max-width: 900px; background: #fff; overflow: auto; }
    .ppt-stack { position: relative; }
    .ppt-viewer { position: relative; width: 100%; background: #fff; overflow: hidden; }

    .slide { padding: 6px 0; border-bottom: 1px solid #f0f1f3; }
    .slide:last-child { border-bottom: 0; }
    .slide-no { color: #0052cc; font-weight: 600; font-size: 11px; margin-bottom: 3px; }
    .slide p { margin: 0 0 3px; }
    .slide p.empty { color: #97a0af; font-style: italic; }
    .slide table { border-collapse: collapse; margin: 4px 0 7px; font-size: 11.5px; }
    .slide td { border: 1px solid #e4e6ea; padding: 2px 6px; vertical-align: top; }
    .card:focus-visible { outline: 2px solid #0052cc; outline-offset: 2px; }
    .tab:focus-visible, .open:focus-visible { outline: 2px solid #0052cc; outline-offset: 1px; }
    .status { padding: 16px 14px; color: #6b778c; }
  `;

  let host = null;
  let shadow = null;
  let cardEl = null;
  let nameEl = null;
  let metaEl = null;
  let openEl = null;
  let bodyEl = null;

  let currentUrl = null;
  let anchorX = 0;
  let anchorY = 0;
  let currentViewer = null;
  let reserveH = 0;   // room reserved for the card, so it does not move
  let reserveW = 0;   // width reserved for the card, so it does not move
  let intentTimer = 0;
  let closeTimer = 0;
  let seq = 0;
  let abortCtrl = null;   // the one fetch/render this card owns

  const cache = new Map(); // url -> { kind, blobUrl?, text?, size, cost }
  const blocked = new Set(); // urls this page can never read (no CORS header)
  let cacheBytes = 0;        // sum of entry.cost, kept below CACHE_BYTES

  /* ---------- card ---------- */

  function ensureCard() {
    if (host) return;
    host = document.createElement("div");
    host.setAttribute("data-hover-file-preview", "");
    if (BUILD) host.setAttribute("data-hfp-build", BUILD);
    shadow = host.attachShadow({ mode: "closed" });

    const style = document.createElement("style");
    style.textContent = CSS;

    cardEl = document.createElement("div");
    cardEl.className = "card";
    cardEl.setAttribute("role", "dialog");
    cardEl.setAttribute("aria-label", t("preview", "File preview"));
    cardEl.tabIndex = -1;

    const head = document.createElement("div");
    head.className = "head";

    nameEl = document.createElement("span");
    nameEl.className = "name";

    metaEl = document.createElement("span");
    metaEl.className = "meta";
    metaEl.setAttribute("aria-live", "polite");

    openEl = document.createElement("a");
    openEl.className = "open";
    openEl.target = "_blank";
    openEl.rel = "noopener noreferrer";
    openEl.textContent = t("open", "Open");

    bodyEl = document.createElement("div");
    bodyEl.className = "body";
    bodyEl.setAttribute("aria-busy", "true");

    head.append(nameEl, metaEl, openEl);
    cardEl.append(head, bodyEl);
    shadow.append(style, cardEl);
    document.documentElement.appendChild(host);
    watchSize();
  }

  function hide() {
    // the blob URLs of a rendered docx belong to the cached entry: revoking
    // them here would leave that HTML pointing at dead URLs, so the pictures
    // come back blank (white boxes) on the next hover. cachePut revokes them
    // when the entry is evicted instead.
    if (sizeWatch) { sizeWatch.disconnect(); sizeWatch = null; }
    if (currentViewer) { currentViewer.destroy(); currentViewer = null; }
    clearTimeout(intentTimer);
    clearTimeout(closeTimer);
    cancelFetch();     // a download for a card nobody is looking at is wasted
    seq++;             // and every parse/render already in flight is stale
    if (host) host.remove();
    host = null;
    shadow = null;
    cardEl = null;
    nameEl = metaEl = openEl = bodyEl = null;
    currentUrl = null;
    focusCard = false;
  }

  // size + build in the header; every kind shows it, including the deck
  function showMeta(size) {
    if (!metaEl) return;
    let m = size ? formatSize(size) : "";
    if (BUILD) m = (m ? m + " \u00b7 " : "") + BUILD;
    metaEl.textContent = m;
  }

  // append a qualifier ("first 100 of 240 slides") without losing the size
  function metaNote(text) {
    if (!metaEl) return;
    metaEl.textContent = (metaEl.textContent ? metaEl.textContent + " \u00b7 " : "") + text;
  }

  // keep the card inside the window: a preview that turns out taller than the
  // room it was placed with moves up (never down and never sideways across the
  // pointer), so it can not end up half off screen
  function clampCard() {
    if (!cardEl) return;
    const r = cardEl.getBoundingClientRect();
    const pad = 8;
    let top = r.top;
    let left = r.left;
    if (r.bottom > innerHeight - pad) top = Math.max(pad, innerHeight - r.height - pad);
    if (r.right > innerWidth - pad) left = Math.max(pad, innerWidth - r.width - pad);
    if (top !== r.top) cardEl.style.top = top + "px";
    if (left !== r.left) cardEl.style.left = left + "px";
  }

  let sizeWatch = null;

  // the card's content arrives in its own time (an image, a pdf embed, a docx
  // rendered in pieces): clamp whenever it actually changes size
  function watchSize() {
    if (!cardEl || typeof ResizeObserver === "undefined") return;
    if (sizeWatch) sizeWatch.disconnect();
    sizeWatch = new ResizeObserver(() => clampCard());
    sizeWatch.observe(cardEl);
  }

  function setStatus(msg) {
    if (!bodyEl) return;
    bodyEl.textContent = "";
    const p = document.createElement("div");
    p.className = "status";
    p.setAttribute("role", "status");
    p.textContent = msg;
    bodyEl.appendChild(p);
  }

  function reveal() {
    if (!cardEl) return;
    cardEl.style.visibility = "";
    if (bodyEl) bodyEl.setAttribute("aria-busy", "false");
    place();
    // a card opened with the keyboard takes the focus, so Tab reaches its
    // links instead of walking the rest of the page first
    if (focusCard) {
      focusCard = false;
      try { cardEl.focus({ preventScroll: true }); } catch (e) { /* older engine */ }
    }
  }

  let placed = false;  // the card is placed once and then never moves
  let focusCard = false; // this card was opened with the keyboard

  function place() {
    if (!cardEl) return;
    if (placed) return;   // the position is final: the card only grows in place
    cardEl.style.left = "-99999px";
    cardEl.style.top = "0";
    const r = cardEl.getBoundingClientRect();
    const pad = 8;
    // reserve the room the card can grow into, so the position does not change
    // when the preview arrives
    const room = reserveH || Math.min(Math.floor(innerHeight * 0.7), r.height || 0);
    const width = reserveW || r.width || 0;
    let left = anchorX + 16;
    let top = anchorY + 16;
    if (left + width > innerWidth - pad) left = Math.max(pad, anchorX - width - 16);
    if (top + room > innerHeight - pad) top = Math.max(pad, anchorY - room - 16);
    cardEl.style.left = left + "px";
    cardEl.style.top = top + "px";
    placed = true;
  }

  async function render(name, kind, entry, href, token) {
    if (!bodyEl || (token !== undefined && token !== seq)) return;
    bodyEl.textContent = "";

    if (kind === "text") {
      if (/\.(md|markdown|mdown|mkd)$/i.test(name)) {
        const wrap = document.createElement("div");
        wrap.className = "md";
        wrap.innerHTML = markdown(entry.text);
        bodyEl.appendChild(wrap);
      } else {
        const pre = document.createElement("pre");
        pre.textContent = entry.text;
        bodyEl.appendChild(pre);
      }
      openEl.href = href;
    } else if (kind === "sheet") {
      const tabs = document.createElement("div");
      tabs.className = "tabs";
      tabs.setAttribute("role", "tablist");
      const pane = document.createElement("div");
      pane.className = "pane";
      pane.id = "hfp-pane";
      pane.setAttribute("role", "tabpanel");
      const select = (i) => {
        Array.from(tabs.children).forEach((b, j) => b.setAttribute("aria-selected", j === i ? "true" : "false"));
        pane.setAttribute("aria-labelledby", "hfp-tab-" + i);
        pane.innerHTML = entry.sheets[i].html;
      };
      entry.sheets.forEach((s, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "tab";
        b.id = "hfp-tab-" + i;
        b.setAttribute("role", "tab");
        b.setAttribute("aria-controls", "hfp-pane");
        b.textContent = s.name;
        b.title = s.name;
        b.addEventListener("click", () => select(i));
        tabs.appendChild(b);
      });
      bodyEl.append(tabs, pane);
      select(0);
      openEl.href = href;
    } else if (kind === "docx") {
      const wrap = document.createElement("div");
      wrap.className = "prose docx-body";
      wrap.innerHTML = entry.html;
      bodyEl.appendChild(wrap);
      // what the file carried vs what the card got, so a picture that never
      // arrives is named instead of leaving a gap
      const pics = entry.pictures;
      if (pics && pics.media && !pics.imgs && metaEl) {
        metaNote(t("picturesUnreadable", "pictures could not be read"));
      }
      // a page CSP that forbids data:/blob: images drops every picture without
      // an error we could catch - say so in the card instead of showing gaps
      setTimeout(() => {
        if (!wrap.isConnected) return;
        const imgs = Array.from(wrap.querySelectorAll("img"));
        const blocked = imgs.filter((i) => i.complete && i.naturalWidth === 0).length;
        if (host) host.setAttribute("data-hfp-pictures", imgs.length + ":" + (imgs.length - blocked));
        if (!blocked) return;
        if (host) host.setAttribute("data-hfp-imgblocked", String(blocked));
        if (metaEl) {
          metaEl.textContent =
            (metaEl.textContent ? metaEl.textContent + " \u00b7 " : "") +
            t("picturesBlocked", "$1 picture(s) blocked by this page", [String(blocked)]);
        }
      }, 400);
      openEl.href = href;
    } else if (kind === "pptx") {
      const boxW = 860;   // the card's max width

      const holder = document.createElement("div");
      holder.className = "ppt-scroll";
      holder.style.visibility = "hidden";
      bodyEl.appendChild(holder);

      const note = document.createElement("div");
      note.className = "status";
      note.textContent = t("renderingSlides", "Rendering slides…");
      bodyEl.insertBefore(note, holder);

      const lib = await loadRenderer();
      if (token !== seq) return;
      if (!lib) {
        setStatus(t("unavailable", "Preview unavailable"));
        reveal();
        return;
      }
      try {
        // "list" stacks the slides in the wrapper itself. No clone per slide, so
        // a chart drawn in place keeps its canvas and no copy can race it; the
        // old per-slide clone loop plus its 120 ms chart sleep is gone.
        // maxSlides bounds the parse and the DOM, not just the loop.
        const viewer = lib.init(holder, { width: boxW, mode: "list", maxSlides: MAX_SLIDES });
        currentViewer = { destroy: () => { try { viewer.destroy(); } catch (e) {} } };
        await viewer.preview(entry.buf);
        if (token !== seq || !bodyEl) return;

        const total = viewer.totalSlideCount || viewer.slideCount || 1;
        const shown = viewer.slideCount || Math.min(total, MAX_SLIDES);
        // the library paints its wrapper black for slide-mode letterboxing; in
        // list mode that black shows through the 10 px gap between slides
        const wrapper = holder.querySelector(".pptx-preview-wrapper");
        if (wrapper) wrapper.style.background = "#fff";
        holder.style.visibility = "";
        note.remove();
        bodyEl.scrollTop = 0;
        host.setAttribute("data-hfp-slides", String(shown));

        host.setAttribute("data-hfp-renderer", "pptx");
        openEl.href = href;
        showMeta(entry.size);
        if (total > shown) metaNote(t("firstSlides", "first $1 of $2 slides", [String(shown), String(total)]));
        place();
        clampCard();
        return;
      } catch (e) {
        if (token !== seq) return;   // a superseded render must not touch the new card
        console.warn("[hover-file-preview] pptx render failed:", e);
        if (host) host.setAttribute("data-hfp-error", String((e && e.message) || e));
        setStatus(t("renderFailed", "Preview unavailable ($1)", [String((e && e.message) || "render error").slice(0, 80)]));
        reveal();
        return;
      }
    } else {
      if (kind === "image") {
        const img = document.createElement("img");
        img.alt = name;
        img.onerror = () => {
          setStatus(t("imageFailed", "Preview failed (blocked by this page?) — click to open"));
          if (host) host.setAttribute("data-hfp-error", "image");
        };
        img.onload = () => clampCard();   // the picture's height is known now
        img.src = entry.blobUrl;
        bodyEl.appendChild(img);
      } else {
        const em = document.createElement("embed");
        em.type = "application/pdf";
        em.src = entry.blobUrl;
        bodyEl.appendChild(em);
      }
      openEl.href = entry.blobUrl;
    }

    showMeta(entry.size);
    if (kind === "sheet" && entry.truncated) {
      metaNote(t("firstRows", "first $1 rows", [String(MAX_SHEET_ROWS)]));
    }
    if (host) host.setAttribute("data-hfp-renderer", "direct");
    place();
    clampCard();
  }

  /* xlsx -> [{name, html}] through our own reader (values, dates, merges). */
  async function buildSheets(buf) {
    const lib = globalThis.HFPxlsx;
    if (!lib) throw Object.assign(new Error("no xlsx reader"), { code: "PARSE" });
    const sheets = await lib.read(buf, { maxRows: MAX_SHEET_ROWS });
    if (!sheets.length) throw Object.assign(new Error("empty workbook"), { code: "PARSE" });
    return sheets;
  }

  /* docx -> HTML through our own reader, with the text outline as a last resort. */
  async function buildDocxHtml(buf, report) {
    const lib = globalThis.HFPdocx;
    if (lib) {
      try {
        const html = await lib.render(buf, report);
        if (html && html.trim()) return html;
      } catch (e) {
        console.warn("[hover-file-preview] docx render failed:", e);
      }
    }

    // the reader is unavailable: keep the text outline rather than no card
    const zipLib = globalThis.HFPZip;
    if (!zipLib) throw Object.assign(new Error("no zip lib"), { code: "PARSE" });
    const zip = await zipLib.load(buf);
    const doc = zip.file("word/document.xml");
    if (!doc) throw Object.assign(new Error("no document.xml"), { code: "PARSE" });

    const holder = document.createElement("div");
    HFP.docxLines(await doc.async("string")).forEach((line) => {
      const p = document.createElement("p");
      p.textContent = line;
      holder.appendChild(p);
    });
    if (!holder.innerHTML.trim()) throw Object.assign(new Error("empty docx"), { code: "PARSE" });
    return holder.innerHTML;
  }

  /* pptx -> the raw buffer; the renderer parses it when the card opens. */
  async function buildSlides(buf) {
    return { buf };
  }

  /* ---------- cache ---------- */

  function cacheGet(url) {
    const hit = cache.get(url);
    if (!hit) return null;
    cache.delete(url);
    cache.set(url, hit);
    return hit;
  }

  function evict(url) {
    const old = cache.get(url);
    if (!old) return;
    cache.delete(url);
    cacheBytes -= old.cost || 0;
    if (old.blobUrl) URL.revokeObjectURL(old.blobUrl);
    if (old.mediaUrls) old.mediaUrls.forEach((u) => URL.revokeObjectURL(u));
  }

  function cachePut(url, entry) {
    if (entry.cost > CACHE_BYTES) return;   // one huge deck must not hold the budget
    evict(url);                              // a re-render replaces the old pictures
    cache.set(url, entry);
    cacheBytes += entry.cost || 0;
    while (cache.size > CACHE_MAX || cacheBytes > CACHE_BYTES) {
      if (!cache.size) break;
      const [oldUrl] = cache.entries().next().value;
      evict(oldUrl);
    }
  }

  /* ---------- fetch ---------- */

  function crossOrigin(href) {
    try {
      return new URL(href, location.href).origin !== location.origin;
    } catch (e) {
      return false;
    }
  }

  function cancelFetch() {
    if (!abortCtrl) return;
    abortCtrl.abort();
    abortCtrl = null;
  }

  async function fetchCapped(href, limit, signal) {
    let res;
    try {
      res = await fetch(href, { credentials: "same-origin", signal });
    } catch (e) {
      if (signal && signal.aborted) {
        throw Object.assign(new Error("aborted"), { code: "ABORTED" });
      }
      // a page CSP or a cross-origin link never reaches the network
      throw Object.assign(new Error("blocked"), { code: "BLOCKED" });
    }
    if (!res.ok) {
      if (res.body) res.body.cancel().catch(() => {});
      throw Object.assign(new Error("HTTP " + res.status), { code: "HTTP", status: res.status });
    }

    const declared = Number(res.headers.get("content-length") || 0);
    if (declared > limit) {
      if (res.body) res.body.cancel().catch(() => {});
      throw Object.assign(new Error("too large"), { code: "TOO_LARGE", size: declared });
    }
    if (!res.body) {
      const buf = await res.arrayBuffer();
      if (buf.byteLength > limit) {
        throw Object.assign(new Error("too large"), { code: "TOO_LARGE", size: buf.byteLength });
      }
      return { buf, size: buf.byteLength };
    }

    const reader = res.body.getReader();
    const chunks = [];
    let total = 0;
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        total += value.byteLength;
        if (total > limit) {
          reader.cancel().catch(() => {});
          throw Object.assign(new Error("too large"), { code: "TOO_LARGE", size: total });
        }
        chunks.push(value);
      }
    } catch (err) {
      if (signal && signal.aborted) {
        reader.cancel().catch(() => {});
        throw Object.assign(new Error("aborted"), { code: "ABORTED" });
      }
      throw err;
    }
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    return { buf: out.buffer, size: total };
  }

  /* ---------- open ---------- */

  async function open(anchor, kind, x, y, viaKeyboard) {
    const href = anchor.href;
    const src = rawUrl(href);      // github.com/<o>/<r>/blob/... -> raw.githubusercontent.com
    const name = fileName(href, location.href);
    if (blocked.has(src)) return;   // known unreadable here: no card at all
    cancelFetch();                  // a card that was replaced must not keep downloading
    abortCtrl = new AbortController();
    const signal = abortCtrl.signal;
    const mine = ++seq;

    anchorX = x;
    anchorY = y;
    currentUrl = href;
    reserveH = kind === "pptx" ? Math.floor(innerHeight * 0.7) : 0;
    reserveW = kind === "pptx" ? 900 : 0;   // the card's max width
    placed = false;
    focusCard = !!viaKeyboard;

    if (currentViewer) { currentViewer.destroy(); currentViewer = null; }
    ensureCard();
    nameEl.textContent = name;
    nameEl.title = name;
    metaEl.textContent = "";
    openEl.href = href;
    setStatus(t("loading", "Loading…"));
    place();

    const cached = cacheGet(src);
    if (cached) {
      render(name, kind, cached, href, mine).catch(() => {}).then(() => { if (mine === seq) reveal(); });
      return;
    }

    let got;
    try {
      got = await fetchCapped(src, LIMITS[kind], signal);
    } catch (err) {
      if (mine !== seq || currentUrl !== href) return;
      if (err.code === "ABORTED") return;   // the card was replaced or closed
      if (err.code === "TOO_LARGE") {
        setStatus(t("tooLarge", "File too large ($1) — click to open", [formatSize(err.size)]));
      } else if (err.code === "BLOCKED") {
        // a file on another origin that sends no CORS header can never be read
        // from here, so do not pop a card for it at all
        if (crossOrigin(src)) {
          blocked.add(src);
          hide();
          return;
        }
        setStatus(t("blockedPage", "Blocked by this page (CSP) — click to open"));
      } else if (err.code === "HTTP") {
        setStatus(t("httpError", "HTTP $1 — click to open", [String(err.status)]));
      } else {
        setStatus(t("unavailable", "Preview unavailable"));
      }
      if (host) host.setAttribute("data-hfp-error", err.code || "fetch");
      reveal();
      return;
    }

    if (mine !== seq || currentUrl !== href) return;

    let entry;
    if (kind === "text") {
      const text = decodeText(got.buf);
      entry = { kind, text, size: got.size, cost: Math.max(got.size, text.length * 2) };
    } else if (kind === "sheet") {
      try {
        const sheets = await buildSheets(got.buf);
        entry = {
          kind,
          sheets,
          size: got.size,
          cost: sheets.reduce((n, s) => n + s.html.length * 2, 0),
        };
        entry.truncated = entry.sheets.some((s) => s.truncated);
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus(t("unavailable", "Preview unavailable"));
        reveal();
        return;
      }
    } else if (kind === "docx") {
      // the report is per render: globals let a superseded render hand its
      // pictures to the next card
      const report = {};
      try {
        const html = await buildDocxHtml(got.buf, report);
        entry = {
          kind,
          html,
          size: got.size,
          mediaUrls: report.mediaUrls || [],
          pictures: report.pictures,
          cost: html.length * 2 + (report.mediaBytes || 0),
        };
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus(t("unavailable", "Preview unavailable"));
        reveal();
        return;
      }
      if (mine !== seq || currentUrl !== href) return;
    } else if (kind === "pptx") {
      try {
        entry = { kind, ...(await buildSlides(got.buf)), size: got.size, cost: got.size };
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus(t("unavailable", "Preview unavailable"));
        reveal();
        return;
      }
      if (mine !== seq || currentUrl !== href) return;
    } else {
      const blob = new Blob([got.buf], { type: mimeFor(kind, name) });
      entry = { kind, blobUrl: URL.createObjectURL(blob), size: got.size, cost: got.size };
    }
    cachePut(src, entry);
    render(name, kind, entry, href, mine).catch(() => {}).then(() => { if (mine === seq) reveal(); });
  }

  /* ---------- events ---------- */

  /* Closed shadow root retargets inner events to the host, so ev.target
     is the card host when the pointer is anywhere inside the card. */
  function overCard(ev) {
    return !!host && ev.target === host;
  }

  function anchorFrom(ev) {
    const path = ev.composedPath ? ev.composedPath() : [];
    for (const node of path) {
      if (node instanceof HTMLAnchorElement && node.href) return node;
    }
    return null;
  }

  let pointerX = 0;
  let pointerY = 0;

  function pointerInCard() {
    if (!cardEl) return false;
    const r = cardEl.getBoundingClientRect();
    return pointerX >= r.left && pointerX <= r.right && pointerY >= r.top && pointerY <= r.bottom;
  }

  function onOver(ev) {
    pointerX = ev.clientX;
    pointerY = ev.clientY;
    clearTimeout(closeTimer); // any movement cancels a pending close

    if (overCard(ev)) {
      clearTimeout(intentTimer); // keep the card while the pointer is on it
      return;
    }
    const anchor = anchorFrom(ev);
    if (anchor) {
      if (anchor.href === currentUrl) return;
      const kind = detectKind(anchor.href, location.href);
      if (!kind) {
        if (currentUrl) hide();
        return;
      }
      clearTimeout(intentTimer);
      const x = ev.clientX;
      const y = ev.clientY;
      intentTimer = setTimeout(() => open(anchor, kind, x, y), INTENT_MS);
      return;
    }

    clearTimeout(intentTimer);
    if (currentUrl) closeTimer = setTimeout(() => { if (!pointerInCard()) hide(); }, LEAVE_MS);
  }

  // a marker on the page: proves the content script is loaded, and which build
  document.documentElement.setAttribute("data-hfp-cs", BUILD || "dev");

  document.addEventListener("mouseover", onOver, true);

  document.addEventListener("keydown", (ev) => {
    if (ev.key !== "Escape") return;
    clearTimeout(intentTimer);
    if (currentUrl) hide();
  }, true);

  // Keyboard parity: tabbing onto a previewable link opens the same card, and
  // the card takes the focus so its links and sheet tabs are reachable.
  document.addEventListener("focusin", (ev) => {
    if (host && (ev.target === host || host.contains(ev.target))) {
      clearTimeout(closeTimer);
      return;
    }
    const anchor = ev.target instanceof HTMLAnchorElement ? ev.target : null;
    if (!anchor || !anchor.href) { clearTimeout(intentTimer); return; }
    if (anchor.href === currentUrl) return;
    const kind = detectKind(anchor.href, location.href);
    if (!kind) return;
    const r = anchor.getBoundingClientRect();
    if (!r.width && !r.height) return;
    clearTimeout(intentTimer);
    intentTimer = setTimeout(() => open(anchor, kind, r.right, r.bottom, true), INTENT_MS);
  }, true);

  document.addEventListener("focusout", () => {
    if (!currentUrl) return;
    setTimeout(() => {
      if (!currentUrl) return;
      const active = document.activeElement;
      if (host && (active === host || (active && host.contains(active)))) return;
      if (active instanceof HTMLAnchorElement && detectKind(active.href, location.href)) return;
      hide();
    }, LEAVE_MS);
  }, true);

  addEventListener("scroll", (ev) => {
    // only a scroll of the page itself closes the card - scrolling the deck
    // inside the card must not
    if (currentUrl && (ev.target === document || ev.target === document.documentElement)) hide();
  }, { passive: true, capture: true });

  addEventListener("resize", () => {
    if (!cardEl) return;
    // a window that only got shorter (a url bar that hides) must not move the
    // card: only re-place it when it would end up outside the window
    const r = cardEl.getBoundingClientRect();
    if (r.top < 0 || r.left < 0 || r.bottom > innerHeight || r.right > innerWidth) {
      placed = false;
      place();
    }
  }, { passive: true });

  addEventListener("pagehide", hide);
})();
