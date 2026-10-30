(() => {
  "use strict";

  const HFP = globalThis.HFP;
  if (!HFP) return;
  const { detectKind, fileName, formatSize, mimeFor, decodeText, markdown, rawUrl } = HFP;

  const BUILD = "";        // filled in by build.sh, shows in the card
  const INTENT_MS = 300;   // hover this long before fetching
  const LEAVE_MS = 200;    // grace period to cross the gap from link to card
  const CACHE_MAX = 30;    // blobs kept in memory (LRU)
  const LIMITS = {
    image: 8 * 1024 * 1024,
    pdf: 25 * 1024 * 1024,
    text: 2 * 1024 * 1024,
    sheet: 10 * 1024 * 1024,
    docx: 20 * 1024 * 1024,
    pptx: 30 * 1024 * 1024
  };
  const MAX_SHEET_ROWS = 2000;  // ponytail: truncate, real pagination if anyone cares
  const MAX_SLIDES = 100;       // ponytail: same

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

  const cache = new Map(); // url -> { kind, blobUrl?, text?, size }
  const blocked = new Set(); // urls this page can never read (no CORS header)

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

    const head = document.createElement("div");
    head.className = "head";

    nameEl = document.createElement("span");
    nameEl.className = "name";

    metaEl = document.createElement("span");
    metaEl.className = "meta";

    openEl = document.createElement("a");
    openEl.className = "open";
    openEl.target = "_blank";
    openEl.rel = "noopener noreferrer";
    openEl.textContent = "Open";

    bodyEl = document.createElement("div");
    bodyEl.className = "body";

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
    if (host) host.remove();
    host = null;
    shadow = null;
    cardEl = null;
    nameEl = metaEl = openEl = bodyEl = null;
    currentUrl = null;
  }

  // size + build in the header; every kind shows it, including the deck
  function showMeta(size) {
    if (!metaEl) return;
    let m = size ? formatSize(size) : "";
    if (BUILD) m = (m ? m + " \u00b7 " : "") + BUILD;
    metaEl.textContent = m;
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
    p.textContent = msg;
    bodyEl.appendChild(p);
  }

  function reveal() {
    if (!cardEl) return;
    cardEl.style.visibility = "";
    place();
  }

  let placed = false;  // the card is placed once and then never moves

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

  async function render(name, kind, entry, href) {
    if (!bodyEl) return;
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
      const pane = document.createElement("div");
      pane.className = "pane";
      const select = (i) => {
        Array.from(tabs.children).forEach((b, j) => b.setAttribute("aria-selected", j === i ? "true" : "false"));
        pane.innerHTML = entry.sheets[i].html;
      };
      entry.sheets.forEach((s, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "tab";
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
      const pics = globalThis.__docxPictures;
      if (pics && pics.media && !pics.imgs && metaEl) {
        metaEl.textContent = (metaEl.textContent ? metaEl.textContent + " \u00b7 " : "") + "pictures could not be read";
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
            blocked + " picture" + (blocked > 1 ? "s" : "") + " blocked by this page";
        }
      }, 400);
      openEl.href = href;
    } else if (kind === "pptx") {
      const maxH = Math.max(280, Math.floor(innerHeight * 0.7) - 40);
      const boxW = 860;
      const boxH = Math.min(maxH, Math.round(boxW * 9 / 16));

      const holder = document.createElement("div");
      holder.className = "ppt-scroll";
      // the deck is built one slide at a time in this holder, so it must not be
      // visible while that happens: the card would flash through every slide
      // (and end on the last one) before the stack is assembled
      holder.style.visibility = "hidden";
      bodyEl.appendChild(holder);

      const note = document.createElement("div");
      note.className = "status";
      note.textContent = "Rendering slides…";
      bodyEl.insertBefore(note, holder);

      const lib = globalThis.pptxPreview;
      if (!lib) {
        setStatus("Preview unavailable");
        reveal();
        return;
      }
      try {
        // mode "slide" keeps exactly one slide in the wrapper, so each clone
        // below is a single slide (list mode renders the whole deck at once).
        const viewer = lib.init(holder, { width: boxW, height: boxH, mode: "slide" });
        await viewer.preview(entry.buf.slice(0));

        const count = viewer.slideCount || 1;
        const stack = document.createElement("div");
        stack.className = "ppt-stack";
        let slideH = boxH;
        for (let i = 0; i < count; i++) {
          viewer.renderSingleSlide(i);
          await new Promise((r) => requestAnimationFrame(r));
          // a chart is drawn in a timer of its own, so give it a beat before the
          // slide is copied - otherwise the copy carries an empty chart box
          if (holder.querySelector(".chart-node")) {
            await new Promise((r) => setTimeout(r, 120));
          }
          const slide = holder.querySelector(".pptx-preview-slide-wrapper-" + i);
          if (!slide) continue;
          const clone = slide.cloneNode(true);
          clone.className = "ppt-viewer";
          clone.style.position = "relative";
          clone.style.margin = "0";
          stack.appendChild(clone);
          slideH = clone.getBoundingClientRect().height || slideH;
        }
        holder.innerHTML = "";
        holder.appendChild(stack);
        holder.style.height = slideH * count + "px";
        holder.style.visibility = "";
        note.remove();
        bodyEl.scrollTop = 0;
        host.setAttribute("data-hfp-slides", count);

        currentViewer = { destroy: () => { try { viewer.destroy(); } catch (e) {} } };
        host.setAttribute("data-hfp-renderer", "pptx");
        openEl.href = href;
        showMeta(entry.size);
        place();
        clampCard();
        return;
      } catch (e) {
        console.warn("[hover-file-preview] pptx render failed:", e);
        if (host) host.setAttribute("data-hfp-error", String((e && e.message) || e));
        setStatus("Preview unavailable (" + String((e && e.message) || "render error").slice(0, 80) + ")");
        reveal();
        return;
      }
    } else {
      if (kind === "image") {
        const img = document.createElement("img");
        img.alt = name;
        img.onerror = () => {
          setStatus("Preview failed (blocked by this page?) — click to open");
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
      metaEl.textContent += " \u00b7 first " + MAX_SHEET_ROWS + " rows";
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
  async function buildDocxHtml(buf) {
    const lib = globalThis.HFPdocx;
    globalThis.__docxMediaUrls = null;   // only this render's pictures may be cached
    if (lib) {
      try {
        const html = await lib.render(buf);
        if (html && html.trim()) return html;
      } catch (e) {
        console.warn("[hover-file-preview] docx render failed:", e);
      }
    }

    const zipLib = globalThis.JSZip;
    if (!zipLib) throw Object.assign(new Error("no zip lib"), { code: "PARSE" });
    const zip = await zipLib.loadAsync(buf);
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
    return { buf: buf.slice(0) };
  }

  /* ---------- cache ---------- */

  function cacheGet(url) {
    const hit = cache.get(url);
    if (!hit) return null;
    cache.delete(url);
    cache.set(url, hit);
    return hit;
  }

  function cachePut(url, entry) {
    const prev = cache.get(url);      // a re-render replaces the old pictures
    if (prev && prev !== entry) {
      cache.delete(url);
      if (prev.blobUrl) URL.revokeObjectURL(prev.blobUrl);
      if (prev.mediaUrls) prev.mediaUrls.forEach((u) => URL.revokeObjectURL(u));
    }
    cache.set(url, entry);
    if (cache.size > CACHE_MAX) {
      const [oldUrl, old] = cache.entries().next().value;
      cache.delete(oldUrl);
      if (old.blobUrl) URL.revokeObjectURL(old.blobUrl);
      if (old.mediaUrls) old.mediaUrls.forEach((u) => URL.revokeObjectURL(u));
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

  async function fetchCapped(href, limit) {
    let res;
    try {
      res = await fetch(href, { credentials: "same-origin" });
    } catch (e) {
      // a page CSP or a cross-origin link never reaches the network
      throw Object.assign(new Error("blocked"), { code: "BLOCKED" });
    }
    if (!res.ok) throw Object.assign(new Error("HTTP " + res.status), { code: "HTTP", status: res.status });

    const declared = Number(res.headers.get("content-length") || 0);
    if (declared > limit) {
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
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    return { buf: out.buffer, size: total };
  }

  /* ---------- open ---------- */

  async function open(anchor, kind, x, y) {
    const href = anchor.href;
    const src = rawUrl(href);      // github.com/<o>/<r>/blob/... -> raw.githubusercontent.com
    const name = fileName(href, location.href);
    if (blocked.has(src)) return;   // known unreadable here: no card at all
    const mine = ++seq;

    anchorX = x;
    anchorY = y;
    currentUrl = href;
    reserveH = kind === "pptx" ? Math.floor(innerHeight * 0.7) : 0;
    reserveW = kind === "pptx" ? 900 : 0;   // the card's max width
    placed = false;

    ensureCard();
    nameEl.textContent = name;
    nameEl.title = name;
    metaEl.textContent = "";
    openEl.href = href;
    setStatus("Loading…");
    place();

    const cached = cacheGet(src);
    if (cached) {
      render(name, kind, cached, href).catch(() => {}).then(reveal);
      return;
    }

    let got;
    try {
      got = await fetchCapped(src, LIMITS[kind]);
    } catch (err) {
      if (mine !== seq || currentUrl !== href) return;
      if (err.code === "TOO_LARGE") {
        setStatus("File too large (" + formatSize(err.size) + ") — click to open");
      } else if (err.code === "BLOCKED") {
        // a file on another origin that sends no CORS header can never be read
        // from here, so do not pop a card for it at all
        if (crossOrigin(src)) {
          blocked.add(src);
          hide();
          return;
        }
        setStatus("Blocked by this page (CSP) — click to open");
      } else if (err.code === "HTTP") {
        setStatus("HTTP " + err.status + " — click to open");
      } else {
        setStatus("Preview unavailable");
      }
      if (host) host.setAttribute("data-hfp-error", err.code || "fetch");
      reveal();
      return;
    }

    if (mine !== seq || currentUrl !== href) return;

    let entry;
    if (kind === "text") {
      entry = { kind, text: decodeText(got.buf), size: got.size };
    } else if (kind === "sheet") {
      try {
        entry = { kind, sheets: await buildSheets(got.buf), size: got.size };
        entry.truncated = entry.sheets.some((s) => s.truncated);
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus("Preview unavailable");
        reveal();
        return;
      }
    } else if (kind === "docx") {
      try {
        entry = { kind, html: await buildDocxHtml(got.buf), size: got.size };
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus("Preview unavailable");
        reveal();
        return;
      }
      // the rendered HTML embeds one blob URL per picture; the cache has to
      // keep them alive for as long as it keeps the HTML
      entry.mediaUrls = (globalThis.__docxMediaUrls || []).slice();
      globalThis.__docxMediaUrls = null;
      if (mine !== seq || currentUrl !== href) return;
    } else if (kind === "pptx") {
      try {
        entry = { kind, ...(await buildSlides(got.buf)), size: got.size };
      } catch (err) {
        if (mine !== seq || currentUrl !== href) return;
        setStatus("Preview unavailable");
        reveal();
        return;
      }
      if (mine !== seq || currentUrl !== href) return;
    } else {
      const blob = new Blob([got.buf], { type: mimeFor(kind, name) });
      entry = { kind, blobUrl: URL.createObjectURL(blob), size: got.size };
    }
    cachePut(src, entry);
    render(name, kind, entry, href).catch(() => {}).then(reveal);
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
    if (ev.key === "Escape" && currentUrl) { hide(); return; }
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
