/* Pure helpers. Loaded as a content script before content.js, and as a
   CommonJS module by test/lib.test.mjs. */
(function (root) {
  "use strict";

  var IMAGE_EXT = /\.(png|jpe?g|gif|webp|bmp|svg)$/i;
  var PDF_EXT = /\.pdf$/i;
  var SHEET_EXT = /\.(xlsx|xlsm)$/i;
  var DOCX_EXT = /\.docx$/i;
  var PPTX_EXT = /\.pptx$/i;
  var TEXT_EXT = /\.(txt|csv|log|md|json|sql)$/i;

  var IMAGE_MIME = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    bmp: "image/bmp",
    svg: "image/svg+xml"
  };

  var TEXT_MIME = {
    json: "application/json",
    csv: "text/csv",
    md: "text/markdown"
  };

  /* Returns "image" | "pdf" | "text" | null for an anchor href. */
  function detectKind(rawHref, base) {
    var u;
    try {
      u = new URL(rawHref, base || (typeof location !== "undefined" ? location.href : undefined));
    } catch (e) {
      return null;
    }
    if (u.protocol !== "http:" && u.protocol !== "https:") return null;

    var last = (u.pathname.split("/").pop() || "");
    var name = last;
    try {
      name = decodeURIComponent(last);
    } catch (e) {
      /* keep the raw segment */
    }
    if (!name) return null;

    if (IMAGE_EXT.test(name)) return "image";
    if (PDF_EXT.test(name)) return "pdf";
    if (SHEET_EXT.test(name)) return "sheet";
    if (DOCX_EXT.test(name)) return "docx";
    if (PPTX_EXT.test(name)) return "pptx";
    if (TEXT_EXT.test(name)) return "text";
    return null;
  }

  /* Filename shown in the card header. */
  function fileName(rawHref, base) {
    var u;
    try {
      u = new URL(rawHref, base || (typeof location !== "undefined" ? location.href : undefined));
    } catch (e) {
      return "";
    }
    var last = (u.pathname.split("/").pop() || "");
    try {
      return decodeURIComponent(last) || last;
    } catch (e) {
      return last;
    }
  }

  function formatSize(n) {
    if (typeof n !== "number" || !isFinite(n) || n < 0) return "";
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / 1024 / 1024).toFixed(1) + " MB";
  }

  function mimeFor(kind, name) {
    if (kind === "pdf") return "application/pdf";
    var ext = String(name || "").toLowerCase().split(".").pop();
    if (kind === "image") return IMAGE_MIME[ext] || "application/octet-stream";
    return (TEXT_MIME[ext] || "text/plain") + ";charset=utf-8";
  }

  /* UTF-8, falling back to Big5 (HK/TC files), then lossy UTF-8. */
  function decodeText(buf) {
    var bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (e) {
      try {
        return new TextDecoder("big5").decode(bytes);
      } catch (e2) {
        return new TextDecoder("utf-8").decode(bytes);
      }
    }
  }

  function docxLines(xml) {
    var lines = [];
    var chunks = String(xml).split("</w:p>");
    for (var i = 0; i < chunks.length; i++) {
      var re = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
      var runs = [];
      var m;
      while ((m = re.exec(chunks[i])) !== null) runs.push(unxml(m[1]));
      var line = runs.join("").trim();
      if (line) lines.push(line);
    }
    return lines;
  }

  /* ---------- where the bytes of a link actually live ---------- */

  /* GitHub shows a file at /<owner>/<repo>/blob/<ref>/<path>, and that URL is
     an HTML page.  The bytes are at raw.githubusercontent.com. */
  function rawUrl(href) {
    var u;
    try {
      u = new URL(href);
    } catch (e) {
      return href;
    }
    if (u.hostname !== "github.com") return href;
    var m = /^\/([^/]+)\/([^/]+)\/(?:blob|raw)\/([^/]+)\/(.+)$/.exec(u.pathname);
    if (!m) return href;
    // github.com/<o>/<r>/raw/... is same-origin (the browser sends the session
    // cookie, so a private repository the reader can see also works) and it
    // redirects to the signed file URL.
    return "https://github.com/" + m[1] + "/" + m[2] + "/raw/" + m[3] + "/" + m[4];
  }

  /* ---------- markdown -> html ---------- */

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  /* inline spans: code first, then links and images, then emphasis.  Every
     protected span is swapped for a placeholder so emphasis cannot reach
     inside it. */
  function inlineMd(text) {
    var kept = [];
    var hold = function (html) {
      kept.push(html);
      return "\u0000" + (kept.length - 1) + "\u0000";
    };

    var s = escapeHtml(text);
    s = s.replace(/`([^`]+)`/g, function (m, code) {
      return hold("<code>" + code + "</code>");
    });
    s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)[^)]*\)/g, function (m, alt, src) {
      return hold('<img alt="' + alt + '" src="' + src + '">');
    });
    s = s.replace(/\[([^\]]+)\]\(([^)\s]+)[^)]*\)/g, function (m, label, href) {
      return hold('<a href="' + href + '" target="_blank" rel="noopener noreferrer">' + label + "</a>");
    });
    s = s.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, function (m, url) {
      return hold('<a href="' + url + '" target="_blank" rel="noopener noreferrer">' + url + "</a>");
    });

    s = s.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
    s = s.replace(/__([^_]+)__/g, "<strong>$1</strong>");
    s = s.replace(/~~([^~]+)~~/g, "<del>$1</del>");
    s = s.replace(/(^|[\s(])\*([^*\s][^*]*?)\*/g, "$1<em>$2</em>");
    s = s.replace(/(^|[\s(])_([^_\s][^_]*?)_/g, "$1<em>$2</em>");

    return s.replace(/\u0000(\d+)\u0000/g, function (m, i) {
      return kept[+i];
    });
  }

  function splitRow(line) {
    var t = line.trim().replace(/^\|/, "").replace(/\|$/, "");
    return t.split("|").map(function (c) {
      return c.trim();
    });
  }

  /* a small block parser: headings, rules, fences, quotes, lists, pipe tables
     and paragraphs.  Raw HTML is escaped, never passed through. */
  function markdown(src) {
    var lines = String(src).replace(/\r\n?/g, "\n").split("\n");
    var out = [];
    var para = [];
    var i = 0;

    function flush() {
      if (para.length) {
        out.push("<p>" + inlineMd(para.join(" ")) + "</p>");
        para = [];
      }
    }

    while (i < lines.length) {
      var line = lines[i];

      var fence = /^\s*(```+|~~~+)\s*([\w+#.-]*)\s*$/.exec(line);
      if (fence) {
        flush();
        var code = [];
        var close = new RegExp("^\\s*" + fence[1].slice(0, 3));
        i++;
        while (i < lines.length && !close.test(lines[i])) code.push(lines[i++]);
        i++;
        out.push(
          '<pre><code class="lang-' + (fence[2] || "text") + '">' + escapeHtml(code.join("\n")) + "</code></pre>"
        );
        continue;
      }

      if (!/\S/.test(line)) {
        flush();
        i++;
        continue;
      }

      var h = /^\s*(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
      if (h) {
        flush();
        var n = h[1].length;
        out.push("<h" + n + ">" + inlineMd(h[2]) + "</h" + n + ">");
        i++;
        continue;
      }

      if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) {
        flush();
        out.push("<hr>");
        i++;
        continue;
      }

      if (/^\s*>\s?/.test(line)) {
        flush();
        var quote = [];
        while (i < lines.length && /^\s*>\s?/.test(lines[i])) {
          quote.push(lines[i].replace(/^\s*>\s?/, ""));
          i++;
        }
        out.push("<blockquote>" + markdown(quote.join("\n")) + "</blockquote>");
        continue;
      }

      if (/\|/.test(line) && i + 1 < lines.length && /^\s*\|?[\s:-]*-[\s:|-]*\|?[\s:|-]*$/.test(lines[i + 1])) {
        flush();
        var head = splitRow(line);
        i += 2;
        var rows = [];
        while (i < lines.length && /\|/.test(lines[i]) && /\S/.test(lines[i])) rows.push(splitRow(lines[i++]));
        out.push(
          "<table><thead><tr>" +
            head.map(function (c) { return "<th>" + inlineMd(c) + "</th>"; }).join("") +
            "</tr></thead><tbody>" +
            rows
              .map(function (r) {
                return "<tr>" + r.map(function (c) { return "<td>" + inlineMd(c) + "</td>"; }).join("") + "</tr>";
              })
              .join("") +
            "</tbody></table>"
        );
        continue;
      }

      var item = /^\s*([-*+]|\d+[.)])\s+(.*)$/.exec(line);
      if (item) {
        flush();
        var ordered = /^\d/.test(item[1]);
        var items = [];
        while (i < lines.length) {
          var m = /^\s*([-*+]|\d+[.)])\s+(.*)$/.exec(lines[i]);
          if (!m) break;
          var text = m[2];
          var task = /^\[([ xX])\]\s+/.exec(text);
          if (task) text = (task[1] === " " ? "\u2610 " : "\u2611 ") + text.slice(task[0].length);
          i++;
          while (
            i < lines.length &&
            /^\s{2,}\S/.test(lines[i]) &&
            !/^\s*([-*+]|\d+[.)])\s+/.test(lines[i])
          ) {
            text += " " + lines[i].trim();
            i++;
          }
          items.push("<li>" + inlineMd(text) + "</li>");
        }
        out.push((ordered ? "<ol>" : "<ul>") + items.join("") + (ordered ? "</ol>" : "</ul>"));
        continue;
      }

      para.push(line.trim());
      i++;
    }

    flush();
    return out.join("\n");
  }

  var api = {
    detectKind: detectKind,
    fileName: fileName,
    formatSize: formatSize,
    mimeFor: mimeFor,
    decodeText: decodeText,
    docxLines: docxLines,
    markdown: markdown,
    rawUrl: rawUrl
  };

  root.HFP = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
