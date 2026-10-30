/* Word preview: reads a .docx (OOXML) package and renders its body as HTML.

   Scope: paragraphs and runs (bold/italic/underline/strike/size/colour/
   highlight/font), heading styles, numbered and bulleted lists, tables (grid,
   horizontal and vertical merges, shading, borders), pictures, page breaks,
   headers/footers, section columns and footnotes.

   Parsed with regular expressions and a small block scanner: a .docx body is a
   flat sequence of <w:p> and <w:tbl> blocks, and this keeps the renderer
   testable in node (no DOM). */
(function (root) {
  "use strict";

  var PX_PER_TWIP = 1 / 15; // 1 twip = 1/20 pt = 1/15 px at 96dpi
  var PAGE_BREAK = "\u0000page\u0000"; // split into sheets of paper by bodyHtml
  var MAX_IMAGE_BYTES = 4 * 1024 * 1024;

  /* ---------------------------------------------------------------- xml */

  function unxml(s) {
    return String(s)
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(+d); })
      .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
      .replace(/&amp;/g, "&");
  }

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function attr(tag, name) {
    var m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag);
    return m ? unxml(m[1]) : "";
  }

  function has(xml, re) {
    return re.test(xml);
  }

  /* the nth self-contained element <tag ...> ... </tag>, counting nesting */
  function slice(xml, start) {
    var tag = /^<([\w:]+)/.exec(xml.slice(start))[1];
    var open = new RegExp("<" + tag + "[ >]", "g");
    var close = new RegExp("</" + tag + ">", "g");
    var depth = 0;
    var i = start;
    while (i < xml.length) {
      open.lastIndex = i;
      close.lastIndex = i;
      var o = open.exec(xml);
      var c = close.exec(xml);
      if (!c) return xml.slice(start);
      if (o && o.index < c.index) {
        depth++;
        i = o.index + o[0].length;
      } else {
        depth--;
        i = c.index + c[0].length;
        if (depth === 0) return xml.slice(start, i);
      }
    }
    return xml.slice(start);
  }

  /* every top level <w:p> / <w:tbl> of a body (or of a table cell): a table
     nests inside a cell, so the scanner tracks both depths instead of counting
     one tag at a time */
  function blocks(xml) {
    var out = [];
    var re = /<(\/?)([\w:]+)([^>]*?)(\/?)>/g;
    var m;
    var depthP = 0;
    var depthT = 0;
    var start = 0;
    var kind = "";

    while ((m = re.exec(xml)) !== null) {
      var closing = m[1] === "/";
      var self = m[4] === "/";
      var name = m[2];
      if (self) continue;

      if (!closing && name === "w:p") {
        if (!depthP && !depthT) {
          start = m.index;
          kind = "p";
        }
        depthP++;
      } else if (!closing && name === "w:tbl") {
        if (!depthP && !depthT) {
          start = m.index;
          kind = "tbl";
        }
        depthT++;
      } else if (closing && name === "w:p") {
        depthP = Math.max(0, depthP - 1);
        if (!depthP && !depthT) out.push({ kind: kind || "p", xml: xml.slice(start, m.index + m[0].length) });
      } else if (closing && name === "w:tbl") {
        depthT = Math.max(0, depthT - 1);
        if (!depthP && !depthT) out.push({ kind: kind || "tbl", xml: xml.slice(start, m.index + m[0].length) });
      }
    }
    return out;
  }

  function find(xml, tag) {
    var m = new RegExp("<" + tag + "[^>]*>").exec(xml);
    return m ? m[0] : "";
  }

  function inner(xml, tag) {
    var m = new RegExp("<" + tag + "[^>]*>([\\s\\S]*?)</" + tag + ">").exec(xml);
    return m ? m[1] : "";
  }

  /* ------------------------------------------------------------ package */

  function relTarget(target) {
    if (!target) return "";
    var t = target.replace(/^\.\//, "");
    return t.charAt(0) === "/" ? t.slice(1) : "word/" + t;
  }

  function rels(xml) {
    var map = {};
    var re = /<Relationship\b[^>]*>/g;
    var m;
    while ((m = re.exec(xml)) !== null) {
      map[attr(m[0], "Id")] = {
        target: relTarget(attr(m[0], "Target")),
        external: attr(m[0], "TargetMode") === "External",
      };
    }
    return map;
  }

  /* ------------------------------------------------------------- styles */

  function parseStyles(xml) {
    var styles = {};
    // w:docDefaults holds the document-wide paragraph spacing and font; it is
    // the weakest layer, below every style and every run's own w:pPr
    var defs = inner(xml, "w:docDefaults");
    styles.__defaults = {
      run: runProps(inner(inner(defs, "w:rPrDefault"), "w:rPr")),
      para: paraProps(inner(inner(defs, "w:pPrDefault"), "w:pPr")),
    };
    var re = /<w:style\b[^>]*>/g;
    var m;
    while ((m = re.exec(xml)) !== null) {
      var id = attr(m[0], "w:styleId");
      if (!id) continue;
      var body = slice(xml, m.index);
      styles[id] = {
        type: attr(m[0], "w:type"),
        name: attr(find(body, "w:name"), "w:val"),
        basedOn: attr(find(body, "w:basedOn"), "w:val"),
        run: runProps(inner(body, "w:rPr")),
        para: paraProps(inner(body, "w:pPr")),
        tblBorders: inner(inner(body, "w:tblPr"), "w:tblBorders"),
      };
    }
    return styles;
  }

  var HEADINGS = /^(heading|titre|überschrift)\s*([1-9])$/i;

  function headingLevel(style, styles) {
    var seen = 0;
    while (style && seen++ < 8) {
      var s = styles[style];
      if (!s) break;
      if (s.name === "Title") return 1;
      var m = HEADINGS.exec(s.name || "");
      if (m) return +m[2];
      style = s.basedOn;
    }
    return 0;
  }

  function styleChain(style, styles) {
    var chain = [];
    var seen = 0;
    while (style && seen++ < 8) {
      var s = styles[style];
      if (!s) break;
      chain.push(s);
      style = s.basedOn;
    }
    return chain;
  }

  function runProps(xml) {
    var out = {};
    if (!xml) return out;
    if (/<w:b\b/.test(xml) && !/<w:b w:val="(0|false)"/.test(xml)) out.bold = 1;
    if (/<w:i\b/.test(xml) && !/<w:i w:val="(0|false)"/.test(xml)) out.italic = 1;
    var u = find(xml, "w:u");
    if (u && attr(u, "w:val") !== "none") out.underline = 1;
    if (/<w:strike\b/.test(xml) && !/<w:strike w:val="(0|false)"/.test(xml)) out.strike = 1;
    var sz = find(xml, "w:sz");
    if (sz) out.size = parseInt(attr(sz, "w:val"), 10) / 2; // half points
    var color = find(xml, "w:color");
    if (color && attr(color, "w:val") !== "auto") out.color = "#" + attr(color, "w:val");
    var hl = find(xml, "w:highlight");
    if (hl) out.highlight = attr(hl, "w:val");
    var fonts = find(xml, "w:rFonts");
    if (fonts) out.font = attr(fonts, "w:ascii") || attr(fonts, "w:hAnsi");
    var va = find(xml, "w:vertAlign");
    if (va) out.valign = attr(va, "w:val");
    var caps = find(xml, "w:caps");
    if (caps) out.caps = 1;
    return out;
  }

  function paraProps(xml) {
    var out = {};
    if (!xml) return out;
    var jc = find(xml, "w:jc");
    if (jc) out.align = attr(jc, "w:val");
    var ind = find(xml, "w:ind");
    if (ind) {
      out.left = parseInt(attr(ind, "w:left") || attr(ind, "w:start") || "0", 10);
      out.right = parseInt(attr(ind, "w:right") || attr(ind, "w:end") || "0", 10);
      out.first = parseInt(attr(ind, "w:firstLine") || "0", 10);
      out.hanging = parseInt(attr(ind, "w:hanging") || "0", 10);
    }
    var spacing = find(xml, "w:spacing");
    if (spacing) {
      out.before = parseInt(attr(spacing, "w:before") || "0", 10);
      out.after = parseInt(attr(spacing, "w:after") || "0", 10);
      out.line = parseInt(attr(spacing, "w:line") || "0", 10);
      out.lineRule = attr(spacing, "w:lineRule");
    }
    var numPr = inner(xml, "w:numPr"); // the content, not the opening tag
    if (numPr) {
      out.numId = parseInt(attr(find(numPr, "w:numId"), "w:val") || "0", 10);
      out.ilvl = parseInt(attr(find(numPr, "w:ilvl"), "w:val") || "0", 10);
    }
    var shd = find(xml, "w:shd");
    if (shd) out.shade = attr(shd, "w:fill");
    if (/<w:pageBreakBefore\b/.test(xml) && !/w:val="(0|false)"/.test(find(xml, "w:pageBreakBefore"))) {
      out.pageBreakBefore = 1;
    }
    return out;
  }

  /* ---------------------------------------------------------- numbering */

  function parseNumbering(xml) {
    var abstract = {};
    var re = /<w:abstractNum\b[^>]*>/g;
    var m;
    while ((m = re.exec(xml)) !== null) {
      var id = attr(m[0], "w:abstractNumId");
      var body = slice(xml, m.index);
      var levels = {};
      var lv = /<w:lvl\b[^>]*>/g;
      var l;
      while ((l = lv.exec(body)) !== null) {
        var lvl = attr(l[0], "w:ilvl");
        var chunk = slice(body, l.index);
        var fmt = attr(find(chunk, "w:numFmt"), "w:val") || "decimal";
        var ind = find(chunk, "w:ind");
        levels[lvl] = {
          format: fmt,
          text: attr(find(chunk, "w:lvlText"), "w:val") || "",
          start: parseInt(attr(find(chunk, "w:start"), "w:val") || "1", 10),
          // the level's own indent; a numbered heading has none and sits at the
          // page margin, unlike a bullet list which always carries w:ind
          left: parseInt(attr(ind, "w:left") || attr(ind, "w:start") || "0", 10),
          hanging: parseInt(attr(ind, "w:hanging") || "0", 10),
        };
      }
      abstract[id] = levels;
    }

    var nums = {};
    var nr = /<w:num\b[^>]*>/g;
    while ((m = nr.exec(xml)) !== null) {
      var numId = attr(m[0], "w:numId");
      var body2 = slice(xml, m.index);
      nums[numId] = attr(find(body2, "w:abstractNumId"), "w:val");
    }
    return { abstract: abstract, nums: nums };
  }

  /* --------------------------------------------------------------- runs */

  function runHtml(rPr, text, ctx) {
    if (!text) return "";
    var st = [];
    if (rPr.bold) st.push("font-weight:bold");
    if (rPr.italic) st.push("font-style:italic");
    if (rPr.strike) st.push("text-decoration:line-through");
    else if (rPr.underline) st.push("text-decoration:underline");
    if (rPr.size) st.push("font-size:" + rPr.size + "pt");
    if (rPr.color) st.push("color:" + rPr.color);
    if (rPr.highlight) st.push("background:" + highlight(rPr.highlight));
    if (rPr.font) st.push("font-family:" + esc(rPr.font));
    if (rPr.caps) text = text.toUpperCase();
    var out = esc(text).replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;");
    if (rPr.valign === "superscript" || rPr.valign === "subscript") {
      out = "<sup>" + out + "</sup>";
    }
    return st.length ? '<span style="' + st.join(";") + '">' + out + "</span>" : out;
  }

  function highlight(name) {
    var map = {
      yellow: "#fff59d", green: "#c8e6c9", cyan: "#b2ebf2", magenta: "#f8bbd0",
      blue: "#bbdefb", red: "#ffcdd2", darkBlue: "#9fa8da", darkRed: "#ef9a9a",
      darkGreen: "#a5d6a7", darkYellow: "#fff176", gray: "#e0e0e0", black: "#9e9e9e",
    };
    return map[name] || "#fff59d";
  }

  /* ------------------------------------------------------------ pictures */

  /* A blob URL keeps the HTML small; a FileReader data URL is the fallback for
     a context where createObjectURL is not usable. */
  function pictureUrl(bytes, type) {
    var blob = new Blob([bytes], { type: type });
    try {
      var url = URL.createObjectURL(blob);
      if (url) return Promise.resolve(url);
    } catch (e) {
      /* fall through */
    }
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(String(reader.result)); };
      reader.onerror = function () { reject(reader.error || new Error("read failed")); };
      reader.readAsDataURL(blob);
    });
  }

  function imageHtml(drawing, ctx) {
    // ctx.media holds a blob URL per picture, made in render()
    var blip = find(drawing, "a:blip");
    var rid = attr(blip, "r:embed") || attr(blip, "r:link");
    var rel = ctx.rels[rid];
    if (!rel || rel.external) return "";
    var data = ctx.media[rel.target];
    if (!data) return "";
    var ext = find(drawing, "wp:extent");
    var cx = parseInt(attr(ext, "cx") || "0", 10) / 9525; // EMU -> px
    var cy = parseInt(attr(ext, "cy") || "0", 10) / 9525;
    if (root.__docxPictures) root.__docxPictures.imgs++;
    // a w:drawing is inline in Word: it follows the text flow, so two pictures
    // in one paragraph sit side by side and a centred paragraph centres them.
    // The card's stylesheet makes .docx-img a block; an inline style wins.
    var style = "display:inline-block;vertical-align:bottom" +
      (cx ? ";width:" + Math.round(cx) + "px;height:" + Math.round(cy) + "px" : "");
    return '<img class="docx-img" src="' + data + '" alt="" style="' + style + '">';
  }

  /* Word builds a number from the level's w:lvlText: "%1." at level 1,
     "%1.%2" at level 2, so a heading under "2." reads "2.1" */
  function markerText(levels, lvl, counters, ilvl) {
    if (!lvl) return "";
    if (lvl.format === "bullet") return bullet(ilvl);
    if (lvl.format === "none") return "";
    var tpl = lvl.text || "";
    if (tpl.indexOf("%") >= 0) {
      return tpl.replace(/%(\d)/g, function (m, d) {
        var at = parseInt(d, 10) - 1;
        var lv = levels[String(at)];
        return formatNumber(counters[at] || 0, lv ? lv.format : "decimal");
      });
    }
    return formatNumber(counters[ilvl] || 1, lvl.format) + ".";
  }

  /* ---------------------------------------------------------- paragraphs */

  function paragraphHtml(p, ctx) {
    var pPr = inner(p, "w:pPr");
    var own = ctx.para(pPr);
    var styleId = attr(find(pPr, "w:pStyle"), "w:val");
    var defaults = ctx.styles.__defaults || { run: {}, para: {} };

    // inheritance, weakest first: document defaults, the style chain, and then
    // the paragraph's own properties
    var props = merge({}, defaults.para);
    var chain = styleChain(styleId, ctx.styles);
    for (var i = chain.length - 1; i >= 0; i--) props = merge(props, chain[i].para);
    props = merge(props, own);
    var baseRun = merge({}, defaults.run);
    for (var j = chain.length - 1; j >= 0; j--) baseRun = merge(baseRun, chain[j].run);

    var html = [];
    var re = /<w:(r|hyperlink|br|tab|fldSimple|drawing)\b[^>]*>/g;
    var m;
    // the run scan must not see paragraph properties: <w:tabs><w:tab/> lives in
    // w:pPr, and a tab stop there used to be taken for a run-level tab, which
    // swallowed the rest of the paragraph (a Word TOC entry lost its text)
    var body = p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "");
    while ((m = re.exec(body)) !== null) {
      var chunk = slice(body, m.index);
      if (m[1] === "r") {
        var rPr = merge(baseRun, ctx.run(inner(chunk, "w:rPr")));
        var parts = [];
        var tr = /<w:(t|tab|br|sym|drawing)(?=[ >/])[^>]*\/?>/g;
        var tm;
        while ((tm = tr.exec(chunk)) !== null) {
          if (tm[1] === "t") {
            parts.push(runHtml(rPr, unxml(inner(chunk.slice(tm.index), "w:t")), ctx));
          } else if (tm[1] === "tab") {
            parts.push(runHtml(rPr, "\t", ctx));
          } else if (tm[1] === "br") {
            parts.push(
              attr(tm[0], "w:type") === "page" ? PAGE_BREAK : "<br>"
            );
          } else if (tm[1] === "drawing") {
            parts.push(imageHtml(chunk.slice(tm.index), ctx));
          }
          tr.lastIndex = tm.index + tm[0].length;
        }
        html.push(parts.join(""));
      } else if (m[1] === "hyperlink") {
        var text = "";
        var hr = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
        var hm;
        while ((hm = hr.exec(chunk)) !== null) text += unxml(hm[1]);
        var rid2 = attr(m[0], "r:id");
        var target = rid2 && ctx.rels[rid2] ? ctx.rels[rid2].target : "";
        html.push(
          target
            ? '<a href="' + esc(target) + '" target="_blank" rel="noopener noreferrer">' + esc(text) + "</a>"
            : esc(text)
        );
      } else if (m[1] === "br") {
        html.push(
          attr(m[0], "w:type") === "page" ? PAGE_BREAK : "<br>"
        );
      } else if (m[1] === "tab") {
        html.push("&nbsp;&nbsp;&nbsp;&nbsp;");
      } else if (m[1] === "fldSimple") {
        var ft = "";
        var fr = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
        var fm;
        while ((fm = fr.exec(chunk)) !== null) ft += unxml(fm[1]);
        html.push(esc(ft));
      } else if (m[1] === "drawing") {
        html.push(imageHtml(chunk, ctx));
      }
      re.lastIndex = m.index + chunk.length;
    }

    var content = html.join("");
    if (!content.trim() && !props.pageBreakBefore) return "";

    var tag = "p";
    var level = headingLevel(styleId, ctx.styles);
    if (level) tag = "h" + Math.min(level, 6);

    var style = [];
    if (props.align) style.push("text-align:" + align(props.align));
    if (props.left) style.push("margin-left:" + Math.round(props.left * PX_PER_TWIP) + "px");
    if (props.right) style.push("margin-right:" + Math.round(props.right * PX_PER_TWIP) + "px");
    if (props.first) style.push("text-indent:" + Math.round(props.first * PX_PER_TWIP) + "px");
    if (props.hanging) style.push("text-indent:-" + Math.round(props.hanging * PX_PER_TWIP) + "px");
    if (props.before) style.push("margin-top:" + Math.round(props.before * PX_PER_TWIP) + "px");
    if (props.after) style.push("margin-bottom:" + Math.round(props.after * PX_PER_TWIP) + "px");
    if (props.line && props.lineRule !== "auto") {
      style.push("line-height:" + (props.line / 240).toFixed(2));
    }
    if (props.shade && props.shade !== "auto") style.push("background:#" + props.shade);

    var list =
      props.numId && ctx.numbering.nums[props.numId] !== undefined
        ? ctx.numbering.abstract[ctx.numbering.nums[props.numId]]
        : null;
    var marker = "";
    if (list) {
      var ilvl = props.ilvl || 0;
      var lvl = list[String(ilvl)] || list["0"];
      // one counter per level under this numId; a deeper level restarts
      var counters = ctx.counters[props.numId] || (ctx.counters[props.numId] = []);
      counters[ilvl] = (counters[ilvl] || (lvl ? lvl.start : 1) - 1) + 1;
      for (var deeper = ilvl + 1; deeper < 9; deeper++) counters[deeper] = 0;
      marker = markerText(list, lvl, counters, ilvl);
      // Word puts the label at the level's indent and hangs it into the margin
      if (lvl && lvl.left) style.push("padding-left:" + Math.round(lvl.left * PX_PER_TWIP) + "px");
      if (lvl && lvl.hanging) style.push("text-indent:-" + Math.round(lvl.hanging * PX_PER_TWIP) + "px");
    }

    var out = "<" + tag + (style.length ? ' style="' + style.join(";") + '"' : "") + ">";
    // a space after the label, the way Word's tab separates number and text
    if (marker) out += '<span class="docx-marker">' + marker + "</span> ";
    out += content + "</" + tag + ">";
    if (props.pageBreakBefore) out = PAGE_BREAK + out;
    return out;
  }

  function merge(a, b) {
    if (!b) return a;
    var out = {};
    Object.keys(a || {}).forEach(function (k) { out[k] = a[k]; });
    Object.keys(b).forEach(function (k) {
      if (b[k] !== undefined && b[k] !== null && b[k] !== "") out[k] = b[k];
    });
    return out;
  }

  function align(v) {
    return v === "both" ? "justify" : v === "center" ? "center" : v === "right" ? "right" : "left";
  }

  function bullet(level) {
    return ["•", "◦", "▪", "•", "◦"][Math.min(level, 4)];
  }

  function formatNumber(n, fmt) {
    if (fmt === "lowerLetter") return String.fromCharCode(96 + ((n - 1) % 26) + 1);
    if (fmt === "upperLetter") return String.fromCharCode(64 + ((n - 1) % 26) + 1);
    if (fmt === "lowerRoman" || fmt === "upperRoman") {
      var roman = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"][n] || String(n);
      return fmt === "upperRoman" ? roman.toUpperCase() : roman;
    }
    return String(n);
  }

  /* -------------------------------------------------------------- tables */

  function tableHtml(tbl, ctx) {
    var tblPr = inner(tbl, "w:tblPr");
    // the borders usually live in the table style (TableGrid and friends), and
    // the table's own tblBorders override them
    var styleId = attr(find(tblPr, "w:tblStyle"), "w:val");
    var styled = styleId && ctx.styles[styleId] ? ctx.styles[styleId].tblBorders : "";
    var own = inner(tblPr, "w:tblBorders");
    var borders = borderMap(own, styled);
    var width = parseInt(attr(find(tblPr, "w:tblW"), "w:w") || "0", 10);

    var grid = [];
    var gr = /<w:gridCol[^>]*>/g;
    var gm;
    var gridXml = inner(tbl, "w:tblGrid");
    while ((gm = gr.exec(gridXml)) !== null) {
      grid.push(parseInt(attr(gm[0], "w:w") || "0", 10));
    }

    var rows = [];
    var trRe = /<w:tr\b[^>]*>/g;
    var trm;
    while ((trm = trRe.exec(tbl)) !== null) {
      var tr = slice(tbl, trm.index);
      var cells = [];
      var tcRe = /<w:tc\b[^>]*>/g;
      var tcm;
      while ((tcm = tcRe.exec(tr)) !== null) {
        var tc = slice(tr, tcm.index);
        var tcPr = inner(tc, "w:tcPr");
        var span = parseInt(attr(find(tcPr, "w:gridSpan"), "w:val") || "1", 10);
        var vMergeTag = find(tcPr, "w:vMerge");
        var vMerge = vMergeTag ? (attr(vMergeTag, "w:val") || "continue") : "";
        var shd = find(tcPr, "w:shd");
        var fill = shd ? attr(shd, "w:fill") : "";
        var cellStyle = [];
        if (fill && fill !== "auto") cellStyle.push("background:#" + fill);
        var cellBorders = inner(tcPr, "w:tcBorders");

        var content = blocks(tc)
          .map(function (b) { return b.kind === "tbl" ? tableHtml(b.xml, ctx) : paragraphHtml(b.xml, ctx); })
          .join("");

        cells.push({
          html: content,
          span: span,
          vMerge: vMerge,
          style: cellStyle.join(";"),
          borders: cellBorders,
        });
        tcRe.lastIndex = tcm.index + tc.length;
      }
      rows.push(cells);
      trRe.lastIndex = trm.index + tr.length;
    }

    // a vertical merge continues the cell above: give the top cell a rowspan
    for (var c = 0; c < (rows[0] || []).length; c++) {
      for (var r = 0; r < rows.length; r++) {
        var cell = rows[r][c];
        if (!cell || cell.vMerge !== "restart") continue;
        var span = 1;
        for (var k = r + 1; k < rows.length; k++) {
          if (rows[k][c] && rows[k][c].vMerge === "continue") {
            span++;
            rows[k][c].drop = true;
          } else break;
        }
        cell.rowspan = span;
      }
    }

    var html = [];
    var colgroup = grid.length
      ? "<colgroup>" +
        grid.map(function (w) { return '<col style="width:' + Math.round(w * PX_PER_TWIP) + 'px">'; }).join("") +
        "</colgroup>"
      : "";
    var tableStyle = [];
    if (width) tableStyle.push("width:" + Math.round(width * PX_PER_TWIP) + "px");
    tableStyle.push("border-collapse:collapse");
    OUTER_SIDES.forEach(function (side) {
      if (borders[side]) tableStyle.push("border-" + side + ":" + borders[side]);
    });
    html.push('<table class="docx-table" style="' + tableStyle.join(";") + '">' + colgroup);
    // only rows that still have a cell count for the frame's top/bottom edges
    var emitRows = rows
      .map(function (row) { return row.filter(function (cell) { return cell && !cell.drop; }); })
      .filter(function (cells) { return cells.length; });
    emitRows.forEach(function (emit, ri) {
      html.push("<tr>");
      emit.forEach(function (cell, ci) {
        var css = cell.style ? [cell.style] : [];
        // in the collapsing border model the table element's own borders are
        // ignored, so the frame has to sit on the outer cells: top on the first
        // row, left on the first column, bottom/right on the last ones
        var frame = {
          top: ri === 0 ? borders.top : "",
          left: ci === 0 ? borders.left : "",
          bottom: ri === emitRows.length - 1 ? borders.bottom : borders.insideH,
          right: ci === emit.length - 1 ? borders.right : borders.insideV,
        };
        OUTER_SIDES.forEach(function (side) {
          // a cell's own w:tcBorders wins (and can cancel with nil/none)
          var v = sideCss(cell.borders, side);
          if (v === null) v = frame[side];
          if (v) css.push("border-" + side + ":" + v);
        });
        html.push(
          "<td" +
            (cell.span > 1 ? ' colspan="' + cell.span + '"' : "") +
            (cell.rowspan > 1 ? ' rowspan="' + cell.rowspan + '"' : "") +
            (css.length ? ' style="' + css.join(";") + '"' : "") +
            ">" + cell.html + "</td>"
        );
      });
      html.push("</tr>");
    });
    html.push("</table>");
    return html.join("");
  }

  /* Word borders are tri-state per side: absent inherits, nil/none cancels,
     anything else draws. A table's own w:tblBorders win over its table style. */
  var OUTER_SIDES = ["top", "left", "bottom", "right"];
  var BORDER_SIDES = OUTER_SIDES.concat(["insideH", "insideV"]);

  function sideCss(xml, side) {
    if (!xml) return null;
    var tag = find(xml, "w:" + side);
    if (!tag) return null;
    var val = attr(tag, "w:val");
    if (!val || val === "none" || val === "nil") return "";
    var sz = parseInt(attr(tag, "w:sz") || "4", 10) / 8;
    var color = attr(tag, "w:color");
    var line =
      val === "dotted" ? "dotted" : val === "dashed" ? "dashed" : val === "double" ? "double" : "solid";
    return Math.max(1, Math.round(sz)) + "px " + line + " " + (color && color !== "auto" ? "#" + color : "#808080");
  }

  function borderMap(own, styled) {
    var out = {};
    BORDER_SIDES.forEach(function (side) {
      var v = sideCss(own, side);
      out[side] = v === null ? sideCss(styled, side) || "" : v;
    });
    return out;
  }

  /* ------------------------------------------------------------ sections */

  /* one section's header, footer and column count */
  function sectionParts(sect, ctx) {
    var out = { header: "", footer: "", cols: 1 };
    if (!sect) return out;
    out.cols = parseInt(attr(find(sect, "w:cols"), "w:num") || "1", 10) || 1;
    ["headerReference", "footerReference"].forEach(function (kind) {
      var re = new RegExp("<w:" + kind + "[^>]*>", "g");
      var m;
      while ((m = re.exec(sect)) !== null) {
        var type = attr(m[0], "w:type") || "default";
        if (type === "first" || type === "even") continue;
        var rel = ctx.rels[attr(m[0], "r:id")];
        if (!rel) continue;
        var part = ctx.parts[rel.target];
        if (!part) continue;
        var html = blocks(part)
          .map(function (b) { return b.kind === "tbl" ? tableHtml(b.xml, ctx) : paragraphHtml(b.xml, ctx); })
          .join("");
        if (!html.trim()) continue;
        var which = kind === "headerReference" ? "header" : "footer";
        out[which] += '<div class="docx-' + which + '">' + html + "</div>";
      }
    });
    return out;
  }

  function wrapSection(html, sect, ctx) {
    var parts = sectionParts(sect, ctx);
    var cols = Math.min(parts.cols, 3);
    return { parts: parts, cols: cols, pieces: html.split(PAGE_BREAK) };
  }

  /* ----------------------------------------------------------- the body */

  /* the sheet the document is set on: width and margins come from the final
     w:sectPr, so the preview's text column breaks where Word's does */
  function pageCss(sect) {
    var pgSz = find(sect, "w:pgSz");
    var mar = find(sect, "w:pgMar");
    var twip = function (tag, name, fallback) {
      var v = attr(tag, name);
      return v ? Math.round(parseInt(v, 10) * PX_PER_TWIP) : fallback;
    };
    var w = pgSz ? Math.round(parseInt(attr(pgSz, "w:w") || "11906", 10) * PX_PER_TWIP) : 794;
    return (
      "width:" + w + "px;max-width:100%;box-sizing:border-box;" +
      "padding:" + twip(mar, "w:top", 18) + "px " + twip(mar, "w:right", 26) + "px " +
      twip(mar, "w:bottom", 18) + "px " + twip(mar, "w:left", 26) + "px"
    );
  }

  function bodyHtml(doc, ctx) {
    var body = inner(doc, "w:body");
    // the last sectPr belongs to the body; the ones before it end a section
    var lastAt = body.lastIndexOf("<w:sectPr");
    var finalSect = lastAt >= 0 ? slice(body, lastAt) : "";
    var pageStyle = pageCss(finalSect);
    var pages = [];
    var current = "";

    function flush() {
      if (current) pages.push(current);
      current = "";
    }

    function addSection(html, sect) {
      var w = wrapSection(html, sect, ctx);
      // w:type continuous means "carry on on this page"; anything else (or the
      // default nextPage) starts a new one
      var continuous = attr(find(sect, "w:type"), "w:val") === "continuous";

      w.pieces.forEach(function (piece, i) {
        var block =
          w.cols > 1
            ? '<div class="docx-cols" style="column-count:' + w.cols + '">' + piece + "</div>"
            : piece;
        if (i > 0) flush();                 // a page break starts a sheet
        else if (!continuous) flush();      // a next page section starts one too
        if (!current) current = w.parts.header;
        current += block;
        if (i === w.pieces.length - 1 && !continuous) current += w.parts.footer;
      });
    }

    var current_ = [];
    blocks(body).forEach(function (b) {
      if (b.kind === "tbl") {
        current_.push(tableHtml(b.xml, ctx));
        return;
      }
      current_.push(paragraphHtml(b.xml, ctx));
      var sect = inner(inner(b.xml, "w:pPr"), "w:sectPr");
      if (sect) {
        addSection(current_.join(""), "<w:sectPr>" + sect + "</w:sectPr>");
        current_ = [];
      }
    });
    if (current_.length) addSection(current_.join(""), finalSect);
    flush();

    if (!pages.length) return "";
    return pages
      .map(function (page, i) {
        return (
          '<div class="docx-page" style="' + pageStyle + '">' +
          (pages.length > 1
            ? '<div class="docx-page-no">' + (i + 1) + " / " + pages.length + "</div>"
            : "") +
          page +
          "</div>"
        );
      })
      .join("");
  }

  function footnotesHtml(doc, ctx) {
    var used = [];
    var re = /<w:footnoteReference[^>]*w:id="(\d+)"/g;
    var m;
    while ((m = re.exec(doc)) !== null) used.push(m[1]);
    if (!used.length) return "";

    var out = ['<div class="docx-footnotes">'];
    used.forEach(function (id2, i) {
      var note = ctx.footnotes[id2];
      if (!note) return;
      out.push(
        '<div class="docx-footnote"><sup>' + (i + 1) + "</sup>" + note + "</div>"
      );
    });
    out.push("</div>");
    return out.join("");
  }

  /* ------------------------------------------------------------- render */

  async function render(buf) {
    var zip = await root.HFPZip.load(buf);
    var read = function (name) {
      var f = zip.file(name);
      return f ? f.async("string") : Promise.resolve("");
    };

    var [doc, stylesXml, numberingXml, docRels] = await Promise.all([
      read("word/document.xml"),
      read("word/styles.xml"),
      read("word/numbering.xml"),
      read("word/_rels/document.xml.rels"),
    ]);
    if (!doc) throw new Error("no word/document.xml");

    var relMap = rels(docRels);

    // every part the body can point at: media for pictures, headers/footers
    var parts = {};
    var media = {};
    var urls = [];
    var extra = Object.keys(relMap).map(function (id) {
      var rel = relMap[id];
      if (rel.external) return Promise.resolve();
      var target = rel.target;
      if (!zip.file(target)) return Promise.resolve();
      if (/\.(png|jpe?g|gif|webp|bmp)$/i.test(target)) {
        // a blob URL, not a data URI: the bytes are already in memory and this
        // keeps the page's HTML small (a 3 MB report carries 22 pictures)
        return zip
          .file(target)
          .async("uint8array")
          .then(function (bytes) {
            if (!bytes || bytes.length > MAX_IMAGE_BYTES) return;
            var type = /\.png$/i.test(target)
              ? "image/png"
              : /\.gif$/i.test(target)
                ? "image/gif"
                : /\.webp$/i.test(target)
                  ? "image/webp"
                  : /\.bmp$/i.test(target)
                    ? "image/bmp"
                    : "image/jpeg";
            return pictureUrl(bytes, type).then(function (url) {
              urls.push(url);
              media[target] = url;
            });
          })
          .catch(function () {});
      }
      return read(target).then(function (text) {
        if (text) parts[target] = inner(text, "w:body") || text;
      });
    });

    var ctx = {
      styles: parseStyles(stylesXml),
      numbering: parseNumbering(numberingXml),
      rels: relMap,
      parts: parts,
      media: media,
      counters: {},
      run: runProps,
      para: paraProps,
    };

    var footnotes = {};
    var fnXml = await read("word/footnotes.xml");
    if (fnXml) {
      var fre = /<w:footnote\b[^>]*>/g;
      var fm;
      while ((fm = fre.exec(fnXml)) !== null) {
        var id3 = attr(fm[0], "w:id");
        var piece = slice(fnXml, fm.index);
        footnotes[id3] = blocks(piece)
          .map(function (b) { return paragraphHtml(b.xml, ctx); })
          .join("");
      }
    }
    ctx.footnotes = footnotes;

    await Promise.all(extra);
    root.__docxMediaUrls = urls;

    root.__docxPictures = { media: Object.keys(media).length, imgs: 0 };
    return (
      '<div class="docx">' +
      bodyHtml(doc, ctx) +
      footnotesHtml(doc, ctx) +
      "</div>"
    );
  }

  root.HFPdocx = {
    render: render,
    _blocks: blocks,
    _props: runProps,
    _pProps: paraProps,
    _para: paragraphHtml,
    _table: tableHtml,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
