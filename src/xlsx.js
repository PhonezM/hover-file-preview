/* Spreadsheet preview: reads an .xlsx (OOXML) package and renders the values of
   every sheet as a plain table.  Values only - cell styles are not applied, but
   a cell whose number format is a date is shown as a date.

   Parsed with regular expressions: OOXML's worksheet parts are a flat list of
   <row>/<c> elements, and this keeps the reader testable in node (no DOM). */
(function (root) {
  "use strict";

  var MAX_COLS = 200;

  function unxml(s) {
    return String(s)
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (m, d) {
        return String.fromCharCode(+d);
      })
      .replace(/&#x([0-9a-f]+);/gi, function (m, h) {
        return String.fromCharCode(parseInt(h, 16));
      })
      .replace(/&amp;/g, "&");
  }

  function attr(tag, name) {
    var m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag);
    return m ? unxml(m[1]) : "";
  }

  function allText(xml, tag) {
    var out = [];
    var re = new RegExp("<" + tag + "[^>]*>([\\s\\S]*?)</" + tag + ">", "g");
    var m;
    while ((m = re.exec(xml)) !== null) out.push(m[1]);
    return out;
  }

  /* A1 -> {row: 0, col: 0} */
  function refToCell(ref) {
    var m = /^([A-Z]+)(\d+)$/.exec(ref || "");
    if (!m) return null;
    var col = 0;
    for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  }

  function colName(index) {
    var s = "";
    for (var n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
      s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    }
    return s;
  }

  /* ------------------------------------------------------- number formats */

  /* A style's numFmtId names a builtin code or one of the workbook's custom
     <numFmt> formatCode strings; a numeric cell is shown through that code the
     way Excel (and SheetJS's SSF) does it. */

  var GENERAL = "General";

  var BUILTIN = {
    0: GENERAL, 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
  };

  var MONTH_LONG = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
    "Sep", "Oct", "Nov", "Dec"];
  var DAY_LONG = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var DAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

  /* one format code per cellXf index, "" when the cell has no style */
  function cellFormats(stylesXml) {
    var custom = {};
    var re = /<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
    var m;
    while ((m = re.exec(stylesXml)) !== null) custom[m[1]] = unxml(m[2]);

    var out = [];
    var cellXfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    if (!cellXfs) return out;
    var xfRe = /<xf\b[^>]*>/g;
    while ((m = xfRe.exec(cellXfs[1])) !== null) {
      var id = attr(m[0], "numFmtId") || "0";
      out.push(custom[id] || BUILTIN[id] || "");
    }
    return out;
  }

  /* split on ';' outside quotes and escapes: positive;negative;zero;text */
  function sections(code) {
    var out = [];
    var cur = "";
    var quoted = false;
    for (var i = 0; i < code.length; i++) {
      var c = code[i];
      if (c === '"') quoted = !quoted;
      if (!quoted && c === "\\") {
        cur += c + (code[i + 1] || "");
        i++;
        continue;
      }
      if (!quoted && c === ";") {
        out.push(cur);
        cur = "";
        continue;
      }
      cur += c;
    }
    out.push(cur);
    return out;
  }

  function pad(n) {
    return n < 10 ? "0" + n : "" + n;
  }

  /* the 1900 date system (25569 = 1970-01-01), read in UTC so a preview does
     not shift with the machine's timezone */
  function dateParts(serial) {
    if (!isFinite(serial) || serial < 0) return null;
    var whole = Math.floor(serial);
    if (whole < 1) {
      // Excel's day zero: 1900-01-00, a Saturday (serials 0 .. 1)
      var secs0 = Math.round((serial - whole) * 86400);
      return {
        y: 1900, mo: 0, day: 0, wd: 6,
        h: Math.floor(secs0 / 3600), mi: Math.floor(secs0 / 60) % 60,
        s: secs0 % 60, ms: 0,
      };
    }
    // Excel counts a 1900-02-29 that never existed, so serials under 60 read
    // one day further along than the plain 25569 epoch suggests
    var days = whole < 60 ? serial + 1 : serial;
    var d = new Date(Math.round((days - 25569) * 86400000));
    if (isNaN(d.getTime())) return null;
    return {
      y: d.getUTCFullYear(), mo: d.getUTCMonth(), day: d.getUTCDate(), wd: d.getUTCDay(),
      h: d.getUTCHours(), mi: d.getUTCMinutes(),
      s: d.getUTCSeconds(), ms: d.getUTCMilliseconds(),
    };
  }

  function isDateCode(code) {
    var bare = code.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
    return /[ymdhs]/i.test(bare);
  }

  /* "12345.678", "1.23457E+12" - the way Excel's General column reads */
  function generalText(v) {
    if (!isFinite(v)) return String(v);
    if (v === 0) return "0";
    var a = Math.abs(v);
    if (a >= 1e11 || a < 1e-10) {
      var m = v.toExponential(5);
      var parts = m.split("e");
      var mant = parts[0].replace(/\.?0+$/, "");
      var exp = parts[1];
      if (/^[+-]\d$/.test(exp)) exp = exp.slice(0, -1) + "0" + exp.slice(-1);
      return mant + "E" + exp;
    }
    var s = String(v);
    return s;
  }

  /* tokens of a date/time code: literals and field runs */
  function dateTokens(code) {
    var toks = [];
    var i = 0;
    while (i < code.length) {
      var c = code[i];
      if (c === '"') {
        var j = code.indexOf('"', i + 1);
        if (j < 0) j = code.length;
        toks.push({ lit: code.slice(i + 1, j) });
        i = j + 1;
        continue;
      }
      if (c === "\\") {
        toks.push({ lit: code[i + 1] || "" });
        i += 2;
        continue;
      }
      if (c === "_") {
        toks.push({ lit: " " });
        i += 2;
        continue;
      }
      if (c === "*") {
        i += 2;
        continue;
      }
      if (c === "[") {
        var k = code.indexOf("]", i);
        var inner = code.slice(i + 1, k < 0 ? code.length : k);
        if (/^[hms]+$/i.test(inner)) toks.push({ field: inner.toLowerCase() }); // elapsed
        i = k < 0 ? code.length : k + 1;
        continue;
      }
      var rest = code.slice(i);
      var m =
        /^(AM\/PM|am\/pm|A\/P|a\/p)/i.exec(rest) ||
        /^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i.exec(rest) ||
        /^(\.0+)/.exec(rest);
      if (m) {
        toks.push({ field: /^am\/pm|^a\/p/i.test(m[0]) ? "ampm" : m[0].toLowerCase() });
        i += m[0].length;
        continue;
      }
      toks.push({ lit: c });
      i++;
    }
    return toks;
  }

  function formatDate(code, serial) {
    if (serial < 0) return "";
    var d = dateParts(serial);
    if (!d) return String(serial);
    var toks = dateTokens(code);
    var ampm = toks.some(function (t) { return t.field === "ampm"; });
    var h24 = d.h;
    var h = ampm ? (h24 % 12 || 12) : h24;
    var secs = d.s + d.ms / 1000;
    var out = "";
    var prevField = "";
    var nextField = function (at) {
      for (var i = at + 1; i < toks.length; i++) if (toks[i].field) return toks[i].field;
      return "";
    };

    for (var i = 0; i < toks.length; i++) {
      var f = toks[i].field;
      if (!f) {
        out += toks[i].lit;
        continue;
      }
      if (f === "ampm") {
        out += h24 < 12 ? "AM" : "PM";
      } else if (f[0] === "y") {
        out += f.length >= 4 ? d.y : pad(d.y % 100);
      } else if (f[0] === "m") {
        var minute = prevField === "h" || prevField === "hh" || nextField(i)[0] === "s";
        if (minute) out += f.length >= 2 ? pad(d.mi) : d.mi;
        else if (f === "mmmm") out += MONTH_LONG[d.mo];
        else if (f === "mmm") out += MONTH_SHORT[d.mo];
        else if (f === "mm") out += pad(d.mo + 1);
        else out += d.mo + 1;
      } else if (f[0] === "d") {
        if (f === "dddd") out += DAY_LONG[d.wd];
        else if (f === "ddd") out += DAY_SHORT[d.wd];
        else if (f.length >= 2) out += pad(d.day);
        else out += d.day;
      } else if (f[0] === "h") {
        out += f.length >= 2 ? pad(h) : h;
      } else if (f[0] === "s") {
        out += f.length >= 2 ? pad(d.s) : d.s;
      } else if (f[0] === ".") {
        // fractional seconds, e.g. ss.00
        var dec = f.length - 1;
        var frac = secs - Math.floor(secs);
        var digits = frac.toFixed(dec).slice(2);
        out += "." + digits;
      }
      prevField = f;
    }
    return out;
  }

  /* best proper fraction for a magnitude below 1, denominator under 10^digits */
  function bestFraction(x, digits) {
    var limit = Math.pow(10, digits);
    var P_2 = 0, P_1 = 1, Q_2 = 1, Q_1 = 0, P = 0, Q = 0, B = x;
    for (var k = 0; k < 40; k++) {
      var A = Math.floor(B);
      P = A * P_1 + P_2;
      Q = A * Q_1 + Q_2;
      if (Q >= limit) {
        Q = Q_1;
        P = P_1;
        break;
      }
      P_2 = P_1;
      P_1 = P;
      Q_2 = Q_1;
      Q_1 = Q;
      if (B - A < 1e-9) break;
      B = 1 / (B - A);
    }
    return Q ? [P, Q] : [0, 1];
  }

  function group(text) {
    var parts = text.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return parts.join(".");
  }

  /* one section of a numeric format applied to a positive magnitude */
  function formatNumberSection(sec, value) {
    var out = "";
    var num = "";
    var literals = { pre: "", suf: "" };
    var seenNumber = false;
    // '%' scales by 100 wherever it sits in the section, so count it up front
    var pct = (sec.replace(/"[^"]*"/g, "").match(/%/g) || []).length;
    var sign = false;
    var cur = "";
    var i = 0;

    function flushLit() {
      if (!cur) return;
      if (seenNumber) literals.suf += cur;
      else literals.pre += cur;
      cur = "";
    }

    while (i < sec.length) {
      var c = sec[i];
      if (c === '"') {
        var j = sec.indexOf('"', i + 1);
        if (j < 0) j = sec.length;
        cur += sec.slice(i + 1, j);
        i = j + 1;
        continue;
      }
      if (c === "\\") {
        cur += sec[i + 1] || "";
        i += 2;
        continue;
      }
      if (c === "_") {
        cur += " ";
        i += 2;
        continue;
      }
      if (c === "*") {
        i += 2;
        continue;
      }
      if (c === "[") {
        var k = sec.indexOf("]", i);
        var inner = sec.slice(i + 1, k < 0 ? sec.length : k);
        var money = /^\$([^-]*)-/.exec(inner);
        if (money) cur += money[1];
        i = k < 0 ? sec.length : k + 1;
        continue;
      }
      if (c === "%") {
        cur += "%";
        i++;
        continue;
      }
      if (/[0#?]/.test(c) || ((c === "." || c === ",") && /[0#?]?/.test(sec[i + 1] || ""))) {
        flushLit();
        seenNumber = true;
        num = "";
        var sci = false;
        while (i < sec.length) {
          var d = sec[i];
          if (/[0#?.,]/.test(d)) {
            num += d;
            i++;
            continue;
          }
          if ((d === "E" || d === "e") && num && /[+-]/.test(sec[i + 1] || "")) {
            num += "E" + sec[i + 1];
            i += 2;
            sci = true;
            while (i < sec.length && /[0#]/.test(sec[i])) {
              num += sec[i];
              i++;
            }
            break;
          }
          break;
        }
        var part = numberFromPattern(num, value, pct, sci);
        if (part.charAt(0) === "-") {
          sign = true;
          part = part.slice(1);
        }
        out += part;
        continue;
      }
      cur += c;
      i++;
    }
    flushLit();
    return (sign ? "-" : "") + literals.pre + out + literals.suf;
  }

  function numberFromPattern(num, value, pct, sci) {
    var n = value;
    for (var p = 0; p < pct; p++) n *= 100;

    // a trailing comma scales by 1000 ("0.0," shows 12345678 as 12345.7)
    var scale = 0;
    while (num.slice(-1) === ",") {
      num = num.slice(0, -1);
      scale++;
    }
    if (scale) n /= Math.pow(1000, scale);

    var intPart = num;
    var fracPart = "";
    var expPart = "";
    var e = num.search(/E/i);
    if (e >= 0) {
      expPart = num.slice(e + 1);
      intPart = num.slice(0, e);
    }

    var dot = intPart.indexOf(".");
    if (dot >= 0) {
      fracPart = intPart.slice(dot + 1);
      intPart = intPart.slice(0, dot);
    }

    var minDec = (fracPart.match(/0/g) || []).length;
    var maxDec = (fracPart.match(/[0#]/g) || []).length;

    if (expPart) {
      var sciNeg = n < 0;
      var mant = Math.abs(n);
      var exp = mant === 0 ? 0 : Math.floor(Math.log10(mant));
      // a pattern with several integer digits (##0) uses engineering steps of 3
      var step = (intPart.match(/[0#]/g) || []).length > 1 ? 3 : 1;
      exp = Math.floor(exp / step) * step;
      mant = Math.abs(n) / Math.pow(10, exp);
      var sign = exp < 0 ? "-" : "+";
      var expDigits = (expPart.match(/0/g) || []).length || 1;
      var expText = String(Math.abs(exp));
      while (expText.length < expDigits) expText = "0" + expText;
      var mantText = maxDec ? mant.toFixed(maxDec) : String(Math.round(mant));
      return (sciNeg && mant !== 0 ? "-" : "") + mantText + "E" + sign + expText;
    }

    // Excel rounds .5 away from zero on the magnitude
    var mag = Math.abs(n);
    var text = maxDec ? mag.toFixed(maxDec) : String(Math.round(mag));
    var neg = n < 0 && Number(text) !== 0;
    if (maxDec > minDec && text.indexOf(".") >= 0) {
      text = text.replace(/0+$/, "");
    }
    var hasDot = fracPart.length > 0;
    if (text.slice(-1) === "." && !hasDot) text = text.slice(0, -1);
    var bits = text.split(".");
    var whole = bits[0] || "";
    var dec = bits[1] || "";
    if (whole === "0" && intPart.indexOf("0") < 0 && (dec || fracPart.length)) whole = "";
    var padZeros = (intPart.match(/0/g) || []).length;
    while (whole.length < padZeros) whole = "0" + whole;
    if (/[0#?]/.test(intPart) && intPart.indexOf(",") >= 0) whole = group(whole);
    text = dec ? whole + "." + dec : hasDot ? whole + "." : whole;
    if (!text) text = "0";
    if (text === ".") text = hasDot ? "." : "0";
    return neg ? "-" + text : text;
  }

  /* ---- fraction patterns (# ?/?) ---------------------------------------- */
  function formatFraction(sec, value) {
    var m = /^([^?0#]*)([#0?]*)\s+([?0#]+)\/([?0#]+)(.*)$/.exec(sec);
    if (!m) return formatNumberSection(sec, value);
    var intLen = m[2].length;
    var numLen = m[3].length;
    var denLen = m[4].length;
    var whole = Math.trunc(Math.abs(value));
    var rest = Math.abs(value) - whole;
    var pq = bestFraction(rest, denLen);
    var num = pq[0];
    var den = pq[1];
    var fracText = num ? String(num) : "";
    var denText = num ? String(den) : "";
    var left = intLen ? String(whole).padStart(intLen, " ") : String(whole);
    if (!intLen) left = "";
    // '#' hides a zero whole part when a fraction follows (" 1/2"), but an
    // integral value still shows its 0 ("0    ")
    if (whole === 0 && m[2].indexOf("0") < 0 && num) left = "";
    var width = numLen + denLen + 1;
    var body = num ? String(num).padStart(numLen, " ") + "/" + String(den).padEnd(denLen, " ") : " ".repeat(width);
    return (value < 0 ? "-" : "") + m[1] + (left ? left + " " : " ") + body + m[5];
  }

  function format(code, value) {
    if (typeof value === "string") return value;
    if (!code || code === GENERAL) return generalText(value);
    if (/\?\/\?/.test(code)) return formatFraction(code, value);

    var sec = sections(code);
    var pick = 0;
    if (value < 0 && sec.length > 1) pick = 1;
    else if (value === 0 && sec.length > 2) pick = 2;
    var chosen = sec[pick] === undefined ? sec[0] : sec[pick];
    if (chosen === "") return "";
    if (chosen === "@") return generalText(value);
    if (isDateCode(chosen)) return formatDate(chosen, value);

    // a dedicated negative section formats the magnitude; section 0 formats
    // the signed value so .5 rounds the way Excel rounds it
    return formatNumberSection(chosen, value < 0 && pick === 0 ? value : Math.abs(value));
  }


  function cellText(c, shared, styles) {
    // a trailing newline inside a cell is invisible in Excel; keep it and the
    // preview grows a blank second line in every cell that has one
    return String(cellValue(c, shared, styles)).replace(/\r?\n+$/, "");
  }

  function cellValue(c, shared, styles) {
    var type = attr(c, "t");
    var v = allText(c, "v")[0];
    if (type === "inlineStr") {
      var is = /<is>([\s\S]*?)<\/is>/.exec(c);
      return unxml(allText(is ? is[1] : c, "t").join(""));
    }
    if (type === "s") return shared[parseInt(v, 10)] || "";
    if (type === "b") return v === "1" ? "TRUE" : "FALSE";
    if (type === "str" || type === "e") return unxml(v || "");
    if (v === undefined || v === "") return "";

    var n = parseFloat(v);
    if (!isFinite(n)) return unxml(v);
    var styleIdx = parseInt(attr(c, "s") || "-1", 10);
    return format(styles[styleIdx] || "", n);
  }

  /* Excel's column width is in characters: 7 px per character plus 5 px of
     padding, the same mapping SheetJS uses */
  function colPx(width) {
    var w = parseFloat(width);
    return w > 0 ? Math.min(600, Math.round(w * 7 + 5)) : 0;
  }

  function sheetTable(xml, shared, styles, maxRows) {
    var rows = [];
    var heights = {};
    var rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
    var m;
    var seen = 0;
    var lastRow = -1;
    var truncated = false;

    // <cols> carries the per-column widths; <sheetFormatPr> the default one
    var widths = [];
    var colsXml = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
    if (colsXml) {
      var colRe = /<col\b[^>]*>/g;
      var cm0;
      while ((cm0 = colRe.exec(colsXml[1])) !== null) {
        var min = parseInt(attr(cm0[0], "min") || "1", 10);
        var max = parseInt(attr(cm0[0], "max") || String(min), 10);
        var px = attr(cm0[0], "hidden") === "1" ? 0 : colPx(attr(cm0[0], "width"));
        for (var ci0 = min; ci0 <= Math.min(max, MAX_COLS); ci0++) widths[ci0 - 1] = px;
      }
    }
    var fmtPr = /<sheetFormatPr\b[^>]*>/.exec(xml);
    var defColW = fmtPr ? colPx(attr(fmtPr[0], "defaultColWidth")) : 0;

    while ((m = rowRe.exec(xml)) !== null) {
      if (++seen > maxRows) {
        truncated = true;
        break;
      }
      var rowIdx = parseInt(attr("<row" + m[1] + ">", "r") || seen, 10) - 1;
      var ht = attr("<row" + m[1] + ">", "ht");
      if (ht) heights[rowIdx] = Math.round(parseFloat(ht) * 96 / 72);
      while (lastRow + 1 < rowIdx && rows.length < maxRows) {
        rows.push([]); // a gap between row numbers still shows as a row
        lastRow++;
      }
      var cells = [];
      var cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      var cm;
      var lastCol = -1;
      while ((cm = cellRe.exec(m[2])) !== null) {
        var tag = "<c" + cm[1] + ">";
        var ref = refToCell(attr(tag, "r"));
        if (!ref) {
          ref = { row: rowIdx, col: lastCol + 1 };
        }
        while (lastCol + 1 < ref.col && cells.length < MAX_COLS) {
          cells.push({ text: "", span: 1, skip: false });
          lastCol++;
        }
        if (cells.length >= MAX_COLS) truncated = true;
        var body = cm[2] || "";
        cells.push({ text: cellText(tag + body + "</c>", shared, styles), span: 1, skip: false });
        lastCol = ref.col;
      }
      rows.push(cells);
      lastRow = rowIdx;
    }

    // merged ranges: the anchor gets colspan/rowspan, the covered cells vanish
    var mergeRe = /<mergeCell\b[^>]*ref="([^"]+)"/g;
    while ((m = mergeRe.exec(xml)) !== null) {
      var parts = m[1].split(":");
      var a = refToCell(parts[0]);
      var b = refToCell(parts[1] || parts[0]);
      if (!a || !b) continue;
      var cell = rows[a.row] && rows[a.row][a.col];
      if (!cell) continue;
      cell.span = Math.max(1, b.col - a.col + 1);
      cell.down = Math.max(1, b.row - a.row + 1);
      for (var r = a.row; r <= b.row; r++) {
        for (var c = a.col; c <= b.col; c++) {
          if (r === a.row && c === a.col) continue;
          if (rows[r] && rows[r][c]) rows[r][c].skip = true;
        }
      }
    }

    var html = ["<table>"];
    var colCount = 0;
    for (var cw = 0; cw < widths.length; cw++) if (widths[cw]) colCount = cw + 1;
    if (colCount) {
      html.push("<colgroup>");
      for (var cc = 0; cc < colCount; cc++) {
        html.push('<col style="width:' + (widths[cc] || defColW || 64) + 'px">');
      }
      html.push("</colgroup>");
    }
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      if (!row.length) continue;
      html.push(heights[ri] ? '<tr style="height:' + heights[ri] + 'px">' : "<tr>");
      for (var ci = 0; ci < row.length; ci++) {
        var cell2 = row[ci];
        if (!cell2 || cell2.skip) continue;
        var span = cell2.span > 1 ? ' colspan="' + cell2.span + '"' : "";
        var down = cell2.down > 1 ? ' rowspan="' + cell2.down + '"' : "";
        html.push("<td" + span + down + ">" + escape(cell2.text) + "</td>");
      }
      html.push("</tr>");
    }
    html.push("</table>");
    return { html: html.join(""), truncated: truncated };
  }

  function escape(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function relTarget(target) {
    if (!target) return "";
    var t = target.replace(/^\.\//, "");
    return t.charAt(0) === "/" ? t.slice(1) : "xl/" + t;
  }

  async function read(buf, options) {
    var maxRows = (options && options.maxRows) || 2000;
    var zip = await root.HFPZip.load(buf);
    var file = function (name) {
      return zip.file(name);
    };
    var str = function (name) {
      var f = file(name);
      return f ? f.async("string") : Promise.resolve("");
    };

    var [wb, rels, sharedXml, stylesXml] = await Promise.all([
      str("xl/workbook.xml"),
      str("xl/_rels/workbook.xml.rels"),
      str("xl/sharedStrings.xml"),
      str("xl/styles.xml"),
    ]);

    var shared = allText(sharedXml, "si").map(function (si) {
      return unxml(allText(si, "t").join(""));
    });
    var styles = cellFormats(stylesXml);

    var targets = {};
    var relRe = /<Relationship\b[^>]*>/g;
    var rm;
    while ((rm = relRe.exec(rels)) !== null) {
      targets[attr(rm[0], "Id")] = relTarget(attr(rm[0], "Target"));
    }

    var sheets = [];
    var sheetRe = /<sheet\b[^>]*>/g;
    var sm;
    while ((sm = sheetRe.exec(wb)) !== null) {
      var name = attr(sm[0], "name");
      var rid = attr(sm[0], "r:id") || attr(sm[0], "id");
      var path = targets[rid];
      if (!path) continue;
      var xml = await str(path);
      if (!xml) continue;
      var table = sheetTable(xml, shared, styles, maxRows);
      sheets.push({
        name: name || "Sheet" + (sheets.length + 1),
        html: table.html,
        truncated: table.truncated,
      });
    }
    return sheets;
  }

  root.HFPxlsx = { read: read, _cell: refToCell, _col: colName, _fmt: format, _styles: cellFormats };
})(typeof globalThis !== "undefined" ? globalThis : this);
