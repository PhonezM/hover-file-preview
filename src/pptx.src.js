/*!
 * OOXML -> DOM renderer for the pptx preview.
 * Short names used throughout: Se=lodash helpers  Tt=__assign  _0=__extends
 * Lm=__spreadArray  oa=__awaiter  sa=__generator  c3e=HFPZip (src/zip.js)  p3e=uuid.v4
 * rVe=echarts.init
 */
/**
 * Minimal replacement for the two lodash functions the library actually uses
 * (`_.get` 279x, `_.omit` 1x). Keeps `Se.get` / `Se.omit` call sites untouched.
 *
 * ponytail: no bracket syntax ("a[0].b") — no call site uses it. Dotted strings
 * are supported because `Se.get(t, "attrs.order")` relies on them; missing that
 * made every node's order 0 and destroyed z-order (tables painted over shapes).
 */

const isIndex = (k) => typeof k === 'number' || /^\d+$/.test(k);

function get(obj, path, defaultValue) {
  if (path == null) return obj === undefined ? defaultValue : obj;
  // lodash accepts a dotted string path ("attrs.order") as well as an array.
  const keys = Array.isArray(path) ? path : String(path).split(".");
  let cur = obj;
  for (const raw of keys) {
    if (cur == null) return defaultValue;
    const k = isIndex(raw) ? Number(raw) : raw;
    cur = cur[k];
  }
  return cur === undefined ? defaultValue : cur;
}

function omit(obj, keys) {
  const drop = new Set(Array.isArray(keys) ? keys : [keys]);
  const out = {};
  for (const k of Object.keys(obj || {})) if (!drop.has(k)) out[k] = obj[k];
  return out;
}

const Se = { get, omit };
import { v4 } from "uuid";
import * as echarts from "echarts";
import {
  __assign,
  __extends,
  __spreadArray,
  __awaiter,
  __generator,
} from "tslib";

const Tt = __assign,
  _0 = __extends,
  Lm = __spreadArray,
  oa = __awaiter,
  sa = __generator;
// the package's own zip reader (src/zip.js) is loaded before this bundle
function c3e() {}
c3e.prototype.loadAsync = function (buffer) {
  return globalThis.HFPZip.load(buffer);
};
const p3e = v4,
  rVe = echarts.init;
var q_ = 1;
function $i(e) {
  var t = 60,
    n = 62,
    r = 45,
    a = 47,
    i = 33,
    o = 39,
    s = 34,
    l = 63,
    u = `\r
	>/= `,
    c = 0;
  return (
    (q_ = 1),
    Iee(
      (function f() {
        for (var d = []; e[c];) {
          if (e.charCodeAt(c) == t) {
            if (e.charCodeAt(c + 1) === a) return ((c = e.indexOf(">", c)), d);
            if (e.charCodeAt(c + 1) === i) {
              if (e.charCodeAt(c + 2) == r) {
                for (
                  ;
                  e.charCodeAt(c) !== n ||
                  e.charCodeAt(c - 1) != r ||
                  e.charCodeAt(c - 2) != r ||
                  c == -1;
                )
                  c = e.indexOf(">", c + 1);
                c === -1 && (c = e.length);
              } else for (c += 2; e.charCodeAt(c) !== n; c++);
              c++;
              continue;
            }
            if (e.charCodeAt(c + 1) === l) {
              ((c = e.indexOf(">", c)), c++);
              continue;
            }
            for (var h = ++c; u.indexOf(e[c]) === -1; c++);
            for (
              var p = e.slice(h, c), g = !1, y = {};
              e.charCodeAt(c) !== n;
              c++
            ) {
              var m = e.charCodeAt(c);
              if ((m > 64 && m < 91) || (m > 96 && m < 123)) {
                for (h = c; u.indexOf(e[c]) === -1; c++);
                for (
                  var v = e.slice(h, c), b = e.charCodeAt(c);
                  b !== o && b !== s;
                )
                  (c++, (b = e.charCodeAt(c)));
                var _ = e[c],
                  S = ++c;
                c = e.indexOf(_, S);
                var C = e.slice(S, c);
                (g || ((y = {}), (g = !0)), (y[v] = C));
              }
            }
            var T = [];
            (e.charCodeAt(c - 1) !== a && (c++, (T = f())),
              d.push({ children: T, tagName: p, attrs: y }));
          } else {
            var k = c;
            (c = e.indexOf("<", c) - 1) === -2 && (c = e.length);
            var M = e.slice(k, c + 1);
            M.length > 0 && d.push(M);
          }
          c++;
        }
        return d;
      })(),
    )
  );
}
function Iee(e) {
  var t = {};
  if (e === void 0) return {};
  if (e.length === 1 && typeof e[0] == "string") return e[0];
  for (var n in (e.forEach(function (r) {
    if ((t[r.tagName] || (t[r.tagName] = []), typeof r == "object")) {
      var a = Iee(r.children);
      (typeof a == "object" &&
        (r.attrs && (a.attrs = r.attrs),
        a.attrs === void 0 ? (a.attrs = { order: q_ }) : (a.attrs.order = q_)),
        q_++,
        t[r.tagName].push(a));
    }
  }),
  t))
    t[n].length == 1 && (t[n] = t[n][0]);
  return t;
}
function Tv(e) {
  return Math.abs(e) > 2e4 ? "emu" : "point";
}
function or(e) {
  return e / 12700;
}
function el(e) {
  return e / 20;
}
function isEmuXfrm(t) {
  var a = (t["a:off"] && t["a:off"].attrs) || {},
    b = (t["a:ext"] && t["a:ext"].attrs) || {},
    n = [a.x, a.y, b.cx, b.cy].map(function (v) {
      return Math.abs(parseInt(v, 10) || 0);
    });
  return n.some(function (v) {
    return v > 2e4;
  });
}
function unitOf(t) {
  // PPT stores EMU, some WPS decks store points. Decide the unit once per shape
  // from all four numbers, never per coordinate: a child sitting at (0,0) inside
  // a group has a tiny offset but an EMU extent, so testing off.x alone (upstream's
  // heuristic) divided the extent by 20 instead of 12700 and blew the shape up 635x.
  // ponytail: upgrade to a per-document flag only if a deck is ever seen mixing both.
  return isEmuXfrm(t) ? or : el;
}
function CS(e) {
  return e / 100;
}
function t2(e) {
  return e / 6e4;
}
function pc(e) {
  return e / 1e5;
}
function Ck(e) {
  var t = Math.ceil(e / 26),
    n = (e % 26 || 26) - 1 + 65;
  return String.fromCharCode(n).repeat(t);
}
function $r(e, t, n) {
  var r = { type: "solidFill" };
  if (e["a:srgbClr"]) r.color = "#" + e["a:srgbClr"].attrs.val;
  else if (e["a:schemeClr"]) {
    var a = e["a:schemeClr"].attrs.val;
    (n && (a = n.getColorThemeName(a)), (r.color = t.getColor(a)));
  } else if (e["a:sysClr"]) r.color = "#" + e["a:sysClr"].attrs.lastClr;
  else if (e["a:prstClr"]) {
    var i = Se.get(e["a:prstClr"], ["attrs", "val"]);
    r.color = (function (d) {
      var h,
        p = [
          "white",
          "AliceBlue",
          "AntiqueWhite",
          "Aqua",
          "Aquamarine",
          "Azure",
          "Beige",
          "Bisque",
          "black",
          "BlanchedAlmond",
          "Blue",
          "BlueViolet",
          "Brown",
          "BurlyWood",
          "CadetBlue",
          "Chartreuse",
          "Chocolate",
          "Coral",
          "CornflowerBlue",
          "Cornsilk",
          "Crimson",
          "Cyan",
          "DarkBlue",
          "DarkCyan",
          "DarkGoldenRod",
          "DarkGray",
          "DarkGrey",
          "DarkGreen",
          "DarkKhaki",
          "DarkMagenta",
          "DarkOliveGreen",
          "DarkOrange",
          "DarkOrchid",
          "DarkRed",
          "DarkSalmon",
          "DarkSeaGreen",
          "DarkSlateBlue",
          "DarkSlateGray",
          "DarkSlateGrey",
          "DarkTurquoise",
          "DarkViolet",
          "DeepPink",
          "DeepSkyBlue",
          "DimGray",
          "DimGrey",
          "DodgerBlue",
          "FireBrick",
          "FloralWhite",
          "ForestGreen",
          "Fuchsia",
          "Gainsboro",
          "GhostWhite",
          "Gold",
          "GoldenRod",
          "Gray",
          "Grey",
          "Green",
          "GreenYellow",
          "HoneyDew",
          "HotPink",
          "IndianRed",
          "Indigo",
          "Ivory",
          "Khaki",
          "Lavender",
          "LavenderBlush",
          "LawnGreen",
          "LemonChiffon",
          "LightBlue",
          "LightCoral",
          "LightCyan",
          "LightGoldenRodYellow",
          "LightGray",
          "LightGrey",
          "LightGreen",
          "LightPink",
          "LightSalmon",
          "LightSeaGreen",
          "LightSkyBlue",
          "LightSlateGray",
          "LightSlateGrey",
          "LightSteelBlue",
          "LightYellow",
          "Lime",
          "LimeGreen",
          "Linen",
          "Magenta",
          "Maroon",
          "MediumAquaMarine",
          "MediumBlue",
          "MediumOrchid",
          "MediumPurple",
          "MediumSeaGreen",
          "MediumSlateBlue",
          "MediumSpringGreen",
          "MediumTurquoise",
          "MediumVioletRed",
          "MidnightBlue",
          "MintCream",
          "MistyRose",
          "Moccasin",
          "NavajoWhite",
          "Navy",
          "OldLace",
          "Olive",
          "OliveDrab",
          "Orange",
          "OrangeRed",
          "Orchid",
          "PaleGoldenRod",
          "PaleGreen",
          "PaleTurquoise",
          "PaleVioletRed",
          "PapayaWhip",
          "PeachPuff",
          "Peru",
          "Pink",
          "Plum",
          "PowderBlue",
          "Purple",
          "RebeccaPurple",
          "Red",
          "RosyBrown",
          "RoyalBlue",
          "SaddleBrown",
          "Salmon",
          "SandyBrown",
          "SeaGreen",
          "SeaShell",
          "Sienna",
          "Silver",
          "SkyBlue",
          "SlateBlue",
          "SlateGray",
          "SlateGrey",
          "Snow",
          "SpringGreen",
          "SteelBlue",
          "Tan",
          "Teal",
          "Thistle",
          "Tomato",
          "Turquoise",
          "Violet",
          "Wheat",
          "White",
          "WhiteSmoke",
          "Yellow",
          "YellowGreen",
        ],
        g = [
          "ffffff",
          "f0f8ff",
          "faebd7",
          "00ffff",
          "7fffd4",
          "f0ffff",
          "f5f5dc",
          "ffe4c4",
          "000000",
          "ffebcd",
          "0000ff",
          "8a2be2",
          "a52a2a",
          "deb887",
          "5f9ea0",
          "7fff00",
          "d2691e",
          "ff7f50",
          "6495ed",
          "fff8dc",
          "dc143c",
          "00ffff",
          "00008b",
          "008b8b",
          "b8860b",
          "a9a9a9",
          "a9a9a9",
          "006400",
          "bdb76b",
          "8b008b",
          "556b2f",
          "ff8c00",
          "9932cc",
          "8b0000",
          "e9967a",
          "8fbc8f",
          "483d8b",
          "2f4f4f",
          "2f4f4f",
          "00ced1",
          "9400d3",
          "ff1493",
          "00bfff",
          "696969",
          "696969",
          "1e90ff",
          "b22222",
          "fffaf0",
          "228b22",
          "ff00ff",
          "dcdcdc",
          "f8f8ff",
          "ffd700",
          "daa520",
          "808080",
          "808080",
          "008000",
          "adff2f",
          "f0fff0",
          "ff69b4",
          "cd5c5c",
          "4b0082",
          "fffff0",
          "f0e68c",
          "e6e6fa",
          "fff0f5",
          "7cfc00",
          "fffacd",
          "add8e6",
          "f08080",
          "e0ffff",
          "fafad2",
          "d3d3d3",
          "d3d3d3",
          "90ee90",
          "ffb6c1",
          "ffa07a",
          "20b2aa",
          "87cefa",
          "778899",
          "778899",
          "b0c4de",
          "ffffe0",
          "00ff00",
          "32cd32",
          "faf0e6",
          "ff00ff",
          "800000",
          "66cdaa",
          "0000cd",
          "ba55d3",
          "9370db",
          "3cb371",
          "7b68ee",
          "00fa9a",
          "48d1cc",
          "c71585",
          "191970",
          "f5fffa",
          "ffe4e1",
          "ffe4b5",
          "ffdead",
          "000080",
          "fdf5e6",
          "808000",
          "6b8e23",
          "ffa500",
          "ff4500",
          "da70d6",
          "eee8aa",
          "98fb98",
          "afeeee",
          "db7093",
          "ffefd5",
          "ffdab9",
          "cd853f",
          "ffc0cb",
          "dda0dd",
          "b0e0e6",
          "800080",
          "663399",
          "ff0000",
          "bc8f8f",
          "4169e1",
          "8b4513",
          "fa8072",
          "f4a460",
          "2e8b57",
          "fff5ee",
          "a0522d",
          "c0c0c0",
          "87ceeb",
          "6a5acd",
          "708090",
          "708090",
          "fffafa",
          "00ff7f",
          "4682b4",
          "d2b48c",
          "008080",
          "d8bfd8",
          "ff6347",
          "40e0d0",
          "ee82ee",
          "f5deb3",
          "ffffff",
          "f5f5f5",
          "ffff00",
          "9acd32",
        ],
        y = p.indexOf(d);
      return (y != -1 && (h = g[y]), "#".concat(h || "000000"));
    })(i);
  }
  var o = e["a:srgbClr"] || e["a:schemeClr"] || e["a:sysClr"],
    s = Se.get(o, ["a:alpha", "attrs", "val"], 1e5);
  r.alpha = s / 1e5;
  var l = Se.get(o, ["a:shade", "attrs", "val"]);
  l && (r.shade = l / 1e5);
  var u = Se.get(o, ["a:lumMod", "attrs", "val"]);
  u && (r.lumMod = u / 1e5);
  var c = Se.get(o, ["a:lumOff", "attrs", "val"]);
  c && (r.lumOff = c / 1e5);
  var f = Se.get(o, ["a:tint", "attrs", "val"]);
  return (f && (r.tint = f / 1e5), r);
}
function W0(e, t, n) {
  var r,
    a = { type: "blipFill" },
    i = Se.get(e, ["a:blip", "attrs", "r:embed"]);
  if (i) {
    var o = (r = n.rels[i]) === null || r === void 0 ? void 0 : r.target;
    o && (a.base64 = t.getMedia(o));
  }
  var s = Se.get(e, ["a:blip", "a:alphaModFix", "attrs", "amt"]);
  s && (a.alpha = s / 1e5);
  var l = Se.get(e, ["a:stretch", "a:fillRect", "attrs"]);
  return (
    l &&
      ((a.fillRect = {}),
      l.b && (a.fillRect.b = l.b / 1e5),
      l.t && (a.fillRect.t = l.t / 1e5),
      l.r && (a.fillRect.r = l.r / 1e5),
      l.l && (a.fillRect.l = l.l / 1e5)),
    a
  );
}
function G0(e, t, n) {
  var r = { type: "gradFill", tileRect: {}, lin: {}, gsList: [] };
  ((r.flip = e.attrs.flip),
    (r.path = Se.get(e, ["a:path", "attrs", "path"]) || "linear"),
    (r.rotWithShape = e.attrs.rotWithShape === "1"),
    Se.get(e, ["a:lin", "attrs", "ang"]) &&
      (r.lin.ang = t2(e["a:lin"].attrs.ang)),
    Se.get(e, ["a:lin", "attrs", "scaled"]) &&
      (r.lin.scaled = e["a:lin"].attrs.scaled === "1"));
  var a = Se.get(e, ["a:gsLst", "a:gs"]) || [];
  return (
    (r.gsList = a.map(function (i) {
      return { color: $r(i, t, n), pos: pc(i.attrs.pos) };
    })),
    Se.get(e, ["a:tileRect", "attrs", "l"]) &&
      (r.tileRect.l = pc(e["a:tileRect"].attrs.l)),
    Se.get(e, ["a:tileRect", "attrs", "t"]) &&
      (r.tileRect.t = pc(e["a:tileRect"].attrs.t)),
    Se.get(e, ["a:tileRect", "attrs", "r"]) &&
      (r.tileRect.r = pc(e["a:tileRect"].attrs.r)),
    Se.get(e, ["a:tileRect", "attrs", "b"]) &&
      (r.tileRect.b = pc(e["a:tileRect"].attrs.b)),
    r
  );
}
function Tk(e) {
  return e <= 0.04045 ? e / 12.92 : Math.pow((e + 0.055) / 1.055, 2.4);
}
function Ak(e) {
  return e < 0.0031308 ? 12.92 * e : 1.055 * Math.pow(e, 1 / 2.4) - 0.055;
}
function cW(e, t) {
  var n = e[0],
    r = e[1],
    a = e[2],
    i = Tk(n / 255) * t,
    o = Tk(r / 255) * t,
    s = Tk(a / 255) * t;
  return [
    Math.round(255 * Ak(i)),
    Math.round(255 * Ak(o)),
    Math.round(255 * Ak(s)),
  ];
}
function kk(e, t, n) {
  var r = e / 255,
    a = t / 255,
    i = n / 255,
    o = Math.max(r, a, i),
    s = Math.min(r, a, i),
    l = o - s,
    u = 0,
    c = (o + s) / 2;
  return (
    l === 0
      ? (u = 0)
      : o === r
        ? (u = ((a - i) / l) % 6)
        : o === a
          ? (u = (i - r) / l + 2)
          : o === i && (u = (r - a) / l + 4),
    (u = Math.round(60 * u)) < 0 && (u += 360),
    {
      h: u,
      s: l === 0 || c === 0 || c === 1 ? 0 : l / (1 - Math.abs(2 * c - 1)),
      l: c,
    }
  );
}
function Mk(e, t, n) {
  var r,
    a,
    i,
    o = (1 - Math.abs(2 * n - 1)) * t,
    s = o * (1 - Math.abs(((e / 60) % 2) - 1)),
    l = n - o / 2;
  return (
    e < 60
      ? ((r = o), (a = s), (i = 0))
      : e < 120
        ? ((r = s), (a = o), (i = 0))
        : e < 180
          ? ((r = 0), (a = o), (i = s))
          : e < 240
            ? ((r = 0), (a = s), (i = o))
            : e < 300
              ? ((r = s), (a = 0), (i = o))
              : ((r = o), (a = 0), (i = s)),
    [
      (r = Math.round(255 * (r + l))),
      (a = Math.round(255 * (a + l))),
      (i = Math.round(255 * (i + l))),
    ]
  );
}
function resolveTypeface(e, t) {
  // `+mn-lt` / `+mj-lt` are placeholders for the theme's minor/major latin font.
  // Upstream only ever read `a:ea`, so Latin runs ended up with no font-family at
  // all and the browser fell back to its default serif (all text 14% too narrow).
  var n = (t && t.fontScheme) || {},
    r = n.minor || {},
    a = n.major || {};
  return e === "+mn-lt"
    ? r.latin || r.ea
    : e === "+mj-lt"
      ? a.latin || a.ea
      : e === "+mn-ea"
        ? r.ea || r.latin
        : e === "+mj-ea"
          ? a.ea || a.latin
          : e || "";
}
function Qn(e, t) {
  if (!e || e.type === "none") return "";
  if (e.type === "solidFill" && /^#[\da-fA-F]{3,6}$/.test(e.color)) {
    var n = parseInt(e.color.substr(1, 2), 16),
      r = parseInt(e.color.substr(3, 2), 16),
      a = parseInt(e.color.substr(5, 2), 16);
    if (e.shade) {
      var i = cW([n, r, a], e.shade);
      ((n = i[0]), (r = i[1]), (a = i[2]));
    }
    if (e.lumMod) {
      var o = (function (c, f, d, h) {
        var p = kk(c, f, d),
          g = p.h,
          y = p.l * h;
        return (y >= 1 && (y = 1), Mk(g, p.s, y));
      })(n, r, a, e.lumMod);
      ((n = o[0]), (r = o[1]), (a = o[2]));
    }
    if (
      (e.lumOff &&
        ((o = (function (c, f, d, h) {
          var p = kk(c, f, d),
            g = p.h,
            y = h + p.l;
          return (y > 1 && (y = 1), Mk(g, p.s, y));
        })(n, r, a, e.lumOff)),
        (n = o[0]),
        (r = o[1]),
        (a = o[2])),
      e.tint || (t != null && t.light))
    ) {
      // PowerPoint blends `a:tint` toward white in linear light, not in HSL.
      // Doing it in HSL (upstream) yields #B4C7E7 for tint 40000 where PowerPoint
      // paints #CFD5EA - every table band came out one step too saturated.
      // ponytail: fixed 2.2 gamma; PowerPoint is not colour-managed beyond this.
      var s = (function (c, f, d, h) {
        h >= 1 && (h = 1);
        var g = function (v) {
            return Math.pow(Math.max(0, Math.min(255, v)) / 255, 2.2);
          },
          y = function (v) {
            return Math.round(
              255 * Math.pow(Math.max(0, Math.min(1, v)), 1 / 2.2),
            );
          };
        return [
          y(g(c) * h + (1 - h)),
          y(g(f) * h + (1 - h)),
          y(g(d) * h + (1 - h)),
        ];
      })(n, r, a, e.tint || (t == null ? void 0 : t.light));
      ((n = s[0]), (r = s[1]), (a = s[2]));
    }
    if (t != null && t.dark) {
      var l = cW([n, r, a], t == null ? void 0 : t.dark);
      ((n = l[0]), (r = l[1]), (a = l[2]));
    }
    var u = e.alpha;
    return "rgba(".concat(n, ",").concat(r, ",").concat(a, ",").concat(u, ")");
  }
}
var U0 = (function () {
    function e(t, n, r) {
      var a, i;
      ((this.uuid = p3e()),
        (this.offset = { x: 0, y: 0 }),
        (this.extend = { w: 0, h: 0 }),
        (this.rotate = 0),
        (this.order = 0),
        (this.flipV = !1),
        (this.flipH = !1),
        (this.source = t),
        (this.ctx = n),
        (this.group = r));
      var o = Se.get(t, ["p:nvSpPr", "p:nvPr"]);
      if (o) {
        var s = Se.get(o, "p:ph");
        (s && s.attrs && ((this.idx = s.attrs.idx), (this.type = s.attrs.type)),
          Se.get(o, ["attrs", "userDrawn"]) &&
            (this.userDrawn = Se.get(o, ["attrs", "userDrawn"]) === "1"));
      }
      if (((this.order = Se.get(t, "attrs.order", 0)), this.source["p:spPr"])) {
        if ((u = this.getXfrm())) {
          var l = this.group ? (this.ctx.pptx && this.ctx.pptx.wps ? el : unitOf(u)) : or;
          ((this.offset = {
            x: Math.round(l(parseInt(u["a:off"].attrs.x))),
            y: Math.round(l(parseInt(u["a:off"].attrs.y))),
          }),
            (this.extend = {
              w: Math.round(l(parseInt(u["a:ext"].attrs.cx))),
              h: Math.round(l(parseInt(u["a:ext"].attrs.cy || "0"))),
            }),
            (this.rotate = t2(parseInt(Se.get(u, "attrs.rot", 0)))),
            (this.flipV = Se.get(u, "attrs.flipV") === "1"),
            (this.flipH = Se.get(u, "attrs.flipH") === "1"));
        }
      } else if (this.source["p:xfrm"]) {
        var u = this.source["p:xfrm"];
        ((l = this.group ? (this.ctx.pptx && this.ctx.pptx.wps ? el : unitOf(u)) : or),
          (this.offset = {
            x: Math.round(l(parseInt(u["a:off"].attrs.x))),
            y: Math.round(l(parseInt(u["a:off"].attrs.y))),
          }),
          (this.extend = {
            w: Math.round(l(parseInt(u["a:ext"].attrs.cx))),
            h: Math.round(l(parseInt(u["a:ext"].attrs.cy))),
          }));
      }
      if (r) {
        var c = r.extend,
          f = r.chExtend,
          d = r.chOffset,
          h = f.w === 0 ? 0 : c.w / f.w,
          p = f.h === 0 ? 0 : c.h / f.h;
        ((this.extend.w = this.extend.w * h),
          (this.extend.h = this.extend.h * p),
          (this.offset.x = (this.offset.x - d.x) * h),
          (this.offset.y = (this.offset.y - d.y) * p));
      }
    }
    return (
      Object.defineProperty(e.prototype, "theme", {
        get: function () {
          return (this.ctx.sliderMaster || this.ctx).theme;
        },
        enumerable: !1,
        configurable: !0,
      }),
      (e.prototype.getColorThemeName = function (t) {
        return this.ctx.getColorThemeName(t);
      }),
      (e.prototype.getXfrm = function () {
        var t = this.source["p:spPr"]["a:xfrm"];
        return (
          t ||
            (this.idx
              ? (t = this.ctx.getNodeInheritAttrsByIdx(this.idx, [
                  "p:spPr",
                  "a:xfrm",
                ]))
              : this.type &&
                (t = this.ctx.getNodeInheritAttrsByType(this.type, [
                  "p:spPr",
                  "a:xfrm",
                ]))),
          t
        );
      }),
      e
    );
  })(),
  cJe = (function () {
    function e(t, n) {
      ((this.props = {}),
        (this.inheritProps = {}),
        (this.source = t),
        (this.node = n),
        this._getInheritBodyProps(),
        this._parseBodyProps(),
        this._parseLstStyle(),
        this._parseText());
      // PowerPoint's last resort for a run that inherits no font is the theme's
      // minor latin face. Upstream left it unset, so the browser used its default
      // serif and every Latin string came out ~14% too narrow.
      var r = n && n.theme && n.theme.fontScheme && n.theme.fontScheme.minor;
      r &&
        r.latin &&
        !this.inheritProps.typeface &&
        !this.props.typeface &&
        ((this.inheritProps.typeface = r.latin),
        r.ea && (this.inheritProps.typefaceEa = r.ea));
    }
    return (
      (e.prototype._getInheritBodyProps = function () {
        var t,
          n = this.node.ctx,
          r = this.node.type,
          a = this.node.idx;
        if (r || a)
          switch (n.slideType) {
            case "slideMaster":
              break;
            case "slideLayout":
              (t = r
                ? n.slideMaster.getNodeByType(r)
                : n.slideMaster.getNodeByIdx(a)) &&
                (this.inheritProps = Se.get(t, ["textBody", "props"]) || {});
              break;
            case "slide":
              ((t = r
                ? n.slideLayout.slideMaster.getNodeByType(r)
                : n.slideLayout.slideMaster.getNodeByIdx(a)) &&
                Object.assign(
                  this.inheritProps,
                  Se.get(t, ["textBody", "props"]) || {},
                ),
                (t = r
                  ? n.slideLayout.getNodeByType(r)
                  : n.slideLayout.getNodeByIdx(a)) &&
                  Object.assign(
                    this.inheritProps,
                    Se.get(t, ["textBody", "props"]) || {},
                  ));
          }
      }),
      (e.prototype._parseBodyProps = function () {
        var t = this,
          n = Se.get(this.source, ["a:bodyPr", "attrs"]) || {};
        Object.keys(n).forEach(function (o) {
          switch (o) {
            case "anchor":
              t.props.anchor = n[o];
              break;
            case "rtlCol":
              t.props.rtlCol = n[o] === "1";
              break;
            case "lIns":
            case "rIns":
            case "tIns":
            case "bIns":
              t.props[o] = or(parseInt(n[o]));
              break;
            case "order":
              break;
            default:
              t.props[o] = n[o];
          }
        });
        var r = Se.get(this.source, ["a:bodyPr", "a:normAutofit", "attrs"]);
        if (r) {
          this.props.normAutofit = {};
          var a = r.fontScale;
          a && (this.props.normAutofit.fontScale = pc(parseInt(a)));
          var i = r.lnSpcReduction;
          i && (this.props.normAutofit.lnSpcReduction = pc(parseInt(i)));
        }
      }),
      (e.prototype._parseLstStyle = function () {
        var t = this,
          n = {},
          r = Se.get(this.source, "a:lstStyle") || {};
        (Object.keys(r).forEach(function (a) {
          if (a.startsWith("a:") && a.endsWith("pPr")) {
            var i = a.substr(2, a.length - 5);
            n[i] = { props: t._formatPPr(r[a]) };
            var o = Se.get(r[a], ["a:defRPr"]);
            n[i].defRPr = t._formatRPr(o);
          }
        }),
          (this.lstStyle = n));
      }),
      (e.prototype._parseText = function () {
        var t = this,
          n = Se.get(this.source, ["a:p"]) || [];
        (Array.isArray(n) || (n = [n]),
          (this.paragraphs = n.map(function (r) {
            return t._parseParagraph(r);
          })));
      }),
      (e.prototype._parseParagraph = function (t) {
        var n = this,
          r = {
            props: {},
            inheritProps: {},
            inheritRProps: {},
            endParaRProps: {},
            rows: [],
          },
          a = Se.get(t, ["a:pPr"]) || {};
        r.props = this._formatPPr(a);
        var i = Se.get(t, ["a:endParaRPr"]);
        r.endParaRProps = this._formatRPr(i);
        var o = Se.get(t, ["a:r"]) || [];
        Array.isArray(o) || (o = [o]);
        // a:fld is a run with generated text (slide number, date). Upstream ignored
        // it, so every slide number in the deck was missing.
        var fld = Se.get(t, ["a:fld"]);
        fld && (o = o.concat(Array.isArray(fld) ? fld : [fld]));
        var s = Se.get(t, ["a:br"]) || [];
        return (
          Array.isArray(s) || (s = [s]),
          (o = o.concat(
            s.map(function (l) {
              return Tt({ isBr: !0 }, l);
            }),
          )).sort(function (l, u) {
            return (
              Se.get(l, ["attrs", "order"]) - Se.get(u, ["attrs", "order"])
            );
          }),
          (r.rows = o.map(function (l) {
            return n._parseRow(l);
          })),
          (r.inheritProps = this._getInheritPProps(r.props.level)),
          (r.inheritRProps = this._getInheritRProps(r.props.level)),
          r
        );
      }),
      (e.prototype._getInheritPProps = function (t) {
        t === void 0 && (t = "0");
        var n,
          r = {},
          a = this.node.ctx,
          i = this.node.type,
          o = this.node.idx;
        switch (a.slideType) {
          case "slideMaster":
            this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                );
            break;
          case "slideLayout":
            (this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.slideMaster.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a.slideMaster, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                ),
              (i || o) &&
                (n = i
                  ? a.slideMaster.getNodeByType(i)
                  : a.slideMaster.getNodeByIdx(o)) &&
                Object.assign(
                  r,
                  Se.get(n, [
                    "textBody",
                    "lstStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                ));
            break;
          case "slide":
            (this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.slideLayout.slideMaster.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a.slideLayout.slideMaster, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "props",
                  ]) || {},
                ),
              (i || o) &&
                (["subTitle", "ctrTitle", "title"].includes(i) &&
                  Object.assign(
                    r,
                    Se.get(a.slideLayout.slideMaster, [
                      "textStyles",
                      "titleStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "props",
                    ]) || {},
                  ),
                (n = i
                  ? a.slideLayout.slideMaster.getNodeByType(i)
                  : a.slideLayout.slideMaster.getNodeByIdx(o)) &&
                  Object.assign(
                    r,
                    Se.get(n, [
                      "textBody",
                      "lstStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "props",
                    ]) || {},
                  ),
                (n = i
                  ? a.slideLayout.getNodeByType(i)
                  : a.slideLayout.getNodeByIdx(o)) &&
                  Object.assign(
                    r,
                    Se.get(n, [
                      "textBody",
                      "lstStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "props",
                    ]) || {},
                  )));
        }
        return r;
      }),
      (e.prototype._getInheritRProps = function (t) {
        t === void 0 && (t = "0");
        var n,
          r = {},
          a = this.node.ctx,
          i = this.node.type,
          o = this.node.idx;
        switch (a.slideType) {
          case "slideMaster":
            this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                );
            break;
          case "slideLayout":
            (this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.slideMaster.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a.slideMaster, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                ),
              (i || o) &&
                (n = i
                  ? a.slideMaster.getNodeByType(i)
                  : a.slideMaster.getNodeByIdx(o)) &&
                (r =
                  Se.get(n, [
                    "textBody",
                    "lstStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {}));
            break;
          case "slide":
            (this.node.isTextBox
              ? Object.assign(
                  r,
                  Se.get(a.slideLayout.slideMaster.defaultTextStyle, [
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                )
              : Object.assign(
                  r,
                  Se.get(a.slideLayout.slideMaster, [
                    "textStyles",
                    "otherStyle",
                    "lvl".concat(t ? +t + 1 : 1),
                    "defRPr",
                  ]) || {},
                ),
              (i || o) &&
                (["subTitle", "ctrTitle", "title"].includes(i) &&
                  Object.assign(
                    r,
                    Se.get(a.slideLayout.slideMaster, [
                      "textStyles",
                      "titleStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "defRPr",
                    ]) || {},
                  ),
                (n = i
                  ? a.slideLayout.slideMaster.getNodeByType(i)
                  : a.slideLayout.slideMaster.getNodeByIdx(o)) &&
                  Object.assign(
                    r,
                    Se.get(n, [
                      "textBody",
                      "lstStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "defRPr",
                    ]) || {},
                  ),
                (n = i
                  ? a.slideLayout.getNodeByType(i)
                  : a.slideLayout.getNodeByIdx(o)) &&
                  Object.assign(
                    r,
                    Se.get(n, [
                      "textBody",
                      "lstStyle",
                      "lvl".concat(t ? +t + 1 : 1),
                      "defRPr",
                    ]) || {},
                  )));
        }
        var s = Se.get(this.node.source, ["p:style", "a:fontRef"]);
        return (
          Se.get(s, "a:schemeClr") &&
            (r.color = $r(s, this.node.theme, this.node)),
          Se.get(this.lstStyle, ["lvl".concat(t ? +t + 1 : 1), "defRPr"]) &&
            Object.assign(
              r,
              Se.get(this.lstStyle, ["lvl".concat(t ? +t + 1 : 1), "defRPr"]),
            ),
          r
        );
      }),
      (e.prototype._formatPPr = function (t) {
        var n = {},
          r = Se.get(t, "attrs") || {};
        return (
          Object.keys(r).forEach(function (a) {
            switch (a) {
              case "algn":
                n.align = r[a];
                break;
              case "marL":
                n.marginLeft = or(parseInt(r[a]));
                break;
              case "indent":
                n.indent = or(parseInt(r[a]));
                break;
              case "lvl":
                n.level = r[a];
            }
          }),
          Se.get(t, ["a:lnSpc", "a:spcPct", "attrs", "val"]) &&
            (n.lineHeight = parseInt(t["a:lnSpc"]["a:spcPct"].attrs.val) / 1e5),
          Se.get(t, ["a:buAutoNum", "attrs", "type"]) &&
            (n.buAutoNum = t["a:buAutoNum"].attrs.type),
          Se.get(t, ["a:buChar", "attrs", "char"]) &&
            (n.buChar = t["a:buChar"].attrs.char),
          Se.get(t, ["a:spcBef", "a:spcPts", "attrs", "val"]) &&
            (n.spaceBefore = CS(parseInt(t["a:spcBef"]["a:spcPts"].attrs.val))),
          Se.get(t, ["a:spcAft", "a:spcPts", "attrs", "val"]) &&
            (n.spaceAfter = CS(parseInt(t["a:spcAft"]["a:spcPts"].attrs.val))),
          Se.get(t, ["a:defRPr"]) &&
            (n.defRPr = this._formatRPr(Se.get(t, ["a:defRPr"]))),
          n
        );
      }),
      (e.prototype._parseRow = function (t) {
        if (t.isBr) return { isBr: !0 };
        var n = { props: {}, text: "" },
          r = Se.get(t, ["a:rPr"]) || {};
        return (
          (n.props = this._formatRPr(r)),
          (n.text = Se.get(t, "a:t") || ""),
          n
        );
      }),
      (e.prototype._formatRPr = function (t) {
        var n = {},
          r = Se.get(t, "attrs") || {};
        Object.keys(r).forEach(function (o) {
          switch (o) {
            case "sz":
              n.size = parseInt(r[o]) / 100;
              break;
            case "b":
              n.bold = r[o] === "1";
              break;
            case "i":
              n.italic = r[o] === "1";
              break;
            case "u":
              n.underline = r[o];
              break;
            case "strike":
              n.strike = r[o];
              break;
            case "order":
            case "dirty":
              break;
            default:
              n[o] = r[o];
          }
        });
        var a = Se.get(t, "a:solidFill");
        a && (n.color = $r(a, this.node.theme, this.node));
        var hl = Se.get(t, "a:hlinkClick"),
          th = this.node.theme;
        hl &&
          ((n.color = {
            type: "solidFill",
            color: (th && th.clrScheme && th.clrScheme.hlink) || "#0563C1",
            alpha: 1,
          }),
          (n.underline = "sng"));
        var i = Se.get(t, "a:highlight");
        return (
          i && (n.background = $r(i, this.node.theme, this.node)),
          (o = Se.get(t, ["a:latin", "attrs", "typeface"])) &&
            (n.typeface = resolveTypeface(o, this.node.theme)),
          (o = Se.get(t, ["a:ea", "attrs", "typeface"])) &&
            (n.typefaceEa = resolveTypeface(o, this.node.theme)),
          n
        );
      }),
      e
    );
  })();
function Av(e, t, n) {
  var r = {};
  if (!Se.get(e, "a:noFill")) {
    Se.get(e, "attrs.w") && (r.width = or(parseInt(Se.get(e, "attrs.w"))));
    var a = Se.get(e, "a:solidFill");
    a && (r.color = $r(a, t, n));
    var i = Se.get(e, "a:prstDash");
    if (
      (i && (r.type = i.attrs.val),
      Se.get(e, ["a:miter"]) && (r.lineJoin = "miter"),
      Se.get(e, ["a:bevel"]) && (r.lineJoin = "bevel"),
      Se.get(e, ["a:round"]) && (r.lineJoin = "round"),
      Se.get(e, ["a:miter", "attrs", "lim"]) &&
        (r.miterLim = or(parseInt(Se.get(e, ["a:miter", "attrs", "lim"])))),
      Se.get(e, ["a:headEnd"]))
    ) {
      var o = Se.get(e, ["a:headEnd", "attrs"]);
      r.headEnd = { type: o.type, len: o.len, w: o.w };
    }
    if (Se.get(e, ["a:tailEnd"])) {
      var s = Se.get(e, ["a:tailEnd", "attrs"]);
      r.tailEnd = { type: s.type, len: s.len, w: s.w };
    }
    return r;
  }
}
var Jy = (function (e) {
    function t(n, r, a, i) {
      var o = e.call(this, n, a, i) || this;
      return (
        (o.border = {}),
        (o.prstGeom = {}),
        (o.isTextBox = !1),
        (o.pptx = r),
        o._parseShape(),
        o._parIsTextBox(),
        o._parsePrstGeom(),
        o._parseBackground(),
        o._parseBorder(),
        o._parseTxt(),
        o
      );
    }
    return (
      _0(t, e),
      (t.prototype._parseShape = function () {
        if (
          ((this.shape = Se.get(this.source, [
            "p:spPr",
            "a:prstGeom",
            "attrs",
            "prst",
          ])),
          !this.shape && Se.get(this.source, ["p:spPr", "a:custGeom"]))
        ) {
          this.shape = "customGeom";
          var n = Se.get(this.source, [
              "p:spPr",
              "a:custGeom",
              "a:pathLst",
              "a:path",
            ]),
            r = [],
            a = [],
            i = function (s) {
              switch (s) {
                case "a:moveTo":
                case "a:cubicBezTo":
                case "a:lnTo":
                  ((a = Array.isArray(n[s]) ? n[s] : [n[s]]),
                    (r = r.concat(
                      a.map(function (l) {
                        return {
                          order: l.attrs.order,
                          type: s.split(":")[1],
                          points: (Array.isArray(l["a:pt"])
                            ? l["a:pt"]
                            : [l["a:pt"]]
                          ).map(function (u) {
                            return [
                              or(parseInt(Se.get(u, ["attrs", "x"]))),
                              or(parseInt(Se.get(u, ["attrs", "y"]))),
                            ];
                          }),
                        };
                      }),
                    )));
                  break;
                case "a:close":
                  ((a = Array.isArray(n[s]) ? n[s] : [n[s]]),
                    (r = r.concat(
                      a.map(function (l) {
                        return { order: l.attrs.order, type: s.split(":")[1] };
                      }),
                    )));
              }
            };
          for (var o in n) i(o);
          (r.sort(function (s, l) {
            return s.order - l.order;
          }),
            (this.prstGeom.pathList = r),
            Se.get(n, ["attrs", "w"]) &&
              (this.prstGeom.w = or(parseInt(Se.get(n, ["attrs", "w"])))),
            Se.get(n, ["attrs", "h"]) &&
              (this.prstGeom.h = or(parseInt(Se.get(n, ["attrs", "h"])))));
        }
      }),
      (t.prototype._parIsTextBox = function () {
        this.isTextBox =
          Se.get(this.source, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]) ===
          "1";
      }),
      (t.prototype._parsePrstGeom = function () {
        var n = this,
          r = Se.get(this.source, ["p:spPr", "a:prstGeom"]),
          a = Se.get(r, ["a:avLst", "a:gd"]);
        a &&
          (Array.isArray(a) || (a = [a]),
          (this.prstGeom.gd = a.map(function (i) {
            var o =
              ["pie", "chord", "arc"].includes(n.shape) ||
              (["blockArc"].includes(n.shape) &&
                ["adj1", "adj2"].includes(i.attrs.name))
                ? t2(parseInt(i.attrs.fmla.split(" ")[1]))
                : pc(parseInt(i.attrs.fmla.split(" ")[1]));
            return { name: i.attrs.name, fmla: o };
          })));
      }),
      (t.prototype._parseBackground = function () {
        if (!Se.get(this.source, ["p:spPr", "a:noFill"]))
          if (Se.get(this.source, ["p:spPr", "a:grpFill"]) && this.group)
            this.background = this.group.getBackground();
          else {
            var n = Se.get(this.source, ["p:spPr", "a:solidFill"]);
            if (n) this.background = $r(n, this.theme, this);
            else {
              var r = Se.get(this.source, ["p:spPr", "a:gradFill"]);
              if (r) this.background = G0(r, this.theme, this);
              else {
                var a = Se.get(this.source, ["p:spPr", "a:blipFill"]);
                if (a) this.background = W0(a, this.pptx, this.ctx);
                else {
                  var i = Se.get(this.source, ["p:style", "a:fillRef"]);
                  i && (this.background = $r(i, this.theme, this));
                }
              }
            }
          }
      }),
      (t.prototype._parseBorder = function () {
        var n = Se.get(this.source, ["p:style", "a:lnRef"]);
        if (n) {
          var r = parseInt(n.attrs.idx),
            a = this.theme.getLineStyle(r);
          ((this.border = Tt(Tt({}, a), this.border)),
            (this.border.color && this.border.color.color) ||
              (this.border.color = $r(n, this.theme, this)));
        }
        var i = Se.get(this.source, ["p:spPr", "a:ln"]);
        (i && Object.assign(this.border, Av(i, this.theme, this)),
          this.border.color &&
            this.border.color.color &&
            !this.border.width &&
            (this.border.width = t.defaultBorderWidth));
      }),
      (t.prototype._parseTxt = function () {
        this.textBody = new cJe(Se.get(this.source, ["p:txBody"]), this);
      }),
      (t.defaultBorderWidth = 0.75),
      t
    );
  })(U0),
  TS = (function (e) {
    function t(n, r, a, i, o) {
      var s,
        l,
        u = e.call(this, r, i, o) || this;
      ((u.userDrawn = !0), (u.pptx = a), (u.path = n));
      var c = Se.get(u.source, ["p:blipFill", "a:srcRect"]);
      c &&
        ((u.clip = {}),
        c.attrs.b && (u.clip.b = parseInt(c.attrs.b) / 1e5),
        c.attrs.t && (u.clip.t = parseInt(c.attrs.t) / 1e5),
        c.attrs.l && (u.clip.l = parseInt(c.attrs.l) / 1e5),
        c.attrs.r && (u.clip.r = parseInt(c.attrs.r) / 1e5));
      var f = Se.get(r, [
        "p:nvPicPr",
        "p:nvPr",
        "a:audioFile",
        "attrs",
        "r:link",
      ]);
      if (f) {
        var d =
          (s = u.ctx.rels[f]) === null || s === void 0 ? void 0 : s.target;
        u.audioFile = u.pptx.getMedia(d);
      }
      var h = Se.get(r, [
        "p:nvPicPr",
        "p:nvPr",
        "a:videoFile",
        "attrs",
        "r:link",
      ]);
      if (h) {
        var p =
          (l = u.ctx.rels[h]) === null || l === void 0 ? void 0 : l.target;
        u.videoFile = u.pptx.getMedia(p);
      }
      return u;
    }
    return (
      _0(t, e),
      Object.defineProperty(t.prototype, "base64", {
        get: function () {
          return this.pptx.getMedia(this.path);
        },
        enumerable: !1,
        configurable: !0,
      }),
      t
    );
  })(U0),
  Eee = (function (e) {
    function t(n, r, a, i) {
      var o = e.call(this, n, a, i) || this;
      return (
        (o.userDrawn = !0),
        (o.props = { tableStyleId: "" }),
        (o.tableGrid = { gridCol: [] }),
        (o.tr = []),
        (o.tableStyles = {}),
        (o.pptx = r),
        o._parseTableProps(),
        o._parseTableGrid(),
        o._parseTr(),
        o._parseInheritStyles(),
        o
      );
    }
    return (
      _0(t, e),
      Object.defineProperty(t.prototype, "slideMaster", {
        get: function () {
          return this.ctx.slideMaster || this.ctx;
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(t.prototype, "theme", {
        get: function () {
          return this.slideMaster.theme;
        },
        enumerable: !1,
        configurable: !0,
      }),
      (t.prototype._parseTableProps = function () {
        var n = Se.get(this.source, [
          "a:graphic",
          "a:graphicData",
          "a:tbl",
          "a:tblPr",
        ]);
        ((this.props.tableStyleId = Se.get(n, "a:tableStyleId")),
          (this.tableStyles =
            Se.get(this.slideMaster.tableStyles, this.props.tableStyleId) ||
            {}),
          Se.get(n, ["attrs", "bandCol"]) === "1" && (this.props.bandCol = !0),
          Se.get(n, ["attrs", "bandRow"]) === "1" && (this.props.bandRow = !0),
          Se.get(n, ["attrs", "firstCol"]) === "1" &&
            (this.props.firstCol = !0),
          Se.get(n, ["attrs", "firstRow"]) === "1" &&
            (this.props.firstRow = !0),
          Se.get(n, ["attrs", "lastCol"]) === "1" && (this.props.lastCol = !0),
          Se.get(n, ["attrs", "lastRow"]) === "1" && (this.props.lastRow = !0));
      }),
      (t.prototype._parseTableGrid = function () {
        var n = Se.get(this.source, [
          "a:graphic",
          "a:graphicData",
          "a:tbl",
          "a:tblGrid",
          "a:gridCol",
        ]);
        if (n)
          for (var r = 0; r < n.length; r++) {
            var a = Se.get(n[r], ["attrs", "w"]);
            this.tableGrid.gridCol.push({ width: or(parseInt(a)) });
          }
      }),
      (t.prototype._parseTr = function () {
        var n = [],
          r = Se.get(this.source, [
            "a:graphic",
            "a:graphicData",
            "a:tbl",
            "a:tr",
          ]);
        Array.isArray(r) || (r = [r]);
        for (var a = 0; a < r.length; a++) {
          var i = { props: {}, td: [] },
            o = r[a];
          i.props.height = or(parseInt(Se.get(o, ["attrs", "h"])));
          var s = Se.get(o, ["a:tc"]);
          Array.isArray(s) || (s = [s]);
          for (var l = 0; l < s.length; l++) i.td.push(this._parseTd(s[l]));
          n.push(i);
        }
        this.tr = n;
      }),
      (t.prototype._parseTd = function (n) {
        var r,
          a,
          i,
          o,
          s = this,
          l = { props: { border: {} }, paragraphs: [] },
          u = Se.get(n, ["a:tcPr", "attrs"]);
        (u != null &&
          u.marB &&
          (l.props.marB = or(parseInt(u == null ? void 0 : u.marB))),
          u != null &&
            u.marT &&
            (l.props.marT = or(parseInt(u == null ? void 0 : u.marT))),
          u != null &&
            u.marL &&
            (l.props.marL = or(parseInt(u == null ? void 0 : u.marL))),
          u != null &&
            u.marR &&
            (l.props.marR = or(parseInt(u == null ? void 0 : u.marR))),
          u != null &&
            u.anchor &&
            (l.props.anchor = u == null ? void 0 : u.anchor));
        var c = Se.get(n, ["a:tcPr"]);
        (Se.get(c, ["a:lnR"]) &&
          (l.props.border.right = Av(
            Se.get(c, ["a:lnR"]),
            this.theme,
            this.ctx,
          )),
          Se.get(c, ["a:lnL"]) &&
            (l.props.border.left = Av(
              Se.get(c, ["a:lnL"]),
              this.theme,
              this.ctx,
            )),
          Se.get(c, ["a:lnT"]) &&
            (l.props.border.top = Av(
              Se.get(c, ["a:lnT"]),
              this.theme,
              this.ctx,
            )),
          Se.get(c, ["a:lnB"]) &&
            (l.props.border.bottom = Av(
              Se.get(c, ["a:lnB"]),
              this.theme,
              this.ctx,
            )),
          !((r = n == null ? void 0 : n.attrs) === null || r === void 0) &&
            r.rowSpan &&
            (l.props.rowSpan = parseInt(n.attrs.rowSpan)),
          !((a = n == null ? void 0 : n.attrs) === null || a === void 0) &&
            a.gridSpan &&
            (l.props.gridSpan = parseInt(n.attrs.gridSpan)),
          !((i = n == null ? void 0 : n.attrs) === null || i === void 0) &&
            i.vMerge &&
            (l.props.vMerge = n.attrs.vMerge === "1"),
          !((o = n == null ? void 0 : n.attrs) === null || o === void 0) &&
            o.hMerge &&
            (l.props.hMerge = n.attrs.hMerge === "1"));
        var f = Se.get(n, ["a:tcPr", "a:solidFill"]);
        f && (l.props.background = $r(f, this.theme, this.ctx));
        var d = Se.get(n, ["a:txBody"]),
          h = Se.get(d, ["a:p"]);
        return (
          Array.isArray(h) || (h = [h]),
          (l.paragraphs = h.map(function (p) {
            return s._parseParagraph(p);
          })),
          l
        );
      }),
      (t.prototype._parseParagraph = function (n) {
        var r = this,
          a = {
            props: {},
            inheritProps: {},
            inheritRProps: {},
            endParaRProps: {},
            rows: [],
          },
          i = Se.get(n, ["a:pPr"]) || {};
        a.props = this._formatPPr(i);
        var o = Se.get(n, ["a:endParaRPr"]);
        a.endParaRProps = this._formatRPr(o);
        var s = Se.get(n, ["a:r"]) || [];
        Array.isArray(s) || (s = [s]);
        var l = Se.get(n, ["a:br"]) || [];
        return (
          Array.isArray(l) || (l = [l]),
          (s = s.concat(
            l.map(function (u) {
              return Tt({ isBr: !0 }, u);
            }),
          )).sort(function (u, c) {
            return (
              Se.get(u, ["attrs", "order"]) - Se.get(c, ["attrs", "order"])
            );
          }),
          (a.rows = s.map(function (u) {
            return r._parseRow(u);
          })),
          a
        );
      }),
      (t.prototype._parseRow = function (n) {
        if (n.isBr) return { isBr: !0 };
        var r = { props: {}, text: "" },
          a = Se.get(n, ["a:rPr"]) || {};
        return (
          (r.props = this._formatRPr(a)),
          (r.text = Se.get(n, "a:t") || ""),
          r
        );
      }),
      (t.prototype._formatPPr = function (n) {
        var r = {},
          a = Se.get(n, "attrs") || {};
        return (
          Object.keys(a).forEach(function (i) {
            i === "algn" && (r.align = a[i]);
          }),
          Se.get(n, ["a:lnSpc", "a:spcPct", "attrs", "val"]) &&
            (r.lineHeight = parseInt(n["a:lnSpc"]["a:spcPct"].attrs.val) / 1e5),
          Se.get(n, ["a:buAutoNum", "attrs", "type"]) &&
            (r.buAutoNum = n["a:buAutoNum"].attrs.type),
          Se.get(n, ["a:buChar", "attrs", "char"]) &&
            (r.buChar = n["a:buChar"].attrs.char),
          Se.get(n, ["a:spcBef", "a:spcPts", "attrs", "val"]) &&
            (r.spaceBefore = CS(parseInt(n["a:spcBef"]["a:spcPts"].attrs.val))),
          Se.get(n, ["a:spcAft", "a:spcPts", "attrs", "val"]) &&
            (r.spaceAfter = CS(parseInt(n["a:spcAft"]["a:spcPts"].attrs.val))),
          r
        );
      }),
      (t.prototype._formatRPr = function (n) {
        var r = {},
          a = Se.get(n, "attrs") || {};
        Object.keys(a).forEach(function (s) {
          switch (s) {
            case "sz":
              r.size = parseInt(a[s]) / 100;
              break;
            case "b":
              r.bold = a[s] === "1";
              break;
            case "i":
              r.italic = a[s] === "1";
              break;
            case "u":
              r.underline = a[s];
              break;
            case "strike":
              r.strike = a[s];
              break;
            case "order":
            case "dirty":
              break;
            default:
              r[s] = a[s];
          }
        });
        var i = Se.get(n, "a:solidFill");
        i && (r.color = $r(i, this.theme, this.ctx));
        var hl = Se.get(n, "a:hlinkClick"),
          th = this.theme;
        hl &&
          ((r.color = {
            type: "solidFill",
            color: (th && th.clrScheme && th.clrScheme.hlink) || "#0563C1",
            alpha: 1,
          }),
          (r.underline = "sng"));
        var o = Se.get(n, "a:highlight");
        return (
          o && (r.background = $r(o, this.theme, this.ctx)),
          (i = Se.get(n, ["a:latin", "attrs", "typeface"])) &&
            (r.typeface = resolveTypeface(i, this.theme)),
          (i = Se.get(n, ["a:ea", "attrs", "typeface"])) &&
            (r.typefaceEa = resolveTypeface(i, this.theme)),
          r
        );
      }),
      (t.prototype._isLastCol = function (n, r) {
        var a, i;
        if (r === n.length - 1) return !0;
        for (var o = r + 1; o < n.length; o++)
          if (
            !(!((a = n[o].props) === null || a === void 0) && a.hMerge) &&
            !(!((i = n[o].props) === null || i === void 0) && i.vMerge)
          )
            return !1;
        return !0;
      }),
      (t.prototype._isBandRow = function (n) {
        var r;
        return !((r = this.props) === null || r === void 0) && r.firstRow
          ? n % 2 == 1
          : n % 2 == 0;
      }),
      (t.prototype._isBandCol = function (n) {
        var r;
        return !((r = this.props) === null || r === void 0) && r.firstCol
          ? n % 2 == 1
          : n % 2 == 0;
      }),
      (t.prototype._parseInheritStyles = function () {
        var n,
          r,
          a,
          i,
          o,
          s,
          l,
          u,
          c = this,
          f =
            (r =
              (n = this.tableStyles) === null || n === void 0
                ? void 0
                : n.wholeTbl) === null || r === void 0
              ? void 0
              : r.tcStyle,
          d =
            (i =
              (a = this.tableStyles) === null || a === void 0
                ? void 0
                : a.wholeTbl) === null || i === void 0
              ? void 0
              : i.tcTxStyle,
          h =
            (s =
              (o = this.slideMaster.defaultTextStyle) === null || o === void 0
                ? void 0
                : o.lvl1) === null || s === void 0
              ? void 0
              : s.props,
          p =
            (u =
              (l = this.slideMaster.defaultTextStyle) === null || l === void 0
                ? void 0
                : l.lvl1) === null || u === void 0
              ? void 0
              : u.defRPr;
        this.tr.forEach(function (g, y) {
          g.td.forEach(function (m, v) {
            var b,
              _,
              S,
              C,
              T,
              k,
              M,
              I,
              L,
              E,
              R,
              D,
              O,
              B,
              V,
              z,
              N,
              H,
              G,
              j,
              J,
              re,
              le,
              Te,
              fe,
              _e,
              ye,
              Oe,
              Ae,
              pe,
              te,
              F,
              ie,
              ce,
              Z,
              q,
              ae,
              we,
              xe,
              he,
              Pe,
              Le,
              be,
              We,
              ut,
              rt,
              kt,
              un,
              an,
              De,
              Ue,
              ge,
              Fe,
              Be,
              yt,
              sn,
              Me = Tt(Tt({}, h), f),
              ct = Tt(Tt({}, p), d);
            (c.props.firstRow && y === 0
              ? ((Me = Tt(
                  Tt(
                    Tt({}, Me),
                    (_ =
                      (b = c.tableStyles) === null || b === void 0
                        ? void 0
                        : b.firstRow) === null || _ === void 0
                      ? void 0
                      : _.tcStyle,
                  ),
                  {
                    border: Tt(
                      Tt({}, Me == null ? void 0 : Me.border),
                      (T =
                        (C =
                          (S = c.tableStyles) === null || S === void 0
                            ? void 0
                            : S.firstRow) === null || C === void 0
                          ? void 0
                          : C.tcStyle) === null || T === void 0
                        ? void 0
                        : T.border,
                    ),
                  },
                )),
                (ct = Tt(
                  Tt({}, ct),
                  (M =
                    (k = c.tableStyles) === null || k === void 0
                      ? void 0
                      : k.firstRow) === null || M === void 0
                    ? void 0
                    : M.tcTxStyle,
                )))
              : c.props.lastRow && y === c.tr.length - 1
                ? ((Me = Tt(
                    Tt(
                      Tt({}, Me),
                      (L =
                        (I = c.tableStyles) === null || I === void 0
                          ? void 0
                          : I.lastRow) === null || L === void 0
                        ? void 0
                        : L.tcStyle,
                    ),
                    {
                      border: Tt(
                        Tt({}, Me == null ? void 0 : Me.border),
                        (D =
                          (R =
                            (E = c.tableStyles) === null || E === void 0
                              ? void 0
                              : E.lastRow) === null || R === void 0
                            ? void 0
                            : R.tcStyle) === null || D === void 0
                          ? void 0
                          : D.border,
                      ),
                    },
                  )),
                  (ct = Tt(
                    Tt({}, ct),
                    (B =
                      (O = c.tableStyles) === null || O === void 0
                        ? void 0
                        : O.lastRow) === null || B === void 0
                      ? void 0
                      : B.tcTxStyle,
                  )))
                : c.props.firstCol && v === 0
                  ? ((Me = Tt(
                      Tt(
                        Tt({}, Me),
                        (z =
                          (V = c.tableStyles) === null || V === void 0
                            ? void 0
                            : V.firstCol) === null || z === void 0
                          ? void 0
                          : z.tcStyle,
                      ),
                      {
                        border: Tt(
                          Tt({}, Me == null ? void 0 : Me.border),
                          (G =
                            (H =
                              (N = c.tableStyles) === null || N === void 0
                                ? void 0
                                : N.firstCol) === null || H === void 0
                              ? void 0
                              : H.tcStyle) === null || G === void 0
                            ? void 0
                            : G.border,
                        ),
                      },
                    )),
                    (ct = Tt(
                      Tt({}, ct),
                      (J =
                        (j = c.tableStyles) === null || j === void 0
                          ? void 0
                          : j.firstCol) === null || J === void 0
                        ? void 0
                        : J.tcTxStyle,
                    )))
                  : c.props.lastCol && c._isLastCol(g.td, v)
                    ? ((Me = Tt(
                        Tt(
                          Tt({}, Me),
                          (le =
                            (re = c.tableStyles) === null || re === void 0
                              ? void 0
                              : re.lastCol) === null || le === void 0
                            ? void 0
                            : le.tcStyle,
                        ),
                        {
                          border: Tt(
                            Tt({}, Me == null ? void 0 : Me.border),
                            (_e =
                              (fe =
                                (Te = c.tableStyles) === null || Te === void 0
                                  ? void 0
                                  : Te.lastCol) === null || fe === void 0
                                ? void 0
                                : fe.tcStyle) === null || _e === void 0
                              ? void 0
                              : _e.border,
                          ),
                        },
                      )),
                      (ct = Tt(
                        Tt({}, ct),
                        (Oe =
                          (ye = c.tableStyles) === null || ye === void 0
                            ? void 0
                            : ye.lastCol) === null || Oe === void 0
                          ? void 0
                          : Oe.tcTxStyle,
                      )))
                    : (c.props.bandRow &&
                        (c._isBandRow(y)
                          ? ((Me = Tt(
                              Tt(
                                Tt({}, Me),
                                (pe =
                                  (Ae = c.tableStyles) === null || Ae === void 0
                                    ? void 0
                                    : Ae.band1H) === null || pe === void 0
                                  ? void 0
                                  : pe.tcStyle,
                              ),
                              {
                                border: Tt(
                                  Tt({}, Me == null ? void 0 : Me.border),
                                  (ie =
                                    (F =
                                      (te = c.tableStyles) === null ||
                                      te === void 0
                                        ? void 0
                                        : te.band1H) === null || F === void 0
                                      ? void 0
                                      : F.tcStyle) === null || ie === void 0
                                    ? void 0
                                    : ie.border,
                                ),
                              },
                            )),
                            (ct = Tt(
                              Tt({}, ct),
                              (Z =
                                (ce = c.tableStyles) === null || ce === void 0
                                  ? void 0
                                  : ce.band1H) === null || Z === void 0
                                ? void 0
                                : Z.tcTxStyle,
                            )))
                          : ((Me = Tt(
                              Tt(
                                Tt({}, Me),
                                (ae =
                                  (q = c.tableStyles) === null || q === void 0
                                    ? void 0
                                    : q.band2V) === null || ae === void 0
                                  ? void 0
                                  : ae.tcStyle,
                              ),
                              {
                                border: Tt(
                                  Tt({}, Me == null ? void 0 : Me.border),
                                  (he =
                                    (xe =
                                      (we = c.tableStyles) === null ||
                                      we === void 0
                                        ? void 0
                                        : we.band2V) === null || xe === void 0
                                      ? void 0
                                      : xe.tcStyle) === null || he === void 0
                                    ? void 0
                                    : he.border,
                                ),
                              },
                            )),
                            (ct = Tt(
                              Tt({}, ct),
                              (Le =
                                (Pe = c.tableStyles) === null || Pe === void 0
                                  ? void 0
                                  : Pe.band2V) === null || Le === void 0
                                ? void 0
                                : Le.tcTxStyle,
                            )))),
                      c.props.bandCol &&
                        (c._isBandCol(v)
                          ? ((Me = Tt(
                              Tt(
                                Tt({}, Me),
                                (We =
                                  (be = c.tableStyles) === null || be === void 0
                                    ? void 0
                                    : be.band1V) === null || We === void 0
                                  ? void 0
                                  : We.tcStyle,
                              ),
                              {
                                border: Tt(
                                  Tt({}, Me == null ? void 0 : Me.border),
                                  (kt =
                                    (rt =
                                      (ut = c.tableStyles) === null ||
                                      ut === void 0
                                        ? void 0
                                        : ut.band1V) === null || rt === void 0
                                      ? void 0
                                      : rt.tcStyle) === null || kt === void 0
                                    ? void 0
                                    : kt.border,
                                ),
                              },
                            )),
                            (ct = Tt(
                              Tt({}, ct),
                              (an =
                                (un = c.tableStyles) === null || un === void 0
                                  ? void 0
                                  : un.band1V) === null || an === void 0
                                ? void 0
                                : an.tcTxStyle,
                            )))
                          : ((Me = Tt(
                              Tt(
                                Tt({}, Me),
                                (Ue =
                                  (De = c.tableStyles) === null || De === void 0
                                    ? void 0
                                    : De.band2H) === null || Ue === void 0
                                  ? void 0
                                  : Ue.tcStyle,
                              ),
                              {
                                border: Tt(
                                  Tt({}, Me == null ? void 0 : Me.border),
                                  (Be =
                                    (Fe =
                                      (ge = c.tableStyles) === null ||
                                      ge === void 0
                                        ? void 0
                                        : ge.band2H) === null || Fe === void 0
                                      ? void 0
                                      : Fe.tcStyle) === null || Be === void 0
                                    ? void 0
                                    : Be.border,
                                ),
                              },
                            )),
                            (ct = Tt(
                              Tt({}, ct),
                              (sn =
                                (yt = c.tableStyles) === null || yt === void 0
                                  ? void 0
                                  : yt.band2H) === null || sn === void 0
                                ? void 0
                                : sn.tcTxStyle,
                            ))))),
              (m.inheritTcStyle = Me),
              (m.inheritTcTxStyle = ct));
          });
        });
      }),
      t
    );
  })(U0),
  uO = (function () {
    function e(t, n, r, a) {
      if (
        ((this.offset = { x: 0, y: 0 }),
        (this.chOffset = { x: 0, y: 0 }),
        (this.extend = { w: 0, h: 0 }),
        (this.chExtend = { w: 0, h: 0 }),
        (this.rotate = 0),
        (this.nodes = []),
        (this.flipV = !1),
        (this.flipH = !1),
        (this.userDrawn = !0),
        (this.order = Se.get(t, ["attrs", "order"])),
        (this.pptx = n),
        (this.ctx = r),
        (this.source = t),
        (this.group = a),
        this.source["p:grpSpPr"])
      ) {
        var i = Se.get(this.source, ["p:grpSpPr", "a:xfrm"]);
        if (
          (i &&
            ((this.offset = {
              x: Math.round(
                this.group && this.pptx.wps
                  ? el(parseInt(i["a:off"].attrs.x))
                  : or(parseInt(i["a:off"].attrs.x)),
              ),
              y: Math.round(
                this.group && this.pptx.wps
                  ? el(parseInt(i["a:off"].attrs.y))
                  : or(parseInt(i["a:off"].attrs.y)),
              ),
            }),
            (this.chOffset = {
              x: Math.round(
                this.pptx.wps || !isEmuXfrm(i)
                  ? el(parseInt(i["a:chOff"].attrs.x))
                  : or(parseInt(i["a:chOff"].attrs.x)),
              ),
              y: Math.round(
                this.pptx.wps || !isEmuXfrm(i)
                  ? el(parseInt(i["a:chOff"].attrs.y))
                  : or(parseInt(i["a:chOff"].attrs.y)),
              ),
            }),
            (this.chExtend = {
              w: Math.round(
                this.pptx.wps || !isEmuXfrm(i)
                  ? el(parseInt(i["a:chExt"].attrs.cx))
                  : or(parseInt(i["a:chExt"].attrs.cx)),
              ),
              h: Math.round(
                this.pptx.wps || !isEmuXfrm(i)
                  ? el(parseInt(i["a:chExt"].attrs.cy))
                  : or(parseInt(i["a:chExt"].attrs.cy)),
              ),
            }),
            (this.extend = {
              w: Math.round(
                this.group && this.pptx.wps
                  ? el(parseInt(i["a:ext"].attrs.cx))
                  : or(parseInt(i["a:ext"].attrs.cx)),
              ),
              h: Math.round(
                this.group && this.pptx.wps
                  ? el(parseInt(i["a:ext"].attrs.cy))
                  : or(parseInt(i["a:ext"].attrs.cy)),
              ),
            }),
            (this.rotate = t2(parseInt(Se.get(i, "attrs.rot", 0)))),
            (this.flipV = Se.get(i, "attrs.flipV") === "1"),
            (this.flipH = Se.get(i, "attrs.flipH") === "1")),
          a)
        ) {
          var o = a.extend,
            s = a.chExtend,
            l = a.chOffset,
            u = o.w / s.w,
            c = o.h / s.h;
          ((this.extend.w = this.extend.w * u),
            (this.extend.h = this.extend.h * c),
            (this.offset.x = (this.offset.x - l.x) * u),
            (this.offset.y = (this.offset.y - l.y) * c));
        }
      }
      (this._parseBackground(), this._parseNodes());
    }
    return (
      (e.prototype.getBackground = function () {
        return this.background && this.background.type !== "none"
          ? this.background
          : this.group
            ? this.group.getBackground()
            : void 0;
      }),
      (e.prototype._parseBackground = function () {
        var t = Se.get(this.source, ["p:grpSpPr"]);
        t && t["a:solidFill"]
          ? (this.background = $r(t["a:solidFill"], this.ctx.theme, this.ctx))
          : t && t["a:gradFill"]
            ? (this.background = G0(t["a:gradFill"], this.ctx.theme, this.ctx))
            : t &&
              t["a:blipFill"] &&
              (this.background = W0(t["a:blipFill"], this.pptx, this.ctx));
      }),
      (e.prototype._parseNodes = function () {
        var t = this.source;
        K0(this.nodes, t, this.pptx, this.ctx, this);
      }),
      e
    );
  })(),
  Lee = (function (e) {
    function t(n, r, a, i) {
      var o = e.call(this, n, a, i) || this;
      return ((o.nodes = []), (o.pptx = r), o);
    }
    return (
      _0(t, e),
      (t.prototype.parseNode = function () {
        return oa(this, void 0, void 0, function () {
          var n, r, a, i, o, s, l, u, c;
          return sa(this, function (f) {
            switch (f.label) {
              case 0:
                return (
                  f.trys.push([0, 3, , 4]),
                  (n = Se.get(this.source, [
                    "a:graphic",
                    "a:graphicData",
                    "dgm:relIds",
                    "attrs",
                    "r:dm",
                  ])) && this.ctx.rels[n]
                    ? ((r = this.ctx.rels[n].target),
                      (i = $i),
                      [4, this.pptx.getXmlByPath(r)])
                    : [2]
                );
              case 1:
                return (
                  (a = i.apply(void 0, [f.sent()])),
                  (o = Se.get(a, [
                    "dgm:dataModel",
                    "dgm:extLst",
                    "a:ext",
                    "dsp:dataModelExt",
                    "attrs",
                    "relId",
                  ])) && this.ctx.rels[o]
                    ? ((s = this.ctx.rels[o].target),
                      [4, this.pptx.getXmlByPath(s)])
                    : [2]
                );
              case 2:
                return (
                  (l = (l = f.sent()).replace(/dsp:/g, "p:")),
                  (u = $i(l)),
                  (c = Se.get(u, ["p:drawing", "p:spTree"])),
                  K0(this.nodes, c, this.pptx, this.ctx),
                  [3, 4]
                );
              case 3:
                return (f.sent(), [3, 4]);
              case 4:
                return [2];
            }
          });
        });
      }),
      t
    );
  })(U0),
  Dee = (function (e) {
    function t(n, r, a, i) {
      var o = e.call(this, n, a, i) || this;
      return (
        (o.options = {
          // no colour list until the file names one: echarts reads an empty
          // list as "no palette" and then draws the axes without any series
          title: {},
          tooltip: {},
          legend: {},
          series: [],
          // echarts' default margins (60px a side) eat a small chart alive; the
          // frame the slide gives us is the chart, so keep the plot inside it
          grid: { left: 6, right: 10, top: 28, bottom: 4, containLabel: true },
          // no entry animation: the card copies a slide moments after it is
          // rendered, and a copy of a half drawn chart is a chart of flat bars
          animation: false,
        }),
        (o.userDrawn = !0),
        (o.pptx = r),
        o
      );
    }
    return (
      _0(t, e),
      Object.defineProperty(t.prototype, "slideMaster", {
        get: function () {
          return this.ctx.slideMaster || this.ctx;
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(t.prototype, "theme", {
        get: function () {
          return this.slideMaster.theme;
        },
        enumerable: !1,
        configurable: !0,
      }),
      (t.prototype.parseNode = function () {
        return oa(this, void 0, void 0, function () {
          var n, r, a, i, o, s;
          return sa(this, function (l) {
            switch (l.label) {
              case 0:
                return (
                  l.trys.push([0, 2, , 3]),
                  (n = Se.get(this.source, [
                    "a:graphic",
                    "a:graphicData",
                    "c:chart",
                    "attrs",
                    "r:id",
                  ])) && this.ctx.rels[n]
                    ? ((r = this.ctx.rels[n].target),
                      (i = $i),
                      [4, this.pptx.getXmlByPath(r)])
                    : [2]
                );
              case 1:
                return (
                  (a = i.apply(void 0, [l.sent()])),
                  (o = Se.get(a, ["c:chartSpace", "c:chart"])),
                  (s = Se.get(o, ["c:plotArea"])),
                  Se.get(s, ["c:lineChart"])
                    ? this.parseLine(Se.get(s, ["c:lineChart"]), o)
                    : Se.get(s, ["c:line3DChart"])
                      ? this.parseLine(Se.get(s, ["c:line3DChart"]), o)
                      : Se.get(s, ["c:areaChart"])
                        ? this.parseAreaLine(Se.get(s, ["c:areaChart"]), o)
                        : Se.get(s, ["c:area3DChart"])
                          ? this.parseAreaLine(Se.get(s, ["c:area3DChart"]), o)
                          : Se.get(s, ["c:barChart"])
                            ? this.parseBar(Se.get(s, ["c:barChart"]), o)
                            : Se.get(s, ["c:bar3DChart"])
                              ? this.parseBar(Se.get(s, ["c:bar3DChart"]), o)
                              : Se.get(s, ["c:pieChart"])
                                ? this.parsePie(Se.get(s, ["c:pieChart"]))
                                : Se.get(s, ["c:pie3DChart"])
                                  ? this.parsePie(Se.get(s, ["c:pie3DChart"]))
                                  : Se.get(s, ["c:doughnutChart"]) &&
                                    this.parseDoughnutChart(
                                      Se.get(s, ["c:doughnutChart"]),
                                    ),
                  [3, 3]
                );
              case 2:
                return (l.sent(), [3, 3]);
              case 3:
                return [2];
            }
          });
        });
      }),
      (t.prototype.parseAreaLine = function (n, r) {
        (this.parseLine(n, r),
          (this.options.series = this.options.series.map(function (a) {
            return ((a.areaStyle = {}), a);
          })));
      }),
      (t.prototype.parseLine = function (n, r) {
        var a = Se.get(n, ["c:ser"]);
        (Array.isArray(a) || (a = [a]),
          (this.options.title = {
            top: "top",
            left: "center",
            text:
              this.parseChartTitle(Se.get(r, ["c:title"])) ||
              Se.get(a[0], ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) ||
              "",
          }),
          (this.options.xAxis = {
            type: "category",
            data: this.getCategory(a[0]),
          }),
          (this.options.yAxis = { type: "value" }),
          (this.options.series = this.parseLineSeries(a, n)),
          this._palette(this.parseLineColors(a)),
          (this.options.legend = (function (legend) {
            if (!legend) return { show: false };
            var pos = Se.get(legend, ["c:legendPos", "attrs", "val"]) || "r";
            return pos === "t" || pos === "tr"
              ? { top: "top", left: pos === "tr" ? "right" : "center" }
              : { bottom: "bottom", left: "center" };
          })(Se.get(r, ["c:legend"]))),
          Se.get(n, ["c:grouping", "attrs", "val"]) === "percentStacked" &&
            (this.options.tooltip.valueFormatter = function (i) {
              return (100 * i).toFixed(2) + "%";
            }));
      }),
      (t.prototype.parseBar = function (n, r) {
        var a = Se.get(n, ["c:ser"]);
        (Array.isArray(a) || (a = [a]),
          (this.options.title = {
            top: "top",
            left: "center",
            text: this.parseChartTitle(Se.get(r, ["c:title"])),
          }),
          Se.get(n, ["c:barDir", "attrs", "val"]) === "bar"
            ? ((this.options.yAxis = {
                type: "category",
                data: this.getCategory(a[0]),
              }),
              (this.options.xAxis = { type: "value" }))
            : ((this.options.xAxis = {
                type: "category",
                data: this.getCategory(a[0]),
              }),
              (this.options.yAxis = { type: "value" })),
          (this.options.series = this.parseBarSeries(a, n)),
          this._palette(this.parseBarColors(a)),
          (this.options.legend = (function (legend) {
            if (!legend) return { show: false };
            var pos = Se.get(legend, ["c:legendPos", "attrs", "val"]) || "r";
            return pos === "t" || pos === "tr"
              ? { top: "top", left: pos === "tr" ? "right" : "center" }
              : { bottom: "bottom", left: "center" };
          })(Se.get(r, ["c:legend"]))),
          Se.get(n, ["c:grouping", "attrs", "val"]) === "percentStacked" &&
            (this.options.tooltip.valueFormatter = function (i) {
              return (100 * i).toFixed(2) + "%";
            }));
      }),
      (t.prototype.parsePie = function (n) {
        var r = Se.get(n, ["c:ser"]);
        ((this.options.title = {
          top: "top",
          left: "center",
          text: this.parsePieTitle(r),
        }),
          this._palette(this.parsePieColors(r)),
          (this.options.series = [this.parsePieSeries(r, n)]),
          (this.options.legend = { bottom: "bottom", left: "center" }));
      }),
      (t.prototype.parseDoughnutChart = function (n) {
        var r = Se.get(n, ["c:ser"]);
        ((this.options.title.text = this.parsePieTitle(r)),
          this._palette(this.parsePieColors(r)),
          (this.options.series = [this.parsePieSeries(r, n)]),
          (this.options.legend = { bottom: "bottom", left: "center" }));
      }),
      (t.prototype.parsePieTitle = function (n) {
        return Se.get(n, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]);
      }),
      // echarts reads an empty colour list as "no palette" and then draws the
      // axes without any series, so only hand it colours the file names
      (t.prototype._palette = function (list) {
        list && list.length && (this.options.color = list);
      }),
      (t.prototype.parseChartTitle = function (n) {
        var r = Se.get(n, ["c:tx", "c:rich", "a:p"]);
        return (
          Array.isArray(r) || (r = [r]),
          r
            .map(function (a) {
              var i = Se.get(a, ["a:r"]);
              return (
                Array.isArray(i) || (i = [i]),
                i
                  .map(function (o) {
                    return Se.get(o, ["a:t"]) || "";
                  })
                  .join("")
              );
            })
            .join("")
        );
      }),
      (t.prototype.parseBarColors = function (n) {
        var r = this;
        return n
          .map(function (a) {
            return Qn($r(Se.get(a, ["c:spPr", "a:solidFill"]), r.theme, r.ctx));
          })
          .filter(Boolean);
      }),
      (t.prototype.parseLineColors = function (n) {
        var r = this;
        return n.map(function (a) {
          return Qn(
            $r(
              Se.get(a, ["c:spPr", "a:ln", "a:solidFill"]) ||
                Se.get(a, ["c:spPr", "a:solidFill"]),
              r.theme,
              r.ctx,
            ),
          );
        }).filter(Boolean);
      }),
      (t.prototype.parsePieColors = function (n) {
        var r = this,
          a = [],
          i = Se.get(n, ["c:dPt"]);
        return (
          Array.isArray(i) || (i = [i]),
          i.forEach(function (o) {
            a.push(
              Qn($r(Se.get(o, ["c:spPr", "a:solidFill"]), r.theme, r.ctx)),
            );
          }),
          a
        );
      }),
      (t.prototype.parsePieSeries = function (n, r) {
        var a = { type: "pie", radius: "80%", startAngle: 90, data: [] };
        Se.get(r, ["c:holeSize", "attrs", "val"]) &&
          (a.radius = [
            "".concat(0.8 * Se.get(r, ["c:holeSize", "attrs", "val"]), "%"),
            "80%",
          ]);
        var i = Se.get(r, ["c:firstSliceAng", "attrs", "val"]);
        i && (a.startAngle = 90 - i);
        for (
          var o = this.getCategory(n), s = this.getVal(n), l = 0;
          l < o.length;
          l++
        )
          a.data.push({ name: o[l], value: s[l] });
        return a;
      }),
      (t.prototype.parseBarSeries = function (n, r) {
        var a,
          i = this,
          o = Se.get(r, ["c:grouping", "attrs", "val"]);
        o === "stacked" ? (a = "Ad") : o === "percentStacked" && (a = "total");
        var showVal =
            Se.get(r, ["c:dLbls", "c:showVal", "attrs", "val"]) === "1" ||
            Se.get(n[0], ["c:dLbls", "c:showVal", "attrs", "val"]) === "1";
        var s = n.map(function (u) {
          return {
            type: "bar",
            name: Se.get(u, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]),
            data: i.getVal(u),
            stack: a,
            label: showVal ? { show: true, position: "top" } : void 0,
          };
        });
        if (o === "percentStacked") {
          var l = [];
          (s.forEach(function (u, c) {
            l =
              c === 0
                ? Lm([], u.data, !0)
                : l.map(function (f, d) {
                    return f + u.data[d];
                  });
          }),
            s.forEach(function (u) {
              u.data = u.data.map(function (c, f) {
                return l[f] <= 0 ? 0 : c / l[f];
              });
            }));
        }
        return s;
      }),
      (t.prototype.parseLineSeries = function (n, r) {
        var a,
          i = this,
          o = Se.get(r, ["c:grouping", "attrs", "val"]);
        o === "stacked" ? (a = "Ad") : o === "percentStacked" && (a = "total");
        var s = n.map(function (u) {
          return {
            type: "line",
            name: Se.get(u, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]),
            data: i.getVal(u),
            stack: a,
          };
        });
        if (a === "total") {
          var l = [];
          (s.forEach(function (u, c) {
            l =
              c === 0
                ? Lm([], u.data, !0)
                : l.map(function (f, d) {
                    return f + u.data[d];
                  });
          }),
            s.forEach(function (u) {
              u.data = u.data.map(function (c, f) {
                return l[f] <= 0 ? 0 : c / l[f];
              });
            }));
        }
        return s;
      }),
      (t.prototype.getCategory = function (n) {
        if (Se.get(n, ["c:cat", "c:strRef"])) {
          var r = Se.get(n, ["c:cat", "c:strRef", "c:strCache", "c:pt"]);
          return (
            Array.isArray(r) || (r = [r]),
            r.map(function (a) {
              return Se.get(a, ["c:v"]);
            })
          );
        }
        if (Se.get(n, ["c:cat", "c:numRef"]))
          return (
            (r = Se.get(n, ["c:cat", "c:numRef", "c:numCache", "c:pt"])),
            Array.isArray(r) || (r = [r]),
            r.map(function (a) {
              return Se.get(a, ["c:v"]);
            })
          );
      }),
      (t.prototype.getVal = function (n) {
        var r = Se.get(n, ["c:val", "c:numRef", "c:numCache", "c:pt"]);
        return (
          Array.isArray(r) || (r = [r]),
          r.map(function (a) {
            return +Se.get(a, ["c:v"]);
          })
        );
      }),
      t
    );
  })(U0);
function fW(e) {
  var t = {},
    n = Se.get(e, "attrs") || {};
  return (
    Object.keys(n).forEach(function (r) {
      switch (r) {
        case "algn":
          t.align = n[r];
          break;
        case "marL":
          t.marginLeft = or(parseInt(n[r]));
          break;
        case "indent":
          t.indent = or(parseInt(n[r]));
          break;
        case "lvl":
          t.level = n[r];
      }
    }),
    Se.get(e, ["a:lnSpc", "a:spcPct", "attrs", "val"]) &&
      (t.lineHeight = parseInt(e["a:lnSpc"]["a:spcPct"].attrs.val) / 1e5),
    t
  );
}
function dW(e, t, n) {
  var r = {},
    a = Se.get(e, "attrs") || {};
  Object.keys(a).forEach(function (o) {
    switch (o) {
      case "sz":
        r.size = parseInt(a[o]) / 100;
        break;
      case "b":
        r.bold = a[o] === "1";
        break;
      case "i":
        r.italic = a[o] === "1";
        break;
      case "u":
        r.underline = a[o];
        break;
      case "strike":
        r.strike = a[o];
        break;
      case "order":
      case "dirty":
        break;
      default:
        r[o] = a[o];
    }
  });
  var i = Se.get(e, "a:solidFill"),
    hl = Se.get(e, "a:hlinkClick"),
    o = Se.get(e, ["a:latin", "attrs", "typeface"]);
  return (
    i && (r.color = $r(i, t, n)),
    // hyperlink runs take the theme hlink colour and an underline (upstream: neither)
    hl &&
      ((r.color = { type: "solidFill", color: (t && t.clrScheme && t.clrScheme.hlink) || "#0563C1" }),
      (r.underline = "sng")),
    o && (r.typeface = resolveTypeface(o, t)),
    (o = Se.get(e, ["a:ea", "attrs", "typeface"])) &&
      (r.typefaceEa = resolveTypeface(o, t)),
    r
  );
}
function K0(e, t, n, r, a) {
  return oa(this, void 0, void 0, function () {
    var i, o, s, l, u, c, f, d, h, p, g, y, m, v, b, _, S, C, T, k, M;
    return sa(this, function (I) {
      switch (I.label) {
        case 0:
          for (s in ((o = []), (i = t))) o.push(s);
          ((l = 0), (I.label = 1));
        case 1:
          if (!(l < o.length)) return [3, 17];
          if (!((s = o[l]) in i)) return [3, 16];
          switch ((u = s)) {
            case "p:sp":
              return [3, 2];
            case "p:pic":
              return [3, 3];
            case "p:cxnSp":
              return [3, 4];
            case "p:graphicFrame":
              return [3, 5];
            case "p:grpSp":
              return [3, 14];
          }
          return [3, 15];
        case 2:
          for (
            c = Array.isArray(t[u]) ? t[u] : [t[u]], k = 0;
            k < c.length;
            k++
          )
            ((m = c[k]), e.push(new Jy(m, n, r, a)));
          return [3, 16];
        case 3:
          for (
            f = Array.isArray(t[u]) ? t[u] : [t[u]], k = 0;
            k < f.length;
            k++
          )
            ((d = f[k]),
              (h = d["p:blipFill"]["a:blip"].attrs["r:embed"]),
              (p = r.rels[h].target),
              (g = new TS(p, d, n, r, a)),
              e.push(g));
          return [3, 16];
        case 4:
          for (
            y = Array.isArray(t[u]) ? t[u] : [t[u]], k = 0;
            k < y.length;
            k++
          )
            ((m = y[k]), e.push(new Jy(m, n, r, a)));
          return [3, 16];
        case 5:
          ((v = Array.isArray(t[u]) ? t[u] : [t[u]]), (k = 0), (I.label = 6));
        case 6:
          if (!(k < v.length)) return [3, 13];
          switch (
            ((b = v[k]),
            (_ = Se.get(b, ["a:graphic", "a:graphicData", "attrs", "uri"])),
            _)
          ) {
            case "http://schemas.openxmlformats.org/drawingml/2006/table":
              return [3, 7];
            case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
              return [3, 8];
            case "http://schemas.openxmlformats.org/drawingml/2006/chart":
              return [3, 10];
          }
          return [3, 12];
        case 7:
          return (e.push(new Eee(b, n, r, a)), [3, 12]);
        case 8:
          return [4, (S = new Lee(b, n, r, a)).parseNode()];
        case 9:
          return (I.sent(), e.push(S), [3, 12]);
        case 10:
          return [4, (C = new Dee(b, n, r, a)).parseNode()];
        case 11:
          return (I.sent(), e.push(C), [3, 12]);
        case 12:
          return (k++, [3, 6]);
        case 13:
          return [3, 16];
        case 14:
          for (
            T = Array.isArray(t[u]) ? t[u] : [t[u]], k = 0;
            k < T.length;
            k++
          )
            ((M = T[k]), e.push(new uO(M, n, r, a)));
          return [3, 16];
        case 15:
          return [3, 16];
        case 16:
          return (l++, [3, 1]);
        case 17:
          return [2];
      }
    });
  });
}
var fJe = (function () {
    function e(t, n, r) {
      ((this.slideType = "slide"),
        (this.rels = {}),
        (this.background = { type: "none" }),
        (this.nodes = []),
        (this.name = t),
        (this.source = n),
        (this.pptx = r));
    }
    return (
      Object.defineProperty(e.prototype, "index", {
        get: function () {
          if (!this.name) return 0;
          var t = this.name.match(/(\d+)/);
          return t ? parseInt(t[0]) : 1;
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(e.prototype, "slideMaster", {
        get: function () {
          return this.slideLayout && this.slideLayout.slideMaster;
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(e.prototype, "theme", {
        get: function () {
          return this.slideMaster.theme;
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(e.prototype, "_relsPath", {
        get: function () {
          return (
            this.name.replace("slides/slide", "slides/_rels/slide") + ".rels"
          );
        },
        enumerable: !1,
        configurable: !0,
      }),
      (e.prototype.load = function () {
        return oa(this, void 0, void 0, function () {
          return sa(this, function (t) {
            switch (t.label) {
              case 0:
                return [4, this._loadRels()];
              case 1:
                return (
                  t.sent(),
                  this._loadBackground(),
                  [4, this._loadNodes()]
                );
              case 2:
                return (t.sent(), [2]);
            }
          });
        });
      }),
      (e.prototype._loadRels = function () {
        return oa(this, void 0, void 0, function () {
          var t,
            n,
            r,
            a = this;
          return sa(this, function (i) {
            switch (i.label) {
              case 0:
                return ((n = $i), [4, this.pptx.getXmlByPath(this._relsPath)]);
              case 1:
                return (
                  (t = n.apply(void 0, [i.sent()])),
                  (r = Se.get(t, ["Relationships", "Relationship"]) || []),
                  Array.isArray(r) || (r = [r]),
                  r.forEach(function (o) {
                    switch (Se.get(o, ["attrs", "Type"])) {
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        var s = o.attrs.Target.replace("../", "ppt/");
                        (s.startsWith("/ppt") && (s = s.substr(1)),
                          (a.slideLayout = a.pptx.getSlideLayout(s)));
                        break;
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio":
                      case "http://schemas.microsoft.com/office/2007/relationships/media":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData":
                      case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
                        var l = o.attrs.Target.replace("../", "ppt/");
                        (l.startsWith("/ppt") && (l = l.substr(1)),
                          (a.rels[o.attrs.Id] = {
                            type: o.attrs.Type.split("/").pop(),
                            target: l,
                          }));
                    }
                  }),
                  [2]
                );
            }
          });
        });
      }),
      (e.prototype._loadBackground = function () {
        var t = Se.get(this.source, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
        t && t["a:solidFill"]
          ? (this.background = $r(t["a:solidFill"], this.theme, this))
          : t && t["a:gradFill"]
            ? (this.background = G0(t["a:gradFill"], this.theme, this))
            : t &&
              t["a:blipFill"] &&
              (this.background = W0(t["a:blipFill"], this.pptx, this));
      }),
      (e.prototype._loadNodes = function () {
        return oa(this, void 0, void 0, function () {
          var t;
          return sa(this, function (n) {
            switch (n.label) {
              case 0:
                return (
                  (t = Se.get(this.source, ["p:sld", "p:cSld", "p:spTree"])),
                  [4, K0(this.nodes, t, this.pptx, this)]
                );
              case 1:
                return (n.sent(), [2]);
            }
          });
        });
      }),
      (e.prototype.getColorThemeName = function (t) {
        return this.slideLayout.getColorThemeName(t);
      }),
      (e.prototype.getNodeInheritAttrsByType = function (t, n) {
        var r = this.slideLayout.getNodeByType(t),
          a = Se.get(r.source, n);
        return a || this.slideLayout.getNodeInheritAttrsByType(t, n);
      }),
      (e.prototype.getNodeInheritAttrsByIdx = function (t, n) {
        var r = this.slideLayout.getNodeByIdx(t),
          a = Se.get(r.source, n);
        return a || this.slideLayout.getNodeInheritAttrsByIdx(t, n);
      }),
      e
    );
  })(),
  dJe = (function () {
    function e(t, n, r) {
      ((this.slideType = "slideLayout"),
        (this.rels = {}),
        (this.background = { type: "none" }),
        (this.nodes = []),
        (this.name = t),
        (this.source = n),
        (this.pptx = r));
    }
    return (
      Object.defineProperty(e.prototype, "_relsPath", {
        get: function () {
          return (
            this.name.replace(
              "slideLayouts/slideLayout",
              "slideLayouts/_rels/slideLayout",
            ) + ".rels"
          );
        },
        enumerable: !1,
        configurable: !0,
      }),
      Object.defineProperty(e.prototype, "theme", {
        get: function () {
          return this.slideMaster.theme;
        },
        enumerable: !1,
        configurable: !0,
      }),
      (e.prototype.load = function () {
        return oa(this, void 0, void 0, function () {
          return sa(this, function (t) {
            switch (t.label) {
              case 0:
                return [4, this._loadRels()];
              case 1:
                return (t.sent(), [4, this._loadBackground()]);
              case 2:
                return (t.sent(), [4, this._loadNodes()]);
              case 3:
                return (t.sent(), [2]);
            }
          });
        });
      }),
      (e.prototype._loadRels = function () {
        return oa(this, void 0, void 0, function () {
          var t,
            n,
            r,
            a = this;
          return sa(this, function (i) {
            switch (i.label) {
              case 0:
                return ((n = $i), [4, this.pptx.getXmlByPath(this._relsPath)]);
              case 1:
                return (
                  (t = n.apply(void 0, [i.sent()])),
                  (r = Se.get(t, ["Relationships", "Relationship"]) || []),
                  Array.isArray(r) || (r = [r]),
                  r.forEach(function (o) {
                    switch (Se.get(o, ["attrs", "Type"])) {
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                        var s = o.attrs.Target.replace("../", "ppt/");
                        (s.startsWith("/ppt") && (s = s.substr(1)),
                          (a.slideMaster = a.pptx.getSlideMaster(s)));
                        break;
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio":
                      case "http://schemas.microsoft.com/office/2007/relationships/media":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video":
                        var l = o.attrs.Target.replace("../", "ppt/");
                        (l.startsWith("/ppt") && (l = l.substr(1)),
                          (a.rels[o.attrs.Id] = {
                            type: o.attrs.Type.split("/").pop(),
                            target: l,
                          }));
                    }
                  }),
                  [2]
                );
            }
          });
        });
      }),
      (e.prototype._loadBackground = function () {
        return oa(this, void 0, void 0, function () {
          var t;
          return sa(this, function (n) {
            return (
              (t = Se.get(this.source, [
                "p:sldLayout",
                "p:cSld",
                "p:bg",
                "p:bgPr",
              ])) && t["a:solidFill"]
                ? (this.background = $r(t["a:solidFill"], this.theme))
                : t && t["a:gradFill"]
                  ? (this.background = G0(t["a:gradFill"], this.theme, this))
                  : t &&
                    t["a:blipFill"] &&
                    (this.background = W0(t["a:blipFill"], this.pptx, this)),
              [2]
            );
          });
        });
      }),
      (e.prototype._loadNodes = function () {
        return oa(this, void 0, void 0, function () {
          var t;
          return sa(this, function (n) {
            return (
              (t = Se.get(this.source, ["p:sldLayout", "p:cSld", "p:spTree"])),
              K0(this.nodes, t, this.pptx, this),
              [2]
            );
          });
        });
      }),
      (e.prototype.getColorThemeName = function (t) {
        return this.slideMaster.getColorThemeName(t);
      }),
      (e.prototype.getNodeByType = function (t) {
        return this.nodes.find(function (n) {
          return n.type === t;
        });
      }),
      (e.prototype.getNodeByIdx = function (t) {
        return this.nodes.find(function (n) {
          return n.idx === t;
        });
      }),
      (e.prototype.getNodeInheritAttrsByType = function (t, n) {
        var r = this.slideMaster.getNodeByType(t);
        return r && Se.get(r.source, n);
      }),
      (e.prototype.getNodeInheritAttrsByIdx = function (t, n) {
        var r = this.slideMaster.getNodeByIdx(t);
        return r && Se.get(r.source, n);
      }),
      e
    );
  })(),
  hJe = (function () {
    function e(t, n, r) {
      ((this.slideType = "slideMaster"),
        (this.rels = {}),
        (this.background = { type: "none" }),
        (this.textStyles = { titleStyle: {}, bodyStyle: {}, otherStyle: {} }),
        (this.defaultTextStyle = {}),
        (this.nodes = []),
        (this.tableStyles = {}),
        (this.name = t),
        (this.source = n),
        (this.pptx = r),
        this.load());
    }
    return (
      Object.defineProperty(e.prototype, "_relsPath", {
        get: function () {
          return (
            this.name.replace(
              "slideMasters/slideMaster",
              "slideMasters/_rels/slideMaster",
            ) + ".rels"
          );
        },
        enumerable: !1,
        configurable: !0,
      }),
      (e.prototype.load = function () {
        return oa(this, void 0, void 0, function () {
          return sa(this, function (t) {
            switch (t.label) {
              case 0:
                return [4, this._parseRels()];
              case 1:
                return (
                  t.sent(),
                  this._parseColorMap(),
                  this._parseBackground(),
                  this._parseTextStyles(),
                  this._parseTableStyles(),
                  this._parseDefaultTextStyle(),
                  this._loadNodes(),
                  [2]
                );
            }
          });
        });
      }),
      (e.prototype._parseRels = function () {
        return oa(this, void 0, void 0, function () {
          var t,
            n,
            r,
            a = this;
          return sa(this, function (i) {
            switch (i.label) {
              case 0:
                return ((n = $i), [4, this.pptx.getXmlByPath(this._relsPath)]);
              case 1:
                return (
                  (t = n.apply(void 0, [i.sent()])),
                  (r = Se.get(t, ["Relationships", "Relationship"]) || []),
                  Array.isArray(r) || (r = [r]),
                  r.forEach(function (o) {
                    switch (Se.get(o, ["attrs", "Type"])) {
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        var s = o.attrs.Target.replace("../", "ppt/");
                        (s.startsWith("/ppt") && (s = s.substr(1)),
                          (a.theme = a.pptx.getTheme(s)));
                        break;
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio":
                      case "http://schemas.microsoft.com/office/2007/relationships/media":
                      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video":
                        var l = o.attrs.Target.replace("../", "ppt/");
                        (l.startsWith("/ppt") && (l = l.substr(1)),
                          (a.rels[o.attrs.Id] = {
                            type: o.attrs.Type.split("/").pop(),
                            target: l,
                          }));
                    }
                  }),
                  [2]
                );
            }
          });
        });
      }),
      (e.prototype._parseColorMap = function () {
        this.colorMap = Se.omit(
          Se.get(this.source, ["p:sldMaster", "p:clrMap", "attrs"]) || {},
          ["order"],
        );
      }),
      (e.prototype.getColorThemeName = function (t) {
        return this.colorMap[t] || t;
      }),
      (e.prototype._parseBackground = function () {
        var t = Se.get(this.source, [
            "p:sldMaster",
            "p:cSld",
            "p:bg",
            "p:bgPr",
          ]),
          n = Se.get(this.source, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);
        t && t["a:solidFill"]
          ? (this.background = $r(t["a:solidFill"], this.theme, this))
          : t && t["a:gradFill"]
            ? (this.background = G0(t["a:gradFill"], this.theme, this))
            : t && t["a:blipFill"]
              ? (this.background = W0(t["a:blipFill"], this.pptx, this))
              : n && (this.background = $r(n, this.theme, this));
      }),
      (e.prototype._parseDefaultTextStyle = function () {
        var t = this,
          n = this.pptx.defaultTextStyleSource;
        Object.keys(n).forEach(function (r) {
          if (r.startsWith("a:") && r.endsWith("pPr")) {
            var a = r.substr(2, r.length - 5),
              i = Se.get(n[r], ["a:defRPr"]);
            t.defaultTextStyle[a] = {
              props: fW(n[r]),
              defRPr: dW(i, t.theme, t),
            };
          }
        });
      }),
      (e.prototype._parseTextStyles = function () {
        var t = this,
          n = Se.get(this.source, ["p:sldMaster", "p:txStyles"]);
        ["titleStyle", "bodyStyle", "otherStyle"].forEach(function (r) {
          var a = t.textStyles[r],
            i = Se.get(n, "p:".concat(r)) || {};
          Object.keys(i).forEach(function (o) {
            if (o.startsWith("a:") && o.endsWith("pPr")) {
              var s = o.substr(2, o.length - 5);
              ((a[s] = {}), (a[s].props = fW(i[o])));
              var l = Se.get(i[o], ["a:defRPr"]);
              a[s].defRPr = dW(l, t.theme, t);
            }
          });
        });
      }),
      (e.prototype._parseTableStyles = function () {
        var t = this,
          n = {},
          r =
            Se.get(this.pptx.tableStyles, ["a:tblStyleLst", "a:tblStyle"]) ||
            [];
        (Array.isArray(r) || (r = [r]),
          r.forEach(function (a) {
            var i = Se.get(a, ["attrs", "styleId"]);
            ((n[i] = {}),
              Object.keys(a).forEach(function (o) {
                if (o.startsWith("a:")) {
                  var s = o.substr(2);
                  n[i][s] = {};
                  var l = Se.get(a[o], ["a:tcStyle"]);
                  if (l) {
                    var u = {};
                    Se.get(l, ["a:fill", "a:solidFill"]) &&
                      (u.background = $r(
                        Se.get(l, ["a:fill", "a:solidFill"]),
                        t.theme,
                        t,
                      ));
                    var c = Se.get(l, "a:tcBdr");
                    (c &&
                      ((u.border = {}),
                      Object.keys(c).forEach(function (h) {
                        if (h.startsWith("a:")) {
                          var p = h.substr(2),
                            g = Se.get(c[h], ["a:ln"]);
                          u.border[p] = Av(g, t.theme, t);
                        }
                      })),
                      (n[i][s].tcStyle = u));
                  }
                  var f = Se.get(a[o], ["a:tcTxStyle"]);
                  if (f) {
                    var d = {};
                    ((d.color = $r(f, t.theme, t)),
                      Se.get(f, ["attrs", "b"]) === "on" && (d.bold = !0),
                      (n[i][s].tcTxStyle = d));
                  }
                }
              }));
          }),
          (this.tableStyles = n));
      }),
      (e.prototype._loadNodes = function () {
        return oa(this, void 0, void 0, function () {
          var t;
          return sa(this, function (n) {
            return (
              (t = Se.get(this.source, ["p:sldMaster", "p:cSld", "p:spTree"])),
              K0(this.nodes, t, this.pptx, this),
              [2]
            );
          });
        });
      }),
      (e.prototype.getNodeByType = function (t) {
        return this.nodes.find(function (n) {
          return n.type === t;
        });
      }),
      (e.prototype.getNodeByIdx = function (t) {
        return this.nodes.find(function (n) {
          return n.idx === t;
        });
      }),
      (e.prototype.getNodeInheritAttrsByType = function (t, n) {}),
      (e.prototype.getNodeInheritAttrsByIdx = function (t, n) {}),
      e
    );
  })(),
  vJe = (function () {
    function e(t, n, r) {
      ((this.clrScheme = {}),
        (this.borderScheme = []),
        (this.fontScheme = {}),
        (this.name = t),
        (this.source = n),
        (this.pptx = r),
        this._parseClrScheme(),
        this._parseFontScheme(),
        this._parseLineStyleLst());
    }
    return (
      (e.prototype._parseClrScheme = function () {
        var t = Se.get(this.source, [
          "a:theme",
          "a:themeElements",
          "a:clrScheme",
        ]);
        for (var n in t)
          if (n.startsWith("a:")) {
            var r = n.substring(2),
              a =
                Se.get(t[n], ["a:sysClr", "attrs", "lastClr"]) ||
                Se.get(t[n], ["a:srgbClr", "attrs", "val"]);
            this.clrScheme[r] = "#" + a;
          }
      }),
      (e.prototype._parseFontScheme = function () {
        var t = Se.get(this.source, ["a:theme", "a:themeElements", "a:fontScheme"]),
          g = function (n) {
            return {
              latin: Se.get(t, [n, "a:latin", "attrs", "typeface"]),
              ea: Se.get(t, [n, "a:ea", "attrs", "typeface"]),
            };
          };
        ((this.fontScheme.major = g("a:majorFont")),
          (this.fontScheme.minor = g("a:minorFont")));
      }),
      (e.prototype._parseLineStyleLst = function () {
        var t = this,
          n =
            Se.get(this.source, [
              "a:theme",
              "a:themeElements",
              "a:fmtScheme",
              "a:lnStyleLst",
              "a:ln",
            ]) || [];
        this.borderScheme = n.map(function (r) {
          var a = { color: {} };
          return (
            Se.get(r, ["attrs", "w"]) &&
              (a.width = or(parseInt(Se.get(r, ["attrs", "w"])))),
            Se.get(r, ["attrs", "algn"]) &&
              (a.algn = Se.get(r, ["attrs", "algn"])),
            Se.get(r, ["attrs", "cap"]) &&
              (a.cap = Se.get(r, ["attrs", "cap"])),
            Se.get(r, ["attrs", "cmpd"]) &&
              (a.cmpd = Se.get(r, ["attrs", "cmpd"])),
            Se.get(r, ["a:miter", "attrs", "lim"]) &&
              (a.miterLim = or(
                parseInt(Se.get(r, ["a:miter", "attrs", "lim"])),
              )),
            Se.get(r, ["a:prstDash", "attrs", "val"]) &&
              (a.type = Se.get(r, ["a:prstDash", "attrs", "val"])),
            Se.get(r, ["a:solidFill"]) &&
              (a.color = $r(Se.get(r, ["a:solidFill"]), t)),
            a
          );
        });
      }),
      (e.prototype.getColor = function (t) {
        if (t !== "phClr") return this.clrScheme[t] || this.defaultColor;
      }),
      (e.prototype.getLineStyle = function (t) {
        return this.borderScheme[t - 1];
      }),
      e
    );
  })(),
  hW = (function () {
    function e() {
      ((this.slides = []),
        (this.slideLayouts = []),
        (this.slideMaster = []),
        (this.themes = []),
        (this.medias = {}),
        (this.wps = !1));
    }
    return (
      (e.prototype.load = function (t) {
        return oa(this, void 0, void 0, function () {
          var n, r;
          return sa(this, function (a) {
            switch (a.label) {
              case 0:
                return ((n = new c3e()), (r = this), [4, n.loadAsync(t)]);
              case 1:
                return (
                  (r._zipContents = a.sent()),
                  [4, this._loadThumbnail()]
                );
              case 2:
                return (a.sent(), [4, this._loadMedia()]);
              case 3:
                return (a.sent(), [4, this._loadPresentation()]);
              case 4:
                return (a.sent(), [4, this._loadContentTypes()]);
              case 5:
                return (a.sent(), [2]);
            }
          });
        });
      }),
      (e.prototype._loadThumbnail = function () {
        return oa(this, void 0, void 0, function () {
          var t;
          return sa(this, function (n) {
            switch (n.label) {
              case 0:
                return (
                  n.trys.push([0, 3, , 4]),
                  this._zipContents.files["docProps/thumbnail.jpeg"]
                    ? [
                        4,
                        this._zipContents.files[
                          "docProps/thumbnail.jpeg"
                        ].async("base64"),
                      ]
                    : [3, 2]
                );
              case 1:
                ((t = n.sent()),
                  (this.thumbnail = "data:image/jpeg;base64," + t),
                  (n.label = 2));
              case 2:
                return [3, 4];
              case 3:
                return (n.sent(), [3, 4]);
              case 4:
                return [2];
            }
          });
        });
      }),
      (e.prototype._loadPresentation = function () {
        return oa(this, void 0, void 0, function () {
          var t, n, r, a, i, o;
          return sa(this, function (s) {
            switch (s.label) {
              case 0:
                return (
                  s.trys.push([0, 4, , 5]),
                  [
                    4,
                    this._zipContents.files["ppt/presentation.xml"].async(
                      "text",
                    ),
                  ]
                );
              case 1:
                return (
                  (t = s.sent()),
                  (n = $i(t)),
                  (r = n["p:presentation"]["p:sldSz"].attrs),
                  (this.width = or(parseInt(r.cx))),
                  (this.height = or(parseInt(r.cy))),
                  (this.defaultTextStyleSource = Se.get(n, [
                    "p:presentation",
                    "p:defaultTextStyle",
                  ])),
                  this._zipContents.files["docProps/app.xml"]
                    ? [
                        4,
                        this._zipContents.files["docProps/app.xml"].async(
                          "text",
                        ),
                      ]
                    : [3, 3]
                );
              case 2:
                ((a = s.sent()),
                  (i = $i(a)),
                  (o = Se.get(i, ["Properties", "Application"])),
                  (this.wps = (o == null ? void 0 : o.includes("WPS")) || !1),
                  (s.label = 3));
              case 3:
                return [3, 5];
              case 4:
                return (s.sent(), [3, 5]);
              case 5:
                return [2];
            }
          });
        });
      }),
      // One broken part must not cost the whole deck: parse each theme,
      // master, layout and slide on its own and skip what throws.
      (e.prototype._readPart = function (part) {
        var f =
          this._zipContents && this._zipContents.files
            ? this._zipContents.files[part]
            : null;
        if (!f) {
          console.warn("[pptx] missing part " + part);
          return Promise.resolve(null);
        }
        return Promise.resolve(f.async("text"))
          .then(function (text) {
            return $i(text);
          })
          .catch(function (err) {
            console.warn("[pptx] unreadable part " + part + ": " + ((err && err.message) || err));
            return null;
          });
      }),
      (e.prototype._skipPart = function (what, part, fn) {
        try {
          return fn();
        } catch (err) {
          console.warn(
            "[pptx] skipped " + what + " " + part + ": " + ((err && err.message) || err),
          );
          return null;
        }
      }),
      (e.prototype._skipPartAsync = function (what, part, make) {
        return Promise.resolve()
          .then(function () {
            var inst = make();
            return Promise.resolve(inst.load()).then(function () {
              return inst;
            });
          })
          .catch(function (err) {
            console.warn(
              "[pptx] skipped " + what + " " + part + ": " + ((err && err.message) || err),
            );
            return null;
          });
      }),
      (e.prototype._loadContentTypes = function () {
        return oa(this, void 0, void 0, function () {
          var t, n, r, a, i, o, s, l, u, c, f, d, h, p, g, y, m, v, b, _, S;
          return sa(this, function (C) {
            switch (C.label) {
              case 0:
                return (
                  C.trys.push([0, 22, , 23]),
                  (n = $i),
                  [
                    4,
                    this._zipContents.files["[Content_Types].xml"].async(
                      "text",
                    ),
                  ]
                );
              case 1:
                if (
                  ((t = n.apply(void 0, [C.sent()])),
                  (r = t.Types.Override),
                  !(a = r.filter(function (T) {
                    return (
                      T.attrs.ContentType ===
                      "application/vnd.openxmlformats-officedocument.theme+xml"
                    );
                  })))
                )
                  return [3, 5];
                ((i = 0), (o = a), (C.label = 2));
              case 2:
                return i < o.length
                  ? ((s = o[i]),
                    (v = s.attrs.PartName.substr(1)),
                    (l = $i),
                    [4, this._readPart(v)])
                  : [3, 5];
              case 3:
                ((b = C.sent()),
                  (y = b && this._skipPart("theme", v, () => new vJe(v, b, this))),
                  y && this.themes.push(y),
                  (C.label = 4));
              case 4:
                return (i++, [3, 2]);
              case 5:
                return (u = r.find(function (T) {
                  return (
                    T.attrs.ContentType ===
                    "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
                  );
                }))
                  ? ((v = u.attrs.PartName.substr(1)),
                    (c = $i),
                    [4, this._readPart(v)])
                  : [3, 7];
              case 6:
                ((b = C.sent()),
                  b && (this.tableStyles = b),
                  (C.label = 7));
              case 7:
                ((f = r.filter(function (T) {
                  return (
                    T.attrs.ContentType ===
                    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
                  );
                })),
                  (m = 0),
                  (C.label = 8));
              case 8:
                return m < f.length
                  ? ((v = f[m].attrs.PartName.substr(1)),
                    (d = $i),
                    [4, this._readPart(v)])
                  : [3, 11];
              case 9:
                ((b = C.sent()),
                  (y = b && this._skipPart("master", v, () => new hJe(v, b, this))),
                  y && this.slideMaster.push(y),
                  (C.label = 10));
              case 10:
                return (m++, [3, 8]);
              case 11:
                ((h = r.filter(function (T) {
                  return (
                    T.attrs.ContentType ===
                    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
                  );
                })),
                  (m = 0),
                  (C.label = 12));
              case 12:
                return m < h.length
                  ? ((v = h[m].attrs.PartName.substr(1)),
                    (p = $i),
                    [4, this._readPart(v)])
                  : [3, 16];
              case 13:
                return (
                  (b = C.sent()),
                  b
                    ? [4, this._skipPartAsync("layout", v, () => new dJe(v, b, this))]
                    : [3, 14]
                );
              case 14:
                ((g = C.sent()), g && this.slideLayouts.push(g), (C.label = 15));
              case 15:
                return (m++, [3, 12]);
              case 16:
                ((y = r.filter(function (T) {
                  return (
                    T.attrs.ContentType ===
                    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
                  );
                })),
                  (m = 0),
                  (C.label = 17));
              case 17:
                return m < y.length
                  ? ((v = y[m].attrs.PartName.substr(1)),
                    (_ = $i),
                    [4, this._readPart(v)])
                  : [3, 21];
              case 18:
                return (
                  (b = C.sent()),
                  b
                    ? [4, this._skipPartAsync("slide", v, () => new fJe(v, b, this))]
                    : [3, 19]
                );
              case 19:
                ((S = C.sent()), S && this.slides.push(S), (C.label = 20));
              case 20:
                return (m++, [3, 17]);
              case 21:
                return (
                  this.slides.sort(function (T, k) {
                    return T.index - k.index;
                  }),
                  [3, 23]
                );
              case 22:
                return (C.sent(), [3, 23]);
              case 23:
                return [2];
            }
          });
        });
      }),
      (e.prototype._loadMedia = function () {
        return oa(this, void 0, void 0, function () {
          var t,
            n,
            r,
            a,
            i,
            o,
            s,
            l,
            u,
            c,
            f,
            d,
            h,
            p = this;
          return sa(this, function (g) {
            switch (g.label) {
              case 0:
                ((t = Object.keys(this._zipContents.files)
                  .filter(function (m) {
                    return m.startsWith("ppt/media/image");
                  })
                  .map(function (m) {
                    return p._zipContents.files[m];
                  })),
                  (n = 0),
                  (r = t),
                  (g.label = 1));
              case 1:
                return n < r.length
                  ? ((a = r[n]),
                    (i = (function (m) {
                      var v;
                      switch (m) {
                        case "jpg":
                        case "jpeg":
                          v = "image/jpeg";
                          break;
                        case "png":
                          v = "image/png";
                          break;
                        case "gif":
                          v = "image/gif";
                          break;
                        case "emf":
                          v = "image/x-emf";
                          break;
                        case "wmf":
                          v = "image/x-wmf";
                          break;
                        default:
                          v = "image/*";
                      }
                      return v;
                    })((y = a.name).substr(2 + (~-y.lastIndexOf(".") >>> 0)))),
                    [4, a.async("base64")])
                  : [3, 4];
              case 2:
                ((o = g.sent()),
                  (this.medias[a.name] = "data:"
                    .concat(i, ";base64,")
                    .concat(o)),
                  (g.label = 3));
              case 3:
                return (n++, [3, 1]);
              case 4:
                ((s = Object.keys(this._zipContents.files)
                  .filter(function (m) {
                    return (
                      m.startsWith("ppt/media/media") &&
                      ["mp3", "wav", "ogg", "mp4", "webm"].includes(
                        m.split(".").pop().toLowerCase(),
                      )
                    );
                  })
                  .map(function (m) {
                    return p._zipContents.files[m];
                  })),
                  (l = 0),
                  (u = s),
                  (g.label = 5));
              case 5:
                return l < u.length
                  ? ((c = u[l]),
                    (f = c.name.split(".").pop().toLowerCase()),
                    [4, c.async("arraybuffer")])
                  : [3, 8];
              case 6:
                ((d = g.sent()),
                  (h = new Blob([d], {
                    type: ""
                      .concat(
                        ["mp3", "wav"].includes(f) ? "audio" : "video",
                        "/",
                      )
                      .concat(f),
                  })),
                  (this.medias[c.name] = URL.createObjectURL(h)),
                  (g.label = 7));
              case 7:
                return (l++, [3, 5]);
              case 8:
                return [2];
            }
            var y;
          });
        });
      }),
      (e.prototype.getXmlByPath = function (t) {
        return oa(this, void 0, void 0, function () {
          return sa(this, function (n) {
            switch (n.label) {
              case 0:
                if (!this._zipContents.files[t]) throw new Error("文件不存在");
                return [4, this._zipContents.files[t].async("text")];
              case 1:
                return [2, n.sent()];
            }
          });
        });
      }),
      (e.prototype.getSlideLayout = function (t) {
        return this.slideLayouts.find(function (n) {
          return n.name === t;
        });
      }),
      (e.prototype.getSlideMaster = function (t) {
        return this.slideMaster.find(function (n) {
          return n.name === t;
        });
      }),
      (e.prototype.getTheme = function (t) {
        return this.themes.find(function (n) {
          return n.name === t;
        });
      }),
      (e.prototype.getMedia = function (t) {
        return this.medias[t];
      }),
      e
    );
  })();
function UE(e) {
  var t = e.extend,
    n = e.offset,
    r = e.clip,
    a = e.base64,
    i = e.audioFile,
    o = e.videoFile,
    s = document.createElement("div");
  (s.style.setProperty("position", "absolute"),
    s.style.setProperty("left", n.x + "px"),
    s.style.setProperty("top", n.y + "px"));
  var l,
    u,
    c,
    f,
    d = document.createElement("div");
  (d.style.setProperty("position", "absolute"),
    d.style.setProperty("left", "0"),
    d.style.setProperty("top", "0"),
    d.style.setProperty("width", t.w + "px"),
    d.style.setProperty("height", t.h + "px"),
    d.style.setProperty("overflow", "hidden"),
    r
      ? ((l =
          t.w / (1 - (r.l === void 0 ? 0 : r.l) - (r.r === void 0 ? 0 : r.r))),
        (u =
          t.h / (1 - (r.t === void 0 ? 0 : r.t) - (r.b === void 0 ? 0 : r.b))),
        (c = -1 * l * (r.l === void 0 ? 0 : r.l)),
        (f = -1 * u * (r.t === void 0 ? 0 : r.t)))
      : ((l = t.w), (u = t.h), (c = 0)));
  var h = document.createElement("img");
  if (
    ((h.src = a),
    (h.width = l),
    (h.height = u),
    h.style.setProperty("position", "absolute"),
    h.style.setProperty("left", c + "px"),
    h.style.setProperty("top", f + "px"),
    d.append(h),
    s.append(d),
    i)
  ) {
    var p = document.createElement("audio");
    ((p.style.position = "absolute"),
      (p.style.left = "0"),
      (p.style.top = "0"),
      (p.src = i),
      (p.controls = !0),
      (p.style.transform = "translate(-50%)"),
      s.append(p));
  }
  if (o) {
    var g = document.createElement("video");
    ((g.style.position = "absolute"),
      (g.style.left = "0"),
      (g.style.top = "0"),
      (g.width = t.w),
      (g.height = t.h),
      (g.src = o),
      (g.controls = !0),
      s.append(g));
  }
  return s;
}
function pJe(e, t, n) {
  var r,
    a = e.props,
    i = e.text,
    o = Tt(Tt({}, t), a),
    s = document.createElement("span");
  s.innerHTML = typeof i == "string" ? i : "";
  var l = 18;
  o.size &&
    (!((r = n == null ? void 0 : n.normAutofit) === null || r === void 0) &&
    r.fontScale
      ? ((l = o.size * n.normAutofit.fontScale), (s.style.fontSize = l + "px"))
      : ((l = o.size), (s.style.fontSize = l + "px")));
  var u = Qn(o.color);
  // runs inherit the body's font when neither the run nor its paragraph sets one
  (o.typeface || (o.typeface = n && n.typeface),
    o.typefaceEa || (o.typefaceEa = n && n.typefaceEa),
    u && (s.style.color = u));
  var c = /^[^\u4e00-\u9fff]+$/;
  if (o.typeface)
    switch (
      ((s.style.fontFamily = [o.typeface, o.typefaceEa]
        .filter(function (v) {
          return v && v.charAt(0) !== "+";
        })
        .join(", ")),
      o.typeface)
    ) {
      case "DengXian":
        c.test(i) && (s.style.letterSpacing = -0.04 * l + "px");
        break;
      case "DengXian Light":
        c.test(i) && (s.style.letterSpacing = -0.05 * l + "px");
        break;
      case "STLiti":
      case "SimSun":
      case "NSimSun":
      case "SimHei":
        c.test(i) &&
          (s.style.fontSize = 0.85 * parseInt(s.style.fontSize) + "px");
        break;
      case "华文中宋":
      case "Fira Sans Extra Condensed Medium":
        s.style.fontSize = 0.85 * parseInt(s.style.fontSize) + "px";
        break;
      case "FangSong":
        s.style.letterSpacing = -0.08 * l + "px";
    }
  else c.test(i) && (s.style.letterSpacing = -0.04 * l + "px");
  return (
    o.bold && (s.style.fontWeight = "bold"),
    o.italic && (s.style.fontStyle = "italic"),
    o.underline &&
      o.underline !== "none" &&
      (s.style.textDecoration = "underline"),
    o.background && (s.style.backgroundColor = Qn(o.background)),
    (s.style.wordBreak = "break-word"),
    s
  );
}
function gJe(e, t, n) {
  var r = document.createElement("span"),
    a = e.firstElementChild;
  switch (
    ((r.style.fontSize = a.style.fontSize),
    (r.style.color = a.style.color),
    (r.style.fontWeight = a.style.fontWeight),
    (r.style.fontStyle = a.style.fontStyle),
    (r.style.marginRight = "10px"),
    t.buAutoNum)
  ) {
    case "arabicPeriod":
    default:
      r.textContent = n + ".";
      break;
    case "circleNumDbPlain":
      r.textContent =
        [
          "①",
          "②",
          "③",
          "④",
          "⑤",
          "⑥",
          "⑦",
          "⑧",
          "⑨",
          "⑩",
          "⑪",
          "⑫",
          "⑬",
          "⑭",
          "⑮",
          "⑯",
          "⑰",
          "⑱",
          "⑲",
          "⑳",
        ][n - 1] || n + "";
      break;
    case "romanUcPeriod":
      r.textContent =
        (function (i) {
          var o = [
            { value: 1e3, numeral: "M" },
            { value: 900, numeral: "CM" },
            { value: 500, numeral: "D" },
            { value: 400, numeral: "CD" },
            { value: 100, numeral: "C" },
            { value: 90, numeral: "XC" },
            { value: 50, numeral: "L" },
            { value: 40, numeral: "XL" },
            { value: 10, numeral: "X" },
            { value: 9, numeral: "IX" },
            { value: 5, numeral: "V" },
            { value: 4, numeral: "IV" },
            { value: 1, numeral: "I" },
          ];
          if (typeof i != "number" || i < 1 || i > 3999)
            throw new Error("Input must be a number between 1 and 3999.");
          for (var s = "", l = 0; l < o.length; l++)
            for (; i >= o[l].value;) ((s += o[l].numeral), (i -= o[l].value));
          return s;
        })(n) + ".";
      break;
    case "alphaUcPeriod":
      r.textContent = Ck(n) + ".";
      break;
    case "alphaLcPeriod":
      r.textContent = Ck(n).toLowerCase() + ".";
      break;
    case "alphaLcParenR":
      r.textContent = Ck(n).toLowerCase() + ")";
      break;
    case "ea1JpnChsDbPeriod":
      r.textContent =
        (function (i) {
          var o = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
          if (!Number.isInteger(i) || i < 0) return "";
          for (var s = "", l = i.toString(), u = 0; u < l.length; u++)
            s += o[parseInt(l[u], 10)];
          return s;
        })(n) + ".";
  }
  e.prepend(r);
}
function AS(e, t, n) {
  var r, a, i, o;
  (t === void 0 && (t = 0), n === void 0 && (n = {}));
  var s = e.inheritProps,
    l = e.inheritRProps,
    u = e.props,
    c = e.rows,
    f = Tt(Tt({}, s), u),
    d = function () {
      for (var C, T, k = 0, M = 0, I = c; M < I.length; M++) {
        var L = I[M];
        L.props && L.props.size && (k = Math.max(k, L.props.size));
      }
      var E =
        ((T =
          (C = n == null ? void 0 : n.bodyProps) === null || C === void 0
            ? void 0
            : C.normAutofit) === null || T === void 0
          ? void 0
          : T.fontScale) || 1;
      return (k || l.size || 18) * E;
    },
    h = document.createElement("div"),
    p = n.isFirst ? 0 : f.spaceBefore || 0,
    g = n.isLast ? 0 : f.spaceAfter || 0;
  ((h.style.margin = "0px"),
    (h.style.padding = ""
      .concat(Math.floor(p), "px 0px ")
      .concat(Math.floor(g), "px 0px")));
  var y = document.createElement("p");
  ((y.style.margin = "0"),
    (y.style.padding = "0px"),
    (y.style.wordBreak = "break-word"),
    (y.style.textAlign =
      (f.align &&
        { ctr: "center", l: "left", r: "right", dist: "justify" }[f.align]) ||
      "center"),
    f.align === "dist" && (y.style.textAlignLast = "justify"));
  // Measured in-browser: 18px Calibri has ascent 17 + descent 5 = 22px natural line
  // box. PowerPoint does NOT use that: spcPct 100000 means 1.0 x font size (18px),
  // and the glyphs simply overflow the box. Keeping 1 matches golden's row heights;
  // "normal" added 22px lines and grew every table row.
  // Measured on slide 6: golden's three bullet lines are 29px apart at a 24px font
  // (1.21x, i.e. the font's own line box), while ours were 24px (1.0x) - the block
  // ended 40px short and looked like its tail had been cut off. Table cells are the
  // exception: there PowerPoint keeps a 1.0x pitch (slide 3's rows match golden at
  // exactly a:tr h). So: font metrics outside tables, 1.0 inside.
  var m = f.hasOwnProperty("lineHeight") ? f.lineHeight : n && n.inCell ? 1 : "normal";
  if (
    (!(
      (a =
        (r = n.bodyProps) === null || r === void 0 ? void 0 : r.normAutofit) ===
        null || a === void 0
    ) &&
      a.lnSpcReduction &&
      (m *=
        1 -
        ((o =
          (i = n.bodyProps) === null || i === void 0
            ? void 0
            : i.normAutofit) === null || o === void 0
          ? void 0
          : o.lnSpcReduction)),
    (y.style.lineHeight = m + ""),
    (y.style.fontSize = d() + "px"),
    c.length)
  ) {
    for (var v = 0, b = c; v < b.length; v++) {
      var _ = b[v];
      _.isBr
        ? y.appendChild(document.createElement("br"))
        : y.appendChild(
            pJe(
              _,
              Tt(Tt({}, l), { marginTop: Math.floor(0.2 * d()) }),
              n.bodyProps,
            ),
          );
    }
    (f.buAutoNum
      ? gJe(y, f, t)
      : f.buChar &&
        (function (C, T) {
          var k = document.createElement("span"),
            M = C.firstElementChild;
          ((k.style.color = M.style.color),
            (k.style.fontSize = M.style.fontSize),
            (k.textContent =
              { n: "■", l: "●", u: "◆", p: "□", ü: "✔", Ø: "➢", "•": "•" }[
                T.buChar
              ] || "■"),
            (k.style.marginRight = "10px"),
            C.prepend(k));
        })(y, f),
      (y.style.paddingLeft = (f.marginLeft || 0) + (f.indent || 0) + "px"));
  } else {
    var S = document.createElement("span");
    ((S.innerHTML = "&nbsp;"),
      (S.style.fontSize = l.size + "px"),
      y.appendChild(S));
  }
  return (h.appendChild(y), h);
}
var Pee = "http://www.w3.org/2000/svg";
function lt(e) {
  return document.createElementNS(Pee, e);
}
function jl(e) {
  var t = e.extend;
  return 0.16667 * Math.min(t.w, t.h);
}
function rn(e, t, n) {
  n === void 0 && (n = 0);
  var r = t.prstGeom,
    a = ((r == null ? void 0 : r.gd) || []).find(function (i) {
      return i.name === e;
    });
  return a ? a.fmla : n;
}
function at(e, t, n) {
  n === void 0 && (n = 0);
  var r = t.extend,
    a = t.prstGeom,
    i = ((a == null ? void 0 : a.gd) || []).find(function (o) {
      return o.name === e;
    });
  return i ? Math.min(r.w, r.h) * i.fmla : n;
}
function vW(e, t, n) {
  n === void 0 && (n = 0);
  var r = t.extend,
    a = t.prstGeom,
    i = ((a == null ? void 0 : a.gd) || []).find(function (o) {
      return o.name === e;
    });
  return i ? Math.max(r.w, r.h) * i.fmla : n;
}
function Cl(e, t) {
  var n = 0;
  switch (e) {
    case "sm":
      n = 1 * t;
      break;
    case "med":
      n = 1.5 * t;
      break;
    case "lg":
      n = 2.5 * t;
  }
  return Math.max(n, 2);
}
function pW(e, t, n, r) {
  r === void 0 && (r = !1);
  var a = e.border,
    i = a === void 0 ? {} : a,
    o = e.uuid,
    s = i.headEnd,
    l = i.width,
    u = i.color,
    c = i.tailEnd,
    f = r ? s : c,
    d = f.len,
    h = d === void 0 ? "med" : d,
    p = f.w,
    g = p === void 0 ? "med" : p,
    y = Cl(h, l),
    m = Cl(g, l),
    v = lt("defs"),
    b = lt("marker"),
    _ = "marker-".concat(o, "-").concat(r ? "start" : "end");
  (b.setAttribute("id", _),
    b.setAttribute("viewBox", "0 0 ".concat(2 * y, " ").concat(2 * m)),
    b.setAttribute("refX", y + "px"),
    b.setAttribute("refY", m + "px"),
    b.setAttribute("markerWidth", 2 * y + "px"),
    b.setAttribute("markerHeight", 2 * m + "px"),
    b.setAttribute("orient", "auto"),
    b.setAttribute("markerUnits", "userSpaceOnUse"));
  var S = lt("ellipse");
  (S.setAttribute("cx", y + "px"),
    S.setAttribute("cy", m + "px"),
    S.setAttribute("rx", y + "px"),
    S.setAttribute("ry", m + "px"),
    S.setAttribute("fill", Qn(u) || "transparent"),
    b.appendChild(S),
    v.appendChild(b),
    t.appendChild(v),
    n.setAttribute(r ? "marker-start" : "marker-end", "url(#".concat(_, ")")));
}
function gW(e, t, n, r) {
  r === void 0 && (r = !1);
  var a = e.border,
    i = a === void 0 ? {} : a,
    o = e.uuid,
    s = i.headEnd,
    l = i.width,
    u = i.color,
    c = i.tailEnd,
    f = r ? s : c,
    d = f.len,
    h = d === void 0 ? "med" : d,
    p = f.w,
    g = p === void 0 ? "med" : p,
    y = Cl(h, l),
    m = Cl(g, l),
    v = lt("defs"),
    b = lt("marker"),
    _ = "marker-".concat(o, "-").concat(r ? "start" : "end");
  (b.setAttribute("id", _),
    b.setAttribute("viewBox", "0 0 ".concat(2 * y, " ").concat(2 * m)),
    b.setAttribute("refX", (r ? 0.9 * y : 1.1 * y) + "px"),
    b.setAttribute("refY", m + "px"),
    b.setAttribute("markerWidth", 2 * y + "px"),
    b.setAttribute("markerHeight", 2 * m + "px"),
    b.setAttribute("orient", "auto"),
    b.setAttribute("markerUnits", "userSpaceOnUse"));
  var S = lt("path"),
    C = r
      ? [
          "M ".concat(2 * y, ",0"),
          "L 0,".concat(m),
          "L ".concat(2 * y, ",").concat(2 * m),
          "Z",
        ].join(" ")
      : [
          "M 0,0",
          "L ".concat(2 * y, ",").concat(m),
          "L 0,".concat(2 * m),
          "Z",
        ].join(" ");
  (S.setAttribute("d", C),
    S.setAttribute("fill", Qn(u) || "transparent"),
    b.appendChild(S),
    v.appendChild(b),
    t.appendChild(v),
    n.setAttribute(r ? "marker-start" : "marker-end", "url(#".concat(_, ")")));
}
function mW(e, t, n, r) {
  r === void 0 && (r = !1);
  var a = e.border,
    i = a === void 0 ? {} : a,
    o = e.uuid,
    s = i.headEnd,
    l = i.width,
    u = i.color,
    c = i.tailEnd,
    f = r ? s : c,
    d = f.len,
    h = d === void 0 ? "med" : d,
    p = f.w,
    g = p === void 0 ? "med" : p,
    y = Cl(h, l),
    m = Cl(g, l),
    v = lt("defs"),
    b = lt("marker"),
    _ = "marker-".concat(o, "-").concat(r ? "start" : "end");
  (b.setAttribute("id", _),
    b.setAttribute("viewBox", "0 0 ".concat(2 * y, " ").concat(2 * m)),
    b.setAttribute("refX", y + "px"),
    b.setAttribute("refY", m + "px"),
    b.setAttribute("markerWidth", 2 * y + "px"),
    b.setAttribute("markerHeight", 2 * m + "px"),
    b.setAttribute("orient", "auto"),
    b.setAttribute("markerUnits", "userSpaceOnUse"));
  var S = lt("path"),
    C = [
      "M 0,".concat(m),
      "L ".concat(y, ",0"),
      "L ".concat(2 * y, ",").concat(m),
      "L ".concat(y, ",").concat(2 * m),
      "Z",
    ].join(" ");
  (S.setAttribute("d", C),
    S.setAttribute("fill", Qn(u) || "transparent"),
    b.appendChild(S),
    v.appendChild(b),
    t.appendChild(v),
    n.setAttribute(r ? "marker-start" : "marker-end", "url(#".concat(_, ")")));
}
function yW(e, t, n, r) {
  r === void 0 && (r = !1);
  var a = e.border,
    i = a === void 0 ? {} : a,
    o = e.uuid,
    s = i.headEnd,
    l = i.width,
    u = i.color,
    c = i.tailEnd,
    f = r ? s : c,
    d = f.len,
    h = d === void 0 ? "med" : d,
    p = f.w,
    g = p === void 0 ? "med" : p,
    y = Cl(h, l),
    m = Cl(g, l),
    v = lt("defs"),
    b = lt("marker"),
    _ = "marker-".concat(o, "-").concat(r ? "start" : "end");
  (b.setAttribute("id", _),
    b.setAttribute(
      "viewBox",
      "0 0 ".concat(2 * y + 2 * l, " ").concat(2 * m + 2 * l),
    ));
  var S = r ? (g === "lg" ? 2 * l : 3 * l) : g === "lg" ? 2 * y : 2 * y - l;
  (b.setAttribute("refX", S + "px"),
    b.setAttribute("refY", m + l + "px"),
    b.setAttribute("markerWidth", 2 * y + "px"),
    b.setAttribute("markerHeight", 2 * m + "px"),
    b.setAttribute("orient", "auto"),
    b.setAttribute("markerUnits", "userSpaceOnUse"));
  var C = lt("path"),
    T = r
      ? [
          "M ".concat(2 * y + l, ", ").concat(l),
          "L ".concat(l, ",").concat(m + l),
          "L ".concat(2 * y + l, ",").concat(2 * m + l),
        ].join(" ")
      : [
          "M ".concat(l, ", ").concat(l),
          "L ".concat(2 * y + l, ",").concat(m + l),
          "L ".concat(l, ",").concat(2 * m + l),
        ].join(" ");
  (C.setAttribute("d", T),
    C.setAttribute("stroke-width", l + "px"),
    C.setAttribute("stroke", Qn(u) || "transparent"),
    C.setAttribute("fill", "transparent"),
    C.setAttribute("stroke-linecap", "round"),
    C.setAttribute("stroke-linejoin", "miter"),
    (C.style.overflow = "visible"),
    b.appendChild(C),
    v.appendChild(b),
    t.appendChild(v),
    n.setAttribute(r ? "marker-start" : "marker-end", "url(#".concat(_, ")")));
}
function bW(e, t, n, r) {
  r === void 0 && (r = !1);
  var a = e.border,
    i = a === void 0 ? {} : a,
    o = e.uuid,
    s = i.headEnd,
    l = i.width,
    u = i.color,
    c = i.tailEnd,
    f = r ? s : c,
    d = f.len,
    h = d === void 0 ? "med" : d,
    p = f.w,
    g = p === void 0 ? "med" : p,
    y = Cl(h, l),
    m = Cl(g, l),
    v = lt("defs"),
    b = lt("marker"),
    _ = "marker-".concat(o, "-").concat(r ? "start" : "end");
  (b.setAttribute("id", _),
    b.setAttribute("viewBox", "0 0 ".concat(2 * y, " ").concat(2 * m)));
  var S = r
    ? g === "sm"
      ? 1.5 * l
      : 2 * l
    : g === "sm"
      ? 2 * y - 1.5 * l
      : 2 * y - 2 * l;
  (b.setAttribute("refX", S + "px"),
    b.setAttribute("refY", m + "px"),
    b.setAttribute("markerWidth", 2 * y + "px"),
    b.setAttribute("markerHeight", 2 * m + "px"),
    b.setAttribute("orient", "auto"),
    b.setAttribute("markerUnits", "userSpaceOnUse"));
  var C = lt("path"),
    T = r
      ? [
          "M 0, ".concat(m),
          "L ".concat(2 * y, ",0"),
          "L ".concat(y, ",").concat(m),
          "L ".concat(2 * y, ",").concat(2 * m),
          "Z",
        ].join(" ")
      : [
          "M 0,0",
          "L ".concat(2 * y, ",").concat(m),
          "L 0,".concat(2 * m),
          "L ".concat(y, ",").concat(m),
          "Z",
        ].join(" ");
  (C.setAttribute("d", T),
    C.setAttribute("fill", Qn(u) || "transparent"),
    (C.style.overflow = "visible"),
    b.appendChild(C),
    v.appendChild(b),
    t.appendChild(v),
    n.setAttribute(r ? "marker-start" : "marker-end", "url(#".concat(_, ")")));
}
function Jg(e, t, n) {
  var r = e.border,
    a = r === void 0 ? {} : r;
  e.extend;
  var i = a.headEnd,
    o = a.tailEnd;
  if (i && i.type !== "none")
    switch (i.type) {
      case "triangle":
        gW(e, t, n, !0);
        break;
      case "oval":
        pW(e, t, n, !0);
        break;
      case "diamond":
        mW(e, t, n, !0);
        break;
      case "arrow":
        yW(e, t, n, !0);
        break;
      case "stealth":
        bW(e, t, n, !0);
    }
  if (o && o.type !== "none")
    switch (o.type) {
      case "triangle":
        gW(e, t, n, !1);
        break;
      case "oval":
        pW(e, t, n, !1);
        break;
      case "diamond":
        mW(e, t, n, !1);
        break;
      case "arrow":
        yW(e, t, n, !1);
        break;
      case "stealth":
        bW(e, t, n, !1);
    }
}
function _W(e) {
  var t = e.extend,
    n = lt("rect");
  return (
    n.setAttribute("x", "0"),
    n.setAttribute("y", "0"),
    n.setAttribute("width", t.w + "px"),
    n.setAttribute("height", t.h + "px"),
    n
  );
}
function Wa(e, t, n, r, a) {
  var i = (e * Math.PI) / 180;
  return [
    i === 0 || i === 2 * Math.PI
      ? t + r
      : i === Math.PI
        ? t - r
        : i === Math.PI / 2 || i === (3 * Math.PI) / 2
          ? t
          : (i > 0 && i < Math.PI / 2) ||
              (i > (3 * Math.PI) / 2 && i < 2 * Math.PI)
            ? t +
              Math.sqrt(
                1 /
                  (1 / Math.pow(r, 2) +
                    Math.pow(Math.tan(i), 2) / Math.pow(a, 2)),
              )
            : t -
              Math.sqrt(
                1 /
                  (1 / Math.pow(r, 2) +
                    Math.pow(Math.tan(i), 2) / Math.pow(a, 2)),
              ),
    i === 0 || i === 2 * Math.PI || i === Math.PI
      ? n
      : i === Math.PI / 2
        ? n + a
        : i === (3 * Math.PI) / 2
          ? n - a
          : i > Math.PI && i < 2 * Math.PI
            ? n -
              Math.sqrt(
                1 /
                  (1 / Math.pow(a, 2) +
                    Math.pow(1 / Math.tan(i), 2) / Math.pow(r, 2)),
              )
            : n +
              Math.sqrt(
                1 /
                  (1 / Math.pow(a, 2) +
                    Math.pow(1 / Math.tan(i), 2) / Math.pow(r, 2)),
              ),
  ];
}
function Qg(e, t) {
  var n = 0;
  return (((t > e && t - e > 180) || (t < e && e - t < 180)) && (n = 1), n);
}
function em(e) {
  var t = e.extend,
    n = 0.146 * t.w,
    r = 0.146 * t.h;
  return {
    top: r,
    bottom: r,
    left: n,
    right: n,
    w: t.w - 2 * n,
    h: t.h - 2 * r,
  };
}
function KE(e) {
  var t = e.extend,
    n = e.offset,
    r = e.border,
    a = e.background,
    i = e.rotate,
    o = e.flipH,
    s = e.flipV,
    l = document.createElement("div"),
    u = n.x,
    c = n.y;
  ((l.className = "shape-wrapper shape-".concat(e.shape)),
    l.style.setProperty("position", "absolute"),
    l.style.setProperty("width", (t.w || 1) + "px"),
    l.style.setProperty("height", (t.h || 1) + "px"),
    l.style.setProperty("left", u + "px"),
    l.style.setProperty("top", c + "px"));
  var f,
    d = document.createElementNS(Pee, "svg");
  (d.style.setProperty("position", "absolute"),
    d.setAttribute("width", "100%"),
    d.setAttribute("height", "100%"),
    d.style.setProperty("left", "0"),
    d.style.setProperty("top", "0"),
    (d.style.overflow = "visible"));
  var h = { left: 0, top: 0, bottom: 0, w: t.w, h: t.h },
    p = !0;
  switch (e.shape) {
    case "customGeom":
      ((f = (function (m) {
        var v = m.prstGeom,
          b = v === void 0 ? {} : v,
          _ = m.extend,
          S = b.pathList,
          C = b.w,
          T = b.h,
          k = lt("path"),
          M = { moveTo: "M", lnTo: "L", cubicBezTo: "C", close: "Z" },
          I = _.w / C,
          L = _.h / T,
          E = S.map(function (R) {
            var D = M[R.type],
              O = Array.isArray(R.points)
                ? R.points
                    .map(function (B) {
                      return "".concat(B[0] * I, ",").concat(B[1] * L);
                    })
                    .join(" ")
                : "";
            return O ? "".concat(D, " ").concat(O) : "".concat(D);
          }).join(" ");
        return (k.setAttribute("d", E), (k.style.fillRule = "evenodd"), k);
      })(e)),
        Jg(e, d, f));
      break;
    case "flowChartProcess":
    case "rect":
      f = _W(e);
      break;
    case "snip1Rect":
      f = (function (m) {
        var v = m.extend,
          b = lt("polygon"),
          _ = at("adj", m, jl(m)),
          S = [
            "0,0",
            "".concat(v.w - _, ",0"),
            "".concat(v.w, ",").concat(_),
            "".concat(v.w, ",").concat(v.h),
            "0,".concat(v.h),
          ].join(" ");
        return (b.setAttribute("points", S), b);
      })(e);
      break;
    case "snip2SameRect":
      f = (function (m) {
        var v = m.extend,
          b = lt("polygon"),
          _ = at("adj1", m, jl(m)),
          S = at("adj2", m, 0),
          C = [
            [_, 0],
            [v.w - _, 0],
            [v.w, _],
            [v.w, v.h - S],
            [v.w - S, v.h],
            [S, v.h],
            [0, v.h - S],
            [0, _],
          ]
            .map(function (T) {
              return "".concat(T[0], ",").concat(T[1]);
            })
            .join(" ");
        return (b.setAttribute("points", C), b);
      })(e);
      break;
    case "snip2DiagRect":
      f = (function (m) {
        var v = m.extend,
          b = lt("polygon"),
          _ = at("adj1", m, 0),
          S = at("adj2", m, jl(m)),
          C = [
            [_, 0],
            [v.w - S, 0],
            [v.w, S],
            [v.w, v.h - _],
            [v.w - _, v.h],
            [S, v.h],
            [0, v.h - S],
            [0, _],
          ]
            .map(function (T) {
              return "".concat(T[0], ",").concat(T[1]);
            })
            .join(" ");
        return (b.setAttribute("points", C), b);
      })(e);
      break;
    case "snipRoundRect":
      f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, jl(m)),
          S = at("adj2", m, jl(m)),
          C = [
            "M ".concat(_, ",0"),
            "L ".concat(v.w - S, ",0"),
            "L ".concat(v.w, ",").concat(S),
            "L ".concat(v.w, ",").concat(v.h),
            "L 0,".concat(v.h),
            "L 0,".concat(_),
            "Q 0,0 ".concat(_, ",0"),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e);
      break;
    case "roundRect":
      f = (function (m) {
        var v = _W(m),
          b = at("adj", m, jl(m));
        return (
          v.setAttribute("rx", b + "px"),
          v.setAttribute("ry", b + "px"),
          v
        );
      })(e);
      break;
    case "wedgeRoundRectCallout":
      // Rounded body plus the wedge tail that points at the anchor.  Both
      // adjustments are measured in short-side units (ss = min(w, h)):
      //   tipX (from the left edge) = ss * adj1
      //   tipY (from the top)       = h / 2 + ss * adj2
      // and the tail's base sits on the left edge around tipY.  The golden set
      // shows the tail on slides 6 and 7 (adj1 = -89815 puts the tip 31 px left
      // of the body), so leaving it out loses the pointer to the red line.
      f = (function (m) {
        var v = m.extend,
          b = lt("g"),
          r3 = at("adj3", m, jl(m)),
          tipX = at("adj1", m),
          tipY = v.h / 2 + at("adj2", m),
          half = r3 / 2,
          body = _W(m),
          tail = lt("polygon");
        (body.setAttribute("rx", r3 + "px"),
          body.setAttribute("ry", r3 + "px"),
          tail.setAttribute(
            "points",
            "0," + (tipY - half) + " " + tipX + "," + tipY + " 0," + (tipY + half),
          ),
          b.appendChild(body),
          b.appendChild(tail));
        return b;
      })(e);
      break;
    case "round1Rect":
      f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, jl(m)),
          S = [
            "M 0,0",
            "L ".concat(v.w - _, ",0"),
            "Q ".concat(v.w, ",0 ").concat(v.w, ",").concat(_),
            "L ".concat(v.w, ",").concat(v.h),
            "L 0,".concat(v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e);
      break;
    case "round2SameRect":
      f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, jl(m)),
          S = at("adj2", m, 0),
          C = [
            "M ".concat(_, ",0"),
            "L ".concat(v.w - _, ",0"),
            "Q ".concat(v.w, ",0 ").concat(v.w, ",").concat(_),
            "L ".concat(v.w, ",").concat(v.h - S),
            "Q "
              .concat(v.w, ",")
              .concat(v.h, " ")
              .concat(v.w - S, ",")
              .concat(v.h),
            "L ".concat(S, ",").concat(v.h),
            "Q 0,".concat(v.h, " 0,").concat(v.h - S),
            "L 0,".concat(_),
            "Q 0,0 ".concat(_, ",0"),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e);
      break;
    case "round2DiagRect":
      f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, jl(m)),
          S = at("adj2", m, 0),
          C = [
            "M ".concat(_, ",0"),
            "L ".concat(v.w - S, ",0"),
            "Q ".concat(v.w, ",0 ").concat(v.w, ",").concat(S),
            "L ".concat(v.w, ",").concat(v.h - _),
            "Q "
              .concat(v.w, ",")
              .concat(v.h, " ")
              .concat(v.w - _, ",")
              .concat(v.h),
            "L ".concat(S, ",").concat(v.h),
            "Q 0,".concat(v.h, " 0,").concat(v.h - S),
            "L 0,".concat(_),
            "Q 0,0 ".concat(_, ",0"),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e);
      break;
    case "triangle":
      f = (function (m) {
        var v = m.extend,
          b = lt("polygon"),
          _ = [
            "".concat(v.w / 2, ",0"),
            "0,".concat(v.h),
            "".concat(v.w, ",").concat(v.h),
          ].join(" ");
        return (b.setAttribute("points", _), b);
      })(e);
      break;
    case "rtTriangle":
      f = (function (m) {
        var v = m.extend,
          b = lt("polygon"),
          _ = ["0,".concat(v.h), "0,0", "".concat(v.w, ",").concat(v.h)].join(
            " ",
          );
        return (b.setAttribute("points", _), b);
      })(e);
      break;
    case "ellipse":
      f = (function (m) {
        var v = m.extend,
          b = lt("ellipse"),
          _ = v.w / 2,
          S = v.h / 2,
          C = _,
          T = S;
        return (
          b.setAttribute("cx", _ + "px"),
          b.setAttribute("cy", S + "px"),
          b.setAttribute("rx", C + "px"),
          b.setAttribute("ry", T + "px"),
          b
        );
      })(e);
      break;
    case "line":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = ["M 0,0", "L ".concat(v.w, ",").concat(v.h)].join(" ");
        return (b.setAttribute("d", _), b);
      })(e)),
        Jg(e, d, f),
        (p = !1));
      break;
    case "straightConnector1":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = ["M 0,0", "L ".concat(v.w, ",").concat(v.h)].join(" ");
        return (b.setAttribute("d", _), b);
      })(e)),
        Jg(e, d, f),
        (p = !1));
      break;
    case "bentConnector3":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = vW("adj1", m, 0.5 * Math.max(v.w, v.h)),
          S = [
            "M 0,0",
            "L ".concat(_, ",0"),
            "L ".concat(_, ",").concat(v.h),
            "L ".concat(v.w, ",").concat(v.h),
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        Jg(e, d, f),
        (p = !1));
      break;
    case "curvedConnector3":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = vW("adj1", m, 0.5 * Math.max(v.w, v.h)),
          S = [
            "M0,0",
            "Q"
              .concat(_, ",0 ")
              .concat(_, ",")
              .concat(v.h / 2),
            "T".concat(v.w, ",").concat(v.h),
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        Jg(e, d, f),
        (p = !1));
      break;
    case "parallelogram":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.5 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w, ",0"),
            "L".concat(v.w - _, ",").concat(v.h),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.5 * Math.min(v.w, v.h)),
            _ = 0.84 * (v.w - b),
            S = 0.08 * v.h + (b / v.w) * v.h * 0.42;
          return {
            top: S,
            bottom: S,
            left: (v.w - _) / 2,
            right: (v.w - _) / 2,
            w: _,
            h: v.h - 2 * S,
          };
        })(e)));
      break;
    case "trapezoid":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(v.h),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b / v.w) * 0.66 * v.h,
            S = 0.66 * b;
          return {
            top: _,
            bottom: 0,
            left: S,
            right: S,
            w: v.w - 2 * S,
            h: v.h - _,
          };
        })(e)));
      break;
    case "diamond":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = [
            "M".concat(v.w / 2, ",0"),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w / 2, ",").concat(v.h),
            "L0,".concat(v.h / 2),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", _), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.25 * v.h,
            _ = 0.25 * v.w;
          return {
            top: b,
            bottom: b,
            left: _,
            right: _,
            w: 0.5 * v.w,
            h: 0.5 * v.h,
          };
        })(e)));
      break;
    case "pentagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = [
            "M".concat(v.w / 2, ",0"),
            "L".concat(v.w, ",").concat(0.3771 * v.h),
            "L".concat(0.808 * v.w, ",").concat(v.h),
            "L".concat(0.192 * v.w, ",").concat(v.h),
            "L0,".concat(0.3771 * v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", _), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.227 * v.h,
            _ = 0.192 * v.w;
          return {
            top: b,
            bottom: 0,
            left: _,
            right: _,
            w: 0.616 * v.w,
            h: 0.773 * v.h,
          };
        })(e)));
      break;
    case "hexagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h),
            "L0,".concat(v.h / 2),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.25 * Math.min(v.w, v.h)),
            _ = 0.098 * v.h + (b / v.w) * 0.38 * v.h,
            S = 0.088 * v.w + 0.422 * b;
          return {
            top: _,
            bottom: _,
            left: S,
            right: S,
            w: v.w - 2 * S,
            h: v.h - 2 * _,
          };
        })(e)));
      break;
    case "heptagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = [
            "M".concat(v.w / 2, ",0"),
            "L".concat(0.9 * v.w, ",").concat(0.2 * v.h),
            "L".concat(v.w, ",").concat(0.642 * v.h),
            "L".concat(0.722 * v.w, ",").concat(v.h),
            "L".concat(0.278 * v.w, ",").concat(v.h),
            "L0,".concat(0.642 * v.h),
            "L".concat(0.1 * v.w, ",").concat(0.2 * v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", _), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend;
          return {
            top: 0.2 * v.h,
            bottom: 0.2 * v.h,
            left: 0.1 * v.w,
            right: 0.1 * v.w,
            w: 0.8 * v.w,
            h: 0.6 * v.h,
          };
        })(e)));
      break;
    case "octagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.29 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(_),
            "L".concat(v.w, ",").concat(v.h - _),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h),
            "L0,".concat(v.h - _),
            "L0,".concat(_),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.29 * Math.min(v.w, v.h));
          return {
            top: 0.5 * b,
            bottom: 0.5 * b,
            left: 0.5 * b,
            right: 0.5 * b,
            w: v.w - b,
            h: v.h - b,
          };
        })(e)));
      break;
    case "decagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = 0.344,
          S = 0.117,
          C = 0.19,
          T = [
            "M".concat(v.w * _, ",0"),
            "L".concat(0.656 * v.w, ",0"),
            "L".concat(0.883 * v.w, ",").concat(v.h * C),
            "L".concat(v.w, ",").concat(0.5 * v.h),
            "L".concat(0.883 * v.w, ",").concat(0.81 * v.h),
            "L".concat(0.656 * v.w, ",").concat(v.h),
            "L".concat(v.w * _, ",").concat(v.h),
            "L".concat(v.w * S, ",").concat(0.81 * v.h),
            "L0,".concat(0.5 * v.h),
            "L".concat(v.w * S, ",").concat(v.h * C),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", T), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.117,
            _ = 0.19;
          return {
            top: v.h * _,
            bottom: v.h * _,
            left: v.w * b,
            right: v.w * b,
            w: 0.766 * v.w,
            h: 0.62 * v.h,
          };
        })(e)));
      break;
    case "dodecagon":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = 0.364,
          S = 0.133,
          C = 0.135,
          T = [
            "M".concat(v.w * _, ",0"),
            "L".concat(0.636 * v.w, ",0"),
            "L".concat(0.867 * v.w, ",").concat(v.h * C),
            "L".concat(v.w, ",").concat(v.h * _),
            "L".concat(v.w, ",").concat(0.636 * v.h),
            "L".concat(0.867 * v.w, ",").concat(0.865 * v.h),
            "L".concat(0.636 * v.w, ",").concat(v.h),
            "L".concat(v.w * _, ",").concat(v.h),
            "L".concat(v.w * S, ",").concat(0.865 * v.h),
            "L0,".concat(0.636 * v.h),
            "L0,".concat(v.h * _),
            "L".concat(v.w * S, ",").concat(v.h * C),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", T), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.133,
            _ = 0.135;
          return {
            top: v.h * _,
            bottom: v.h * _,
            left: v.w * b,
            right: v.w * b,
            w: 0.734 * v.w,
            h: 0.73 * v.h,
          };
        })(e)));
      break;
    case "pie":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = rn("adj1", m, 360),
          S = rn("adj2", m, 270),
          C = v.w / 2,
          T = v.h / 2,
          k = v.w / 2,
          M = v.h / 2,
          I = Wa(_, k, M, C, T),
          L = I[0],
          E = I[1],
          R = Wa(S, k, M, C, T),
          D = R[0],
          O = R[1],
          B = Qg(_, S),
          V = "M"
            .concat(k, ",")
            .concat(M, ", L")
            .concat(L, " ")
            .concat(E, " A ")
            .concat(C, " ")
            .concat(T, " 0 ")
            .concat(B, " ")
            .concat(1, " ")
            .concat(D, " ")
            .concat(O, " Z");
        return (b.setAttribute("d", V), b);
      })(e)),
        (h = em(e)));
      break;
    case "arc":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = rn("adj1", m, 270),
          S = rn("adj2", m, 0),
          C = v.w / 2,
          T = v.h / 2,
          k = v.w / 2,
          M = v.h / 2,
          I = Wa(_, k, M, C, T),
          L = I[0],
          E = I[1],
          R = Wa(S, k, M, C, T),
          D = R[0],
          O = R[1],
          B = Qg(_, S),
          V = "M"
            .concat(L, ",")
            .concat(E, " A ")
            .concat(C, " ")
            .concat(T, " 0 ")
            .concat(B, " ")
            .concat(1, " ")
            .concat(D, " ")
            .concat(O);
        return (b.setAttribute("d", V), b);
      })(e)),
        (p = !1));
      break;
    case "bracketPair":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.16667 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",").concat(v.h),
            "Q"
              .concat(0, ",")
              .concat(v.h, " ")
              .concat(0, ",")
              .concat(v.h - _),
            "L".concat(0, ",").concat(_),
            "Q".concat(0, ",").concat(0, " ").concat(_, ",").concat(0),
            "M".concat(v.w - _, ",").concat(0),
            "Q".concat(v.w, ",").concat(0, " ").concat(v.w, ",").concat(_),
            "L".concat(v.w, ",").concat(v.h - _),
            "Q"
              .concat(v.w, ",")
              .concat(v.h, " ")
              .concat(v.w - _, ",")
              .concat(v.h),
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.285 * at("adj", m, 0.16667 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: b,
            left: b,
            right: b,
            w: v.w - 2 * b,
            h: v.h - 2 * b,
          };
        })(e)),
        (p = !1));
      break;
    case "bracePair":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.083335 * Math.min(v.w, v.h)),
          S = [
            "M".concat(2 * _, ",").concat(v.h),
            "Q"
              .concat(_, ",")
              .concat(v.h, " ")
              .concat(_, ",")
              .concat(v.h - _),
            "L".concat(_, ",").concat(v.h / 2 + _),
            "Q"
              .concat(_, ",")
              .concat(v.h / 2, " ")
              .concat(0, ",")
              .concat(v.h / 2),
            "Q"
              .concat(_, ",")
              .concat(v.h / 2, " ")
              .concat(_, ",")
              .concat(v.h / 2 - _),
            "L".concat(_, ",").concat(_),
            "Q"
              .concat(_, ",")
              .concat(0, " ")
              .concat(2 * _, ",")
              .concat(0),
            "M".concat(v.w - 2 * _, ",").concat(0),
            "Q"
              .concat(v.w - _, ",")
              .concat(0, " ")
              .concat(v.w - _, ",")
              .concat(_),
            "L".concat(v.w - _, ",").concat(v.h / 2 - _),
            "Q"
              .concat(v.w - _, ",")
              .concat(v.h / 2, " ")
              .concat(v.w, ",")
              .concat(v.h / 2),
            "Q"
              .concat(v.w - _, ",")
              .concat(v.h / 2, " ")
              .concat(v.w - _, ",")
              .concat(v.h / 2 + _),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "Q"
              .concat(v.w - _, ",")
              .concat(v.h, " ")
              .concat(v.w - 2 * _, ",")
              .concat(v.h),
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = 0.285 * at("adj", m, 0.16667 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: b,
            left: b,
            right: b,
            w: v.w - 2 * b,
            h: v.h - 2 * b,
          };
        })(e)),
        (p = !1));
      break;
    case "chord":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = rn("adj1", m, 45),
          S = rn("adj2", m, 270),
          C = v.w / 2,
          T = v.h / 2,
          k = v.w / 2,
          M = v.h / 2,
          I = Wa(_, k, M, C, T),
          L = I[0],
          E = I[1],
          R = Wa(S, k, M, C, T),
          D = R[0],
          O = R[1],
          B = Qg(_, S),
          V = "M"
            .concat(L, " ")
            .concat(E, " A ")
            .concat(C, " ")
            .concat(T, " 0 ")
            .concat(B, " ")
            .concat(1, " ")
            .concat(D, " ")
            .concat(O, " Z");
        return (b.setAttribute("d", V), b);
      })(e)),
        (h = em(e)));
      break;
    case "teardrop":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = rn("adj", m, 1),
          S = v.w / 2,
          C = v.h / 2,
          T = v.w / 2,
          k = v.h / 2,
          M = Wa(0, T, k, S, C),
          I = M[0],
          L = M[1],
          E = Wa(270, T, k, S, C),
          R = E[0],
          D = E[1],
          O = Qg(0, 270),
          B = "M"
            .concat(I, " ")
            .concat(L, " A ")
            .concat(S, " ")
            .concat(C, " 0 ")
            .concat(O, " ")
            .concat(1, " ")
            .concat(R, " ")
            .concat(D),
          V = S * _,
          z = T + V,
          N = k - (C * V) / (v.w / 2),
          H = (v.w / 2 + z) / 2,
          G = (v.h / 2 + N) / 2;
        return (
          (B += " Q".concat(H, ",0 ").concat(z, ",").concat(N)),
          (B += " Q"
            .concat(v.w, ",")
            .concat(G, " ")
            .concat(T + S, ",")
            .concat(k)),
          b.setAttribute("d", B),
          b
        );
      })(e)),
        (h = em(e)));
      break;
    case "frame":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, 0.12 * Math.min(v.w, v.h)),
          S = [
            "M0,0",
            "L".concat(v.w, ",0"),
            "L".concat(v.w, ",").concat(v.h),
            "L0,".concat(v.h),
            "Z",
            "M".concat(_, ",").concat(_),
            "L".concat(_, ",").concat(v.h - _),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "L".concat(v.w - _, ",").concat(_),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.12 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: b,
            left: b,
            right: b,
            w: v.w - 2 * b,
            h: v.h - 2 * b,
          };
        })(e)));
      break;
    case "halfFrame":
      f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, 0.333 * Math.min(v.w, v.h)),
          S = Math.min(
            at("adj2", m, 0.333 * Math.min(v.w, v.h)),
            v.w * (1 - _ / v.h),
          ),
          C = [
            "M0,0",
            "L".concat(v.w, ",0"),
            "L".concat(v.w * (1 - _ / v.h), ",").concat(_),
            "L".concat(S, ",").concat(_),
            "L".concat(S, ",").concat(v.h * (1 - S / v.w)),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e);
      break;
    case "corner":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          S = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          C = [
            "M0,0",
            "L".concat(S, ",0"),
            "L".concat(S, ",").concat(v.h - _),
            "L".concat(v.w, ",").concat(v.h - _),
            "L".concat(v.w, ",").concat(v.h),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h));
          return { top: v.h - b, bottom: 0, left: 0, right: 0, w: v.w, h: b };
        })(e)));
      break;
    case "diagStripe":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.5 * Math.min(v.w, v.h)),
          S = (v.w * _) / v.h,
          C = [
            "M".concat(S, ",0"),
            "L".concat(v.w, ",0"),
            "L0,".concat(v.h),
            "L0,".concat(_),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", C), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj", m, 0.5),
            _ = 0.5 * (1 - b) * v.h,
            S = 0.5 * (1 - b) * v.w;
          return {
            top: 0,
            bottom: _,
            left: 0,
            right: S,
            w: v.w - S,
            h: v.h - _,
          };
        })(e)));
      break;
    case "plus":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w, ",").concat(_),
            "L".concat(v.w, ",").concat(v.h - _),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h - _),
            "L0,".concat(v.h - _),
            "L0,".concat(_),
            "L".concat(_, ",").concat(_),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.25 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: b,
            left: 0,
            right: 0,
            w: v.w,
            h: v.h - 2 * b,
          };
        })(e)));
      break;
    case "plaque":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("path"),
          _ = at("adj", m, 0.16667 * Math.min(v.w, v.h)),
          S = [
            "M".concat(_, ",0"),
            "L".concat(v.w - _, ",0"),
            "Q"
              .concat(v.w - _, ",")
              .concat(_, " ")
              .concat(v.w, ",")
              .concat(_),
            "L".concat(v.w, ",").concat(v.h - _),
            "Q"
              .concat(v.w - _, ",")
              .concat(v.h - _, " ")
              .concat(v.w - _, ",")
              .concat(v.h),
            "L".concat(_, ",").concat(v.h),
            "Q"
              .concat(_, ",")
              .concat(v.h - _, " 0,")
              .concat(v.h - _),
            "L0,".concat(_),
            "Q".concat(_, ",").concat(_, " ").concat(_, ",0"),
            "Z",
          ].join(" ");
        return (b.setAttribute("d", S), b);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.16667 * Math.min(v.w, v.h));
          return {
            top: 0.72 * b,
            bottom: 0.72 * b,
            left: 0.72 * b,
            right: 0.72 * b,
            w: v.w - 1.44 * b,
            h: v.w - 1.44 * b,
          };
        })(e)));
      break;
    case "can":
      ((f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = lt("g"),
          S = lt("path"),
          C = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          T = [
            "M0,".concat(C / 2),
            "L0,".concat(v.h - C / 2),
            "A"
              .concat(v.w / 2, ",")
              .concat(C / 2, " 0 0 0 ")
              .concat(v.w, ",")
              .concat(v.h - C / 2),
            "L".concat(v.w, ",").concat(C / 2),
            "A"
              .concat(v.w / 2, ",")
              .concat(C / 2, " 0 0 1 0,")
              .concat(C / 2),
            "Z",
          ].join(" ");
        S.setAttribute("d", T);
        var k = lt("ellipse"),
          M = v.w / 2,
          I = C / 2,
          L = v.w / 2,
          E = C / 2;
        return (
          k.setAttribute("cx", M + "px"),
          k.setAttribute("cy", I + "px"),
          k.setAttribute("rx", L + "px"),
          k.setAttribute("ry", E + "px"),
          (b == null ? void 0 : b.type) === "solidFill" &&
            k.setAttribute("fill", Qn(b, { light: 0.5 }) || "transparent"),
          _.appendChild(S),
          _.appendChild(k),
          _
        );
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.25 * Math.min(v.w, v.h));
          return { top: b, bottom: 0, left: 0, right: 0, w: v.w, h: v.h - b };
        })(e)));
      break;
    case "cube":
      ((f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          S = lt("g"),
          C = lt("path"),
          T = [
            "M0,".concat(_),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w - _, ",").concat(v.h),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        C.setAttribute("d", T);
        var k = lt("path"),
          M = [
            "M0,".concat(_),
            "L".concat(_, ",0"),
            "L".concat(v.w, ",0"),
            "L".concat(v.w - _, ",").concat(_),
            "Z",
          ].join(" ");
        (k.setAttribute("d", M),
          (b == null ? void 0 : b.type) === "solidFill" &&
            k.setAttribute("fill", Qn(b, { light: 0.8 }) || "transparent"));
        var I = lt("path"),
          L = [
            "M".concat(v.w, ",0"),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(v.w, ",").concat(v.h - _),
            "Z",
          ].join(" ");
        return (
          I.setAttribute("d", L),
          (b == null ? void 0 : b.type) === "solidFill" &&
            I.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"),
          S.appendChild(C),
          S.appendChild(k),
          S.appendChild(I),
          S
        );
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.25 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: 0,
            left: 0,
            right: b,
            w: v.w - b,
            h: v.h - b,
          };
        })(e)));
      break;
    case "bevel":
      ((f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj", m, 0.125 * Math.min(v.w, v.h)),
          S = lt("g"),
          C = lt("path"),
          T = [
            "M".concat(_, ",").concat(_),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "L".concat(_, ",").concat(v.h - _),
            "Z",
          ].join(" ");
        C.setAttribute("d", T);
        var k = lt("path"),
          M = [
            "M0,0",
            "L".concat(_, ",").concat(_),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w, ",0"),
            "Z",
          ].join(" ");
        (k.setAttribute("d", M),
          (b == null ? void 0 : b.type) === "solidFill" &&
            k.setAttribute("fill", Qn(b, { light: 0.8 }) || "transparent"));
        var I = lt("path"),
          L = [
            "M".concat(v.w, ",0"),
            "L".concat(v.w - _, ",").concat(_),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "L".concat(v.w, ",").concat(v.h),
            "Z",
          ].join(" ");
        (I.setAttribute("d", L),
          (b == null ? void 0 : b.type) === "solidFill" &&
            I.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"));
        var E = lt("path"),
          R = [
            "M".concat(v.w, ",").concat(v.h),
            "L".concat(v.w - _, ",").concat(v.h - _),
            "L".concat(_, ",").concat(v.h - _),
            "L0,".concat(v.h),
            "Z",
          ].join(" ");
        (E.setAttribute("d", R),
          (b == null ? void 0 : b.type) === "solidFill" &&
            E.setAttribute("fill", Qn(b, { dark: 0.625 }) || "transparent"));
        var D = lt("path"),
          O = [
            "M0,".concat(v.h),
            "L".concat(_, ",").concat(v.h - _),
            "L".concat(_, ",").concat(_),
            "L0,0",
            "Z",
          ].join(" ");
        return (
          D.setAttribute("d", O),
          (b == null ? void 0 : b.type) === "solidFill" &&
            D.setAttribute("fill", Qn(b, { light: 0.6 }) || "transparent"),
          S.appendChild(C),
          S.appendChild(k),
          S.appendChild(I),
          S.appendChild(E),
          S.appendChild(D),
          S
        );
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.125 * Math.min(v.w, v.h));
          return {
            top: b,
            bottom: b,
            left: b,
            right: b,
            w: v.w - 2 * b,
            h: v.h - 2 * b,
          };
        })(e)));
      break;
    case "donut":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          _ = lt("path"),
          S = [
            "M0,".concat(v.h / 2),
            "A"
              .concat(v.w / 2, ",")
              .concat(v.h / 2, " 0 1,1 0,")
              .concat(v.h / 2 + 1),
            "Z",
            "M".concat(v.w - b, ",").concat(v.h / 2),
            "A"
              .concat(v.w / 2 - b, ",")
              .concat(v.h / 2 - b, " 0 1,0 ")
              .concat(v.w - b, ",")
              .concat(v.h / 2 + 1),
            "Z",
          ].join(" ");
        return (_.setAttribute("d", S), _);
      })(e)),
        (h = em(e)));
      break;
    case "noSmoking":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj", m, 0.25 * Math.min(v.w, v.h)),
          _ = lt("path"),
          S = Math.atan(v.h / v.w),
          C = b / 2 / Math.sin(S),
          T = v.w / 2,
          k = v.h / 2,
          M = (-1 * v.h) / v.w,
          I = (v.h * C) / v.w,
          L = v.w / 2 - b,
          E = v.h / 2 - b,
          R = -2 * L * L * M * I,
          D = Math.sqrt(
            Math.pow(2 * L * L * M * I, 2) -
              4 * (E * E + L * L * M * M) * L * L * (I * I - E * E),
          ),
          O = 2 * (E * E + L * L * M * M),
          B = (R - D) / O,
          V = M * B + I,
          z = (R + D) / O,
          N = M * z + I,
          H = (-v.h * C) / v.w,
          G = -2 * L * L * M * H,
          j = Math.sqrt(
            Math.pow(2 * L * L * M * H, 2) -
              4 * (E * E + L * L * M * M) * L * L * (H * H - E * E),
          ),
          J = 2 * (E * E + L * L * M * M),
          re = (G - j) / J,
          le = M * re + H,
          Te = (G + j) / J,
          fe = M * Te + H,
          _e = [
            "M0,".concat(v.h / 2),
            "A"
              .concat(v.w / 2, ",")
              .concat(v.h / 2, " 0 1,1 0,")
              .concat(v.h / 2 + 1),
            "Z",
            "M".concat(T + z, ",").concat(k - N),
            "A"
              .concat(L, ",")
              .concat(E, " 0 0 0 ")
              .concat(T + B, ",")
              .concat(k - V),
            "Z",
            "M".concat(T + re, ",").concat(k - le),
            "A"
              .concat(L, ",")
              .concat(E, " 0 0 0 ")
              .concat(T + Te, ",")
              .concat(k - fe),
            "Z",
          ].join(" ");
        return (_.setAttribute("d", _e), _);
      })(e)),
        (h = em(e)));
      break;
    case "rightArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M0,".concat(v.h / 2 - b / 2),
            "L".concat(v.w - _, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(v.w - _, ",").concat(v.h / 2 + b / 2),
            "L0,".concat(v.h / 2 + b / 2),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.h;
          return {
            top: v.h / 2 - b / 2,
            bottom: v.h / 2 - b / 2,
            left: 0,
            right: _,
            w: v.w - _,
            h: b,
          };
        })(e)));
      break;
    case "leftArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M0,".concat(v.h / 2),
            "L".concat(_, ",0"),
            "L".concat(_, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w, ",").concat(v.h / 2 + b / 2),
            "L".concat(_, ",").concat(v.h / 2 + b / 2),
            "L".concat(_, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.h;
          return {
            top: v.h / 2 - b / 2,
            bottom: v.h / 2 - b / 2,
            left: _,
            right: 0,
            w: v.w - _,
            h: b,
          };
        })(e)));
      break;
    case "upArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M".concat(v.w / 2, ",0"),
            "L".concat(v.w, ",").concat(_),
            "L".concat(v.w / 2 + b / 2, ",").concat(_),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - b / 2, ",").concat(_),
            "L0,".concat(_),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.w;
          return {
            top: _,
            bottom: 0,
            left: v.w / 2 - b / 2,
            right: v.w / 2 - b / 2,
            w: b,
            h: v.h - _,
          };
        })(e)));
      break;
    case "downArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M".concat(v.w / 2, ",").concat(v.h),
            "L0,".concat(v.h - _),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h - _),
            "L".concat(v.w / 2 - b / 2, ",0"),
            "L".concat(v.w / 2 + b / 2, ",0"),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h - _),
            "L".concat(v.w, ",").concat(v.h - _),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.w;
          return {
            top: 0,
            bottom: _,
            left: v.w / 2 - b / 2,
            right: v.w / 2 - b / 2,
            w: b,
            h: v.h - _,
          };
        })(e)));
      break;
    case "leftRightArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M0,".concat(v.h / 2),
            "L".concat(_, ",0"),
            "L".concat(_, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w - _, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(v.w - _, ",").concat(v.h / 2 + b / 2),
            "L".concat(_, ",").concat(v.h / 2 + b / 2),
            "L".concat(_, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.h;
          return {
            top: v.h / 2 - b / 2,
            bottom: v.h / 2 - b / 2,
            left: _,
            right: _,
            w: v.w - 2 * _,
            h: b,
          };
        })(e)));
      break;
    case "upDownArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          S = lt("path"),
          C = [
            "M".concat(v.w / 2, ",").concat(v.h),
            "L0,".concat(v.h - _),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h - _),
            "L".concat(v.w / 2 - b / 2, ",").concat(_),
            "L0,".concat(_),
            "L".concat(v.w / 2, ",0"),
            "L".concat(v.w, ",").concat(_),
            "L".concat(v.w / 2 + b / 2, ",").concat(_),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h - _),
            "L".concat(v.w, ",").concat(v.h - _),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.5 * Math.min(v.w, v.h)),
            _ = (b * at("adj2", m, 0.5 * Math.min(v.w, v.h))) / v.w;
          return {
            top: _,
            bottom: _,
            left: v.w / 2 - b / 2,
            right: v.w / 2 - b / 2,
            w: b,
            h: v.h - 2 * _,
          };
        })(e)));
      break;
    case "quadArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.225 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.225 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.225 * Math.min(v.w, v.h)),
          C = lt("path"),
          T = [
            "M0,".concat(v.h / 2),
            "L".concat(S, ",").concat(v.h / 2 - _),
            "L".concat(S, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w / 2 - b / 2, ",").concat(S),
            "L".concat(v.w / 2 - _, ",").concat(S),
            "L".concat(v.w / 2, ",0"),
            "L".concat(v.w / 2 + _, ",").concat(S),
            "L".concat(v.w / 2 + b / 2, ",").concat(S),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w - S, ",").concat(v.h / 2 - b / 2),
            "L".concat(v.w - S, ",").concat(v.h / 2 - _),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - S, ",").concat(v.h / 2 + _),
            "L".concat(v.w - S, ",").concat(v.h / 2 + b / 2),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h / 2 + b / 2),
            "L".concat(v.w / 2 + b / 2, ",").concat(v.h - S),
            "L".concat(v.w / 2 + _, ",").concat(v.h - S),
            "L".concat(v.w / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - _, ",").concat(v.h - S),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h - S),
            "L".concat(v.w / 2 - b / 2, ",").concat(v.h / 2 + b / 2),
            "L".concat(S, ",").concat(v.h / 2 + b / 2),
            "L".concat(S, ",").concat(v.h / 2 + _),
            "Z",
          ].join(" ");
        return (C.setAttribute("d", T), C);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.225 * Math.min(v.w, v.h)),
            _ = at("adj2", m, 0.225 * Math.min(v.w, v.h)),
            S = at("adj3", m, 0.225 * Math.min(v.w, v.h)),
            C = _ === 0 ? 0 : (b * S) / _ / 2;
          return {
            top: v.h / 2 - b / 2,
            bottom: v.h / 2 - b / 2,
            left: C,
            right: C,
            w: v.w - 2 * C,
            h: b,
          };
        })(e)));
      break;
    case "leftRightUpArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.225 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.225 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.225 * Math.min(v.w, v.h)),
          C = lt("path");
        b > 2 * _ && (b = 2 * _);
        var T = [
          "M0,".concat(v.h - _),
          "L".concat(S, ",").concat(v.h - 2 * _),
          "L".concat(S, ",").concat(v.h - _ - b / 2),
          "L".concat(v.w / 2 - b / 2, ",").concat(v.h - _ - b / 2),
          "L".concat(v.w / 2 - b / 2, ",").concat(S),
          "L".concat(v.w / 2 - _, ",").concat(S),
          "L".concat(v.w / 2, ",0"),
          "L".concat(v.w / 2 + _, ",").concat(S),
          "L".concat(v.w / 2 + b / 2, ",").concat(S),
          "L".concat(v.w / 2 + b / 2, ",").concat(v.h - _ - b / 2),
          "L".concat(v.w - S, ",").concat(v.h - _ - b / 2),
          "L".concat(v.w - S, ",").concat(v.h - 2 * _),
          "L".concat(v.w, ",").concat(v.h - _),
          "L".concat(v.w - S, ",").concat(v.h),
          "L".concat(v.w - S, ",").concat(v.h - _ + b / 2),
          "L".concat(S, ",").concat(v.h - _ + b / 2),
          "L".concat(S, ",").concat(v.h),
          "Z",
        ].join(" ");
        return (C.setAttribute("d", T), C);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.225 * Math.min(v.w, v.h)),
            _ = at("adj2", m, 0.225 * Math.min(v.w, v.h)),
            S = at("adj3", m, 0.225 * Math.min(v.w, v.h)),
            C = _ === 0 ? 0 : (b * S) / _ / 2;
          return {
            top: v.h - _ - b / 2,
            bottom: _ - b / 2,
            left: C,
            right: C,
            w: v.w - 2 * C,
            h: b,
          };
        })(e)));
      break;
    case "bentArrow":
      f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.25 * Math.min(v.w, v.h)),
          C = at("adj4", m, 0.4375 * Math.min(v.w, v.h)),
          T = lt("path");
        b > 2 * _ && (b = 2 * _);
        var k = C - b;
        k < 0 && (k = 0);
        var M = [
          "M0,".concat(v.h),
          "L0,".concat(_ - b / 2 + C),
          "A"
            .concat(C, " ")
            .concat(C, " 0 0 1 ")
            .concat(C, " ")
            .concat(_ - b / 2),
          "L".concat(v.w - S, ",").concat(_ - b / 2),
          "L".concat(v.w - S, ",0"),
          "L".concat(v.w, ",").concat(_),
          "L".concat(v.w - S, ",").concat(2 * _),
          "L".concat(v.w - S, ",").concat(_ + b / 2),
          "L".concat(b + k, ",").concat(_ + b / 2),
          "A"
            .concat(k, " ")
            .concat(k, "  0 0 0 ")
            .concat(b, " ")
            .concat(_ + b / 2 + k),
          "L".concat(b, ",").concat(v.h),
          "Z",
        ].join(" ");
        return (T.setAttribute("d", M), T);
      })(e);
      break;
    case "uturnArrow":
      f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.25 * Math.min(v.w, v.h)),
          C = at("adj4", m, 0.4375 * Math.min(v.w, v.h)),
          T = at("adj5", m, 0.75 * Math.min(v.w, v.h)),
          k = lt("path");
        (b > 2 * _ && (b = 2 * _),
          T < S && (T = S + b),
          C > T - S && (C = T - S));
        var M = C - b;
        (M > T - S - b && (M = T - S - b), M < 0 && (M = 0));
        var I = _ - b / 2,
          L = [
            "M0,".concat(v.h),
            "L0,".concat(C),
            "A".concat(C, " ").concat(C, " 0 0 1 ").concat(C, " 0"),
            "L".concat(v.w - C - I, ",0"),
            "A"
              .concat(C, " ")
              .concat(C, " 0 0 1 ")
              .concat(v.w - I, " ")
              .concat(C),
            "L".concat(v.w - I, ",").concat(T - S),
            "L".concat(v.w, ",").concat(T - S),
            "L".concat(v.w - _, ",").concat(T),
            "L".concat(v.w - 2 * _, ",").concat(T - S),
            "L".concat(v.w - _ - b / 2, ",").concat(T - S),
            "L".concat(v.w - _ - b / 2, ",").concat(b + M),
            "A"
              .concat(M, " ")
              .concat(M, "  0 0 0 ")
              .concat(v.w - M - _ - b / 2, " ")
              .concat(b),
            "L".concat(b + M, ",").concat(b),
            "A"
              .concat(M, " ")
              .concat(M, "  0 0 0 ")
              .concat(b, " ")
              .concat(b + M),
            "L".concat(b, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", L), k);
      })(e);
      break;
    case "leftUpArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        b > 2 * _ && (b = 2 * _);
        var C = Math.min(v.w, v.h) - 2 * _;
        (S > C && (S = C), S < 0 && (S = 0));
        var T = _ - b / 2,
          k = lt("path"),
          M = [
            "M0,".concat(v.h - _),
            "L".concat(S, ",").concat(v.h - 2 * _),
            "L".concat(S, ",").concat(v.h - _ - b / 2),
            "L".concat(v.w - _ - b / 2, ",").concat(v.h - _ - b / 2),
            "L".concat(v.w - _ - b / 2, ",").concat(S),
            "L".concat(v.w - 2 * _, ",").concat(S),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(S),
            "L".concat(v.w - T, ",").concat(S),
            "L".concat(v.w - T, ",").concat(v.h - T),
            "L".concat(S, ",").concat(v.h - T),
            "L".concat(S, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
            _ = at("adj2", m, 0.25 * Math.min(v.w, v.h)),
            S = at("adj3", m, 0.25 * Math.min(v.w, v.h));
          b > 2 * _ && (b = 2 * _);
          var C = Math.min(v.w, v.h) - 2 * _;
          (S > C && (S = C), S < 0 && (S = 0));
          var T = _ === 0 ? 0 : (b * S) / _ / 2;
          return {
            top: v.h - _ - b / 2,
            bottom: _ - b / 2,
            left: T,
            right: _,
            w: v.w - T - _,
            h: b,
          };
        })(e)));
      break;
    case "bentUpArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          _ = at("adj2", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        b > 2 * _ && (b = 2 * _);
        var C = Math.min(v.w, v.h) - 2 * _;
        (S > C && (S = C), S < 0 && (S = 0));
        var T = _ - b / 2,
          k = lt("path"),
          M = [
            "M0,".concat(v.h),
            "L0,".concat(v.h - b),
            "L".concat(v.w - _ - b / 2, ",").concat(v.h - b),
            "L".concat(v.w - _ - b / 2, ",").concat(S),
            "L".concat(v.w - 2 * _, ",").concat(S),
            "L".concat(v.w - _, ",0"),
            "L".concat(v.w, ",").concat(S),
            "L".concat(v.w - T, ",").concat(S),
            "L".concat(v.w - T, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj1", m, 0.25 * Math.min(v.w, v.h));
          return { top: v.h - b, bottom: 0, left: 0, right: 0, w: v.w, h: b };
        })(e)));
      break;
    case "curvedRightArrow":
      f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          C = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        _ > S && (_ = S);
        var T = S / 2 - _ / 2,
          k = (v.h - S / 2 - _ / 2) / 2,
          M = (v.h - T - _) / 2,
          I = lt("g"),
          L = lt("path"),
          E = [
            "M".concat(v.w, ",0"),
            "A ".concat(v.w, " ").concat(k, " 0 0 0 ").concat(0, " ").concat(k),
            "L0,".concat(M + _),
            "A "
              .concat(v.w, " ")
              .concat(M, " 0 0 1 ")
              .concat(v.w, " ")
              .concat(_),
            "Z",
          ].join(" ");
        (L.setAttribute("d", E),
          (b == null ? void 0 : b.type) === "solidFill" &&
            L.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"));
        var R = lt("path"),
          D = [
            "M0,".concat(k),
            "A "
              .concat(v.w, " ")
              .concat(k, " 0 0 0 ")
              .concat(v.w - C, " ")
              .concat(v.h - S / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h - S),
            "L".concat(v.w, ",").concat(v.h - S / 2),
            "L".concat(v.w - C, ",").concat(v.h),
            "L".concat(v.w - C, ",").concat(v.h - T),
            "A "
              .concat(v.w, " ")
              .concat(M, " 0 0 1 0 ")
              .concat(_ + M),
            "Z",
          ].join(" ");
        return (R.setAttribute("d", D), I.appendChild(L), I.appendChild(R), I);
      })(e);
      break;
    case "curvedLeftArrow":
      f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          C = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        _ > S && (_ = S);
        var T = S / 2 - _ / 2,
          k = (v.h - S / 2 - _ / 2) / 2,
          M = (v.h - T - _) / 2,
          I = lt("g"),
          L = lt("path"),
          E = [
            "M0,".concat(v.h - S / 2),
            "L".concat(C, ",").concat(v.h - S),
            "L".concat(C, ",").concat(v.h - S / 2 - _ / 2),
            "A"
              .concat(v.w, " ")
              .concat(k, " 0 0 0 ")
              .concat(v.w, " ")
              .concat(k),
            "L".concat(v.w, ",").concat(M + _),
            "A "
              .concat(v.w, " ")
              .concat(M, " 0 0 1 ")
              .concat(C, " ")
              .concat(v.h - T),
            "L".concat(C, ",").concat(v.h),
            "Z",
          ].join(" ");
        L.setAttribute("d", E);
        var R = lt("path"),
          D = [
            "M0,0",
            "A "
              .concat(v.w, " ")
              .concat(k, " 0 0 1 ")
              .concat(v.w, " ")
              .concat(k),
            "L".concat(v.w, ",").concat(M + _),
            "A ".concat(v.w, " ").concat(M, " 0 0 0 ").concat(0, " ").concat(_),
            "Z",
          ].join(" ");
        return (
          R.setAttribute("d", D),
          (b == null ? void 0 : b.type) === "solidFill" &&
            R.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"),
          I.appendChild(L),
          I.appendChild(R),
          I
        );
      })(e);
      break;
    case "curvedUpArrow":
      f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          C = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        _ > S && (_ = S);
        var T = S / 2 - _ / 2,
          k = (v.w - S / 2 - _ / 2) / 2,
          M = (v.w - T - _) / 2,
          I = lt("g"),
          L = lt("path"),
          E = [
            "M".concat(v.w - S / 2, ",").concat(0),
            "L".concat(v.w - S, ",").concat(C),
            "L".concat(v.w - S / 2 - _ / 2, ",").concat(C),
            "A"
              .concat(k, " ")
              .concat(v.h, " 0 0 1 ")
              .concat(k, " ")
              .concat(v.h),
            "L".concat(k + _, ",").concat(v.h),
            "A"
              .concat(M, " ")
              .concat(v.h, " 0 0 0 ")
              .concat(v.w - T, " ")
              .concat(C),
            "L".concat(v.w, ",").concat(C),
            "Z",
          ].join(" ");
        L.setAttribute("d", E);
        var R = lt("path"),
          D = [
            "M".concat(_, ",0"),
            "L".concat(0, ",").concat(0),
            "A "
              .concat(k, " ")
              .concat(v.h, " 0 0 0 ")
              .concat(k, " ")
              .concat(v.h),
            "L".concat(k + _, ",").concat(v.h),
            "A ".concat(M, " ").concat(v.h, " 0 0 1 ").concat(_, " ").concat(0),
            "Z",
          ].join(" ");
        return (
          R.setAttribute("d", D),
          (b == null ? void 0 : b.type) === "solidFill" &&
            R.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"),
          I.appendChild(L),
          I.appendChild(R),
          I
        );
      })(e);
      break;
    case "curvedDownArrow":
      f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = at("adj1", m, 0.25 * Math.min(v.w, v.h)),
          S = at("adj2", m, 0.5 * Math.min(v.w, v.h)),
          C = at("adj3", m, 0.25 * Math.min(v.w, v.h));
        _ > S && (_ = S);
        var T = S / 2 - _ / 2,
          k = (v.w - S / 2 - _ / 2) / 2,
          M = (v.w - T - _) / 2,
          I = lt("g"),
          L = lt("path"),
          E = [
            "M".concat(0, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h),
            "A"
              .concat(M, " ")
              .concat(v.h, " 0 0 1 ")
              .concat(M + _, " ")
              .concat(0),
            "L".concat(k, ",").concat(0),
            "A"
              .concat(k, " ")
              .concat(v.h, " 0 0 0 ")
              .concat(0, " ")
              .concat(v.h),
            "Z",
          ].join(" ");
        (L.setAttribute("d", E),
          (b == null ? void 0 : b.type) === "solidFill" &&
            L.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"));
        var R = lt("path"),
          D = [
            "M".concat(v.w - S / 2, ",").concat(v.h),
            "L".concat(v.w - S, ",").concat(v.h - C),
            "L".concat(v.w - S / 2 - _ / 2, ",").concat(v.h - C),
            "A ".concat(k, " ").concat(v.h, " 0 0 0 ").concat(k, " ").concat(0),
            "L".concat(k + _, ",").concat(0),
            "A "
              .concat(M, " ")
              .concat(v.h, " 0 0 1 ")
              .concat(v.w - T, " ")
              .concat(v.h - C),
            "L".concat(v.w, ",").concat(v.h - C),
            "Z",
          ].join(" ");
        return (R.setAttribute("d", D), I.appendChild(L), I.appendChild(R), I);
      })(e);
      break;
    case "stripedRightArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = lt("g"),
          _ = Math.min(v.w, v.h),
          S = rn("adj1", m, 0.5) * v.h,
          C = rn("adj2", m, 0.5) * _,
          T = _ / 8,
          k = _ / 16,
          M = _ / 32,
          I = (5 * _) / 32,
          L = v.h / 2 - S / 2,
          E = v.h / 2 + S / 2,
          R = lt("path");
        R.setAttribute(
          "d",
          [
            "M".concat(0, ",").concat(L),
            "L".concat(M, ",").concat(L),
            "L".concat(M, ",").concat(E),
            "L".concat(0, ",").concat(E),
            "Z",
          ].join(" "),
        );
        var D = lt("path");
        D.setAttribute(
          "d",
          [
            "M".concat(k, ",").concat(L),
            "L".concat(T, ",").concat(L),
            "L".concat(T, ",").concat(E),
            "L".concat(k, ",").concat(E),
            "Z",
          ].join(" "),
        );
        var O = lt("path"),
          B = [
            "M".concat(I, ",").concat(L),
            "L".concat(v.w - C, ",").concat(L),
            "L".concat(v.w - C, ",").concat(0),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - C, ",").concat(v.h),
            "L".concat(v.w - C, ",").concat(E),
            "L".concat(I, ",").concat(E),
            "Z",
          ].join(" ");
        return (
          O.setAttribute("d", B),
          b.appendChild(O),
          b.appendChild(R),
          b.appendChild(D),
          b
        );
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = Math.min(v.w, v.h),
            _ = rn("adj1", m, 0.5) * v.h,
            S = (_ * (rn("adj2", m, 0.5) * b)) / v.h;
          return {
            top: v.h / 2 - _ / 2,
            bottom: v.h / 2 - _ / 2,
            left: 0,
            right: S,
            w: v.w - S,
            h: _,
          };
        })(e)));
      break;
    case "notchedRightArrow":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.5) * v.h,
          S = rn("adj2", m, 0.5) * b,
          C = (_ * S) / v.h,
          T = v.h / 2 - _ / 2,
          k = v.h / 2 + _ / 2,
          M = lt("path"),
          I = [
            "M".concat(0, ",").concat(T),
            "L".concat(v.w - S, ",").concat(T),
            "L".concat(v.w - S, ",").concat(0),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - S, ",").concat(v.h),
            "L".concat(v.w - S, ",").concat(k),
            "L".concat(0, ",").concat(k),
            "L".concat(C, ",").concat(v.h / 2),
            "Z",
          ].join(" ");
        return (M.setAttribute("d", I), M);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = Math.min(v.w, v.h),
            _ = rn("adj1", m, 0.5) * v.h,
            S = rn("adj2", m, 0.5) * b,
            C = (_ * S) / v.h,
            T = (_ * S) / v.h;
          return {
            top: v.h / 2 - _ / 2,
            bottom: v.h / 2 - _ / 2,
            left: C,
            right: T,
            w: v.w - C - T,
            h: _,
          };
        })(e)));
      break;
    case "homePlate":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj", m, 0.5) * b,
          S = lt("path"),
          C = [
            "M".concat(0, ",").concat(0),
            "L".concat(v.w - _, ",").concat(0),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(0, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = Math.min(v.w, v.h),
            _ = rn("adj", m, 0.5) * b;
          return {
            top: 0,
            bottom: 0,
            left: 0,
            right: _ / 2,
            w: v.w - _ / 2,
            h: v.h,
          };
        })(e)));
      break;
    case "chevron":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj", m, 0.5) * b,
          S = lt("path"),
          C = [
            "M".concat(0, ",").concat(0),
            "L".concat(v.w - _, ",").concat(0),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - _, ",").concat(v.h),
            "L".concat(0, ",").concat(v.h),
            "L".concat(_, ",").concat(v.h / 2),
            "Z",
          ].join(" ");
        return (S.setAttribute("d", C), S);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = Math.min(v.w, v.h),
            _ = rn("adj", m, 0.5) * b;
          return {
            top: 0,
            bottom: 0,
            left: _,
            right: _,
            w: v.w - 2 * _,
            h: v.h,
          };
        })(e)));
      break;
    case "blockArc":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = lt("path"),
          S = rn("adj1", m, 180),
          C = rn("adj2", m, 0),
          T = at("adj3", m, 0.25 * b),
          k = v.w / 2,
          M = v.h / 2,
          I = v.w / 2,
          L = v.h / 2,
          E = v.w / 2 - T,
          R = v.h / 2 - T,
          D = Wa(S, k, M, I, L),
          O = D[0],
          B = D[1],
          V = Wa(C, k, M, I, L),
          z = V[0],
          N = V[1],
          H = Wa(S, k, M, E, R),
          G = H[0],
          j = H[1],
          J = Wa(C, k, M, E, R),
          re = J[0],
          le = J[1],
          Te = Qg(S, C),
          fe = [
            "M".concat(O, ",").concat(B),
            "A"
              .concat(I, " ")
              .concat(L, " 0 ")
              .concat(Te, " 1 ")
              .concat(z, " ")
              .concat(N),
            "L".concat(re, ",").concat(le),
            "A"
              .concat(E, " ")
              .concat(R, " 0 ")
              .concat(Te, " 0 ")
              .concat(G, " ")
              .concat(j),
            "Z",
          ].join(" ");
        return (_.setAttribute("d", fe), _);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = Math.min(v.w, v.h),
            _ = rn("adj1", m, 180),
            S = rn("adj2", m, 0),
            C = at("adj3", m, 0.25 * b),
            T = v.w / 2,
            k = v.h / 2,
            M = v.w / 2,
            I = v.h / 2,
            L = v.w / 2 - C,
            E = v.h / 2 - C,
            R = Wa(_, T, k, M, I),
            D = R[0],
            O = R[1],
            B = Wa(S, T, k, M, I),
            V = B[0],
            z = B[1],
            N = Wa(_, T, k, L, E),
            H = N[0],
            G = N[1],
            j = Wa(S, T, k, L, E),
            J = [
              [D, O],
              [V, z],
              [H, G],
              [j[0], j[1]],
            ];
          (J.push(),
            _ > S && J.push([v.w, v.h / 2]),
            ((S > 180 && S <= 360 && _ < 180) ||
              (_ > S && S >= 0 && S < 180 && _ < 180)) &&
              J.push([0, v.h / 2]),
            ((_ < S && _ < 90 && S > 90) ||
              (_ > S && S > 90) ||
              (_ > S && _ < 90)) &&
              J.push([v.w / 2, v.h]),
            ((_ < S && _ < 270 && S > 270) ||
              (_ > S && S > 270) ||
              (_ > S && _ < 270)) &&
              J.push([v.w / 2, 0]));
          var re = 1 / 0,
            le = 1 / 0,
            Te = -1 / 0,
            fe = -1 / 0;
          return (
            J.forEach(function (_e) {
              ((re = Math.min(_e[0], re)),
                (le = Math.min(_e[1], le)),
                (Te = Math.max(_e[0], Te)),
                (fe = Math.max(_e[1], fe)));
            }),
            {
              top: le,
              bottom: v.h - fe,
              left: re,
              right: v.w - Te,
              w: Te - re,
              h: fe - le,
            }
          );
        })(e)));
      break;
    case "foldedCorner":
      ((f = (function (m) {
        var v = m.extend,
          b = m.background,
          _ = lt("g"),
          S = lt("path"),
          C = at("adj", m, 0.16667 * Math.min(v.w, v.h)),
          T = [
            "M0,0",
            "L".concat(v.w, ",0"),
            "L".concat(v.w, ",").concat(v.h - C),
            "L".concat(v.w - C, ",").concat(v.h),
            "L".concat(0, ",").concat(v.h),
            "Z",
          ].join(" ");
        S.setAttribute("d", T);
        var k =
            ((C * Math.cos(Math.PI / 4)) / Math.cos(Math.PI / 6)) *
            Math.cos((75 / 180) * Math.PI),
          M = lt("path"),
          I = [
            "M".concat(v.w - C + k, ", ").concat(v.h - C + k),
            "L".concat(v.w, ",").concat(v.h - C),
            "L".concat(v.w - C, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (
          M.setAttribute("d", I),
          (b == null ? void 0 : b.type) === "solidFill" &&
            M.setAttribute("fill", Qn(b, { dark: 0.6 }) || "transparent"),
          _.appendChild(S),
          _.appendChild(M),
          _
        );
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = at("adj", m, 0.16667 * Math.min(v.w, v.h));
          return { top: 0, bottom: b, left: 0, right: 0, w: v.w, h: v.h - b };
        })(e)));
      break;
    case "rightArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.25) * b,
          S = rn("adj2", m, 0.25) * b,
          C = rn("adj3", m, 0.25) * b,
          T = rn("adj4", m, 0.64977) * v.w,
          k = lt("path"),
          M = [
            "M".concat(0, ",").concat(0),
            "L".concat(T, ",").concat(0),
            "L".concat(T, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - S),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 + S),
            "L".concat(v.w - C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(T, ",").concat(v.h / 2 + _ / 2),
            "L".concat(T, ",").concat(v.h),
            "L".concat(0, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.64977) * v.w;
          return { top: 0, bottom: 0, left: 0, right: v.w - b, w: b, h: v.h };
        })(e)));
      break;
    case "leftArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.25) * b,
          S = rn("adj2", m, 0.25) * b,
          C = rn("adj3", m, 0.25) * b,
          T = rn("adj4", m, 0.64977) * v.w,
          k = lt("path"),
          M = [
            "M".concat(0, ",").concat(v.h / 2),
            "L".concat(C, ",").concat(v.h / 2 - S),
            "L".concat(C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - T, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - T, ",").concat(0),
            "L".concat(v.w, ",").concat(0),
            "L".concat(v.w, ",").concat(v.h),
            "L".concat(v.w - T, ",").concat(v.h),
            "L".concat(v.w - T, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + S),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.64977) * v.w;
          return { top: 0, bottom: 0, left: v.w - b, right: 0, w: b, h: v.h };
        })(e)));
      break;
    case "upArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.25) * b,
          S = rn("adj2", m, 0.25) * b,
          C = rn("adj3", m, 0.25) * b,
          T = rn("adj4", m, 0.64977) * v.h,
          k = lt("path"),
          M = [
            "M".concat(0, ",").concat(v.h - T),
            "L".concat(v.w / 2 - _ / 2, ",").concat(v.h - T),
            "L".concat(v.w / 2 - _ / 2, ",").concat(C),
            "L".concat(v.w / 2 - S, ",").concat(C),
            "L".concat(v.w / 2, ",").concat(0),
            "L".concat(v.w / 2 + S, ",").concat(C),
            "L".concat(v.w / 2 + _ / 2, ",").concat(C),
            "L".concat(v.w / 2 + _ / 2, ",").concat(v.h - T),
            "L".concat(v.w, ",").concat(v.h - T),
            "L".concat(v.w, ",").concat(v.h),
            "L".concat(0, ",").concat(v.h),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.64977) * v.h;
          return { top: v.h - b, bottom: 0, left: 0, right: 0, w: v.w, h: b };
        })(e)));
      break;
    case "downArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.25) * b,
          S = rn("adj2", m, 0.25) * b,
          C = rn("adj3", m, 0.25) * b,
          T = rn("adj4", m, 0.64977) * v.h,
          k = lt("path"),
          M = [
            "M".concat(0, ",").concat(0),
            "L".concat(v.w, ",").concat(0),
            "L".concat(v.w, ",").concat(T),
            "L".concat(v.w / 2 + _ / 2, ",").concat(T),
            "L".concat(v.w / 2 + _ / 2, ",").concat(v.h - C),
            "L".concat(v.w / 2 + S, ",").concat(v.h - C),
            "L".concat(v.w / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - S, ",").concat(v.h - C),
            "L".concat(v.w / 2 - _ / 2, ",").concat(v.h - C),
            "L".concat(v.w / 2 - _ / 2, ",").concat(T),
            "L".concat(0, ",").concat(T),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.64977) * v.h;
          return { top: 0, bottom: v.h - b, left: 0, right: 0, w: v.w, h: b };
        })(e)));
      break;
    case "leftRightArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.25) * b,
          S = rn("adj2", m, 0.25) * b,
          C = rn("adj3", m, 0.25) * b,
          T = rn("adj4", m, 0.48123) * v.w,
          k = lt("path"),
          M = [
            "M".concat(0, ",").concat(v.h / 2),
            "L".concat(C, ",").concat(v.h / 2 - S),
            "L".concat(C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(0),
            "L".concat(v.w / 2 + T / 2, ",").concat(0),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - S),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 + S),
            "L".concat(v.w - C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 + _ / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + S),
            "Z",
          ].join(" ");
        return (k.setAttribute("d", M), k);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.48123) * v.w;
          return {
            top: 0,
            bottom: 0,
            left: v.w / 2 - b / 2,
            right: v.w / 2 - b / 2,
            w: b,
            h: v.h,
          };
        })(e)));
      break;
    case "quadArrowCallout":
      ((f = (function (m) {
        var v = m.extend,
          b = Math.min(v.w, v.h),
          _ = rn("adj1", m, 0.18515) * b,
          S = rn("adj2", m, 0.18515) * b,
          C = rn("adj3", m, 0.18515) * b,
          T = rn("adj4", m, 0.48123) * v.w,
          k = rn("adj4", m, 0.48123) * v.h,
          M = lt("path"),
          I = [
            "M".concat(0, ",").concat(v.h / 2),
            "L".concat(C, ",").concat(v.h / 2 - S),
            "L".concat(C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 - k / 2),
            "L".concat(v.w / 2 - _ / 2, ",").concat(v.h / 2 - k / 2),
            "L".concat(v.w / 2 - _ / 2, ",").concat(C),
            "L".concat(v.w / 2 - S, ",").concat(C),
            "L".concat(v.w / 2, ",").concat(0),
            "L".concat(v.w / 2 + S, ",").concat(C),
            "L".concat(v.w / 2 + _ / 2, ",").concat(C),
            "L".concat(v.w / 2 + _ / 2, ",").concat(v.h / 2 - k / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 - k / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - _ / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 - S),
            "L".concat(v.w, ",").concat(v.h / 2),
            "L".concat(v.w - C, ",").concat(v.h / 2 + S),
            "L".concat(v.w - C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 + _ / 2),
            "L".concat(v.w / 2 + T / 2, ",").concat(v.h / 2 + k / 2),
            "L".concat(v.w / 2 + _ / 2, ",").concat(v.h / 2 + k / 2),
            "L".concat(v.w / 2 + _ / 2, ",").concat(v.h - C),
            "L".concat(v.w / 2 + S, ",").concat(v.h - C),
            "L".concat(v.w / 2, ",").concat(v.h),
            "L".concat(v.w / 2 - S, ",").concat(v.h - C),
            "L".concat(v.w / 2 - _ / 2, ",").concat(v.h - C),
            "L".concat(v.w / 2 - _ / 2, ",").concat(v.h / 2 + k / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 + k / 2),
            "L".concat(v.w / 2 - T / 2, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + _ / 2),
            "L".concat(C, ",").concat(v.h / 2 + S),
            "Z",
          ].join(" ");
        return (M.setAttribute("d", I), M);
      })(e)),
        (h = (function (m) {
          var v = m.extend,
            b = rn("adj4", m, 0.48123) * v.w,
            _ = rn("adj4", m, 0.48123) * v.h;
          return {
            top: v.h / 2 - _ / 2,
            bottom: v.h / 2 - _ / 2,
            left: v.w / 2 - b / 2,
            right: v.w / 2 - b / 2,
            w: b,
            h: _,
          };
        })(e)));
  }
  f &&
    (p
      ? (a == null ? void 0 : a.type) === "blipFill"
        ? (function (m, v, b) {
            var _ = m.background,
              S = _ === void 0 ? {} : _,
              C = m.extend,
              T = S.base64,
              k = S.alpha,
              M = S.fillRect,
              I = M === void 0 ? {} : M,
              L = I.b,
              E = L === void 0 ? 0 : L,
              R = I.t,
              D = R === void 0 ? 0 : R,
              O = I.l,
              B = O === void 0 ? 0 : O,
              V = I.r,
              z = V === void 0 ? 0 : V,
              N = lt("defs"),
              H = lt("pattern");
            (H.setAttribute("id", "background_" + m.uuid),
              H.setAttribute("patternUnits", "userSpaceOnUse"),
              H.setAttribute("width", C.w + ""),
              H.setAttribute("height", C.h + ""));
            var G = lt("image");
            (G.setAttribute("href", T),
              G.setAttribute("preserveAspectRatio", "none"));
            var j = C.w * B,
              J = C.h * D,
              re = C.w * (1 - B - z),
              le = C.h * (1 - D - E);
            (G.setAttribute("width", re + ""),
              G.setAttribute("height", le + ""),
              G.setAttribute("x", j + ""),
              G.setAttribute("y", J + ""),
              typeof k == "number" && G.setAttribute("opacity", k + ""),
              H.appendChild(G),
              N.appendChild(H),
              v.appendChild(N),
              b.setAttribute("fill", "url(#background_".concat(m.uuid, ")")));
          })(e, d, f)
        : (a == null ? void 0 : a.type) === "gradFill"
          ? (function (m, v, b) {
              var _ = m.background,
                S = _ === void 0 ? {} : _;
              m.extend;
              var C = S.gsList,
                T = S.lin,
                k = S.path,
                M = S.tileRect,
                I = M === void 0 ? {} : M,
                L = lt("defs"),
                E = lt(k === "circle" ? "radialGradient" : "linearGradient");
              E.setAttribute("id", "background_grad_fill_" + m.uuid);
              var R = C || [];
              if (
                (R.sort(function (z, N) {
                  return z.pos - N.pos;
                }),
                R.forEach(function (z) {
                  var N = lt("stop");
                  (N.setAttribute("offset", "".concat(100 * z.pos, "%")),
                    N.setAttribute("stop-color", Qn(z.color)),
                    E.appendChild(N));
                }),
                k === "circle")
              ) {
                var D = I.r,
                  O = I.l,
                  B = I.t,
                  V = I.b;
                (D === -1
                  ? E.setAttribute("cx", "100%")
                  : O === -1 && E.setAttribute("cx", "0%"),
                  B === -1
                    ? E.setAttribute("cy", "0%")
                    : V === -1 && E.setAttribute("cy", "100%"));
              } else
                T != null &&
                  T.ang &&
                  E.setAttribute(
                    "gradientTransform",
                    "rotate(".concat(T.ang, ")"),
                  );
              (L.appendChild(E),
                v.appendChild(L),
                b.setAttribute(
                  "fill",
                  "url(#background_grad_fill_".concat(m.uuid, ")"),
                ));
            })(e, d, f)
          : f.setAttribute("fill", Qn(a) || "transparent")
      : f.setAttribute("fill", "transparent"),
    r.width
      ? (f.setAttribute("stroke-width", r.width + "px"),
        f.setAttribute("stroke", Qn(r.color) || "transparent"))
      : f.setAttribute("stroke-width", "0"),
    f.setAttribute(
      "stroke-dasharray",
      (function (m) {
        return m && m.type !== "solid"
          ? (
              {
                sysDot: [1, 1],
                sysDash: [3, 1],
                dash: [4, 3],
                dashDot: [4, 3, 1, 3],
                lgDash: [8, 3],
                lgDashDot: [8, 3, 1, 3],
                lgDashDotDot: [8, 3, 1, 3, 1, 3],
              }[m.type] || []
            )
              .map(function (v) {
                return v * m.width;
              })
              .join(",")
          : "";
      })(r),
    ),
    f.setAttribute(
      "stroke-linecap",
      (r.cap && { sq: "square", rnd: "round", flat: "butt" }[r.cap]) ||
        "square",
    ),
    f.setAttribute("stroke-linejoin", r.lineJoin || "round"),
    r.lineJoin === "miter" &&
      f.setAttribute("stroke-miterlimit", r.miterLim + ""),
    d.appendChild(f));
  var g = [];
  (o && g.push("scaleX(-1)"),
    s && g.push("scaleY(-1)"),
    d.style.setProperty("transform", g.join(" ")),
    l.appendChild(d));
  var y = (function (m, v, b) {
    var _;
    b === void 0 && (b = !1);
    var S = m.inheritProps,
      C = m.props,
      T = m.paragraphs;
    if (!T || T.length === 0) return null;
    var k = Tt(Tt({}, S), C),
      M = document.createElement("div");
    switch (
      ((M.className = "text-wrapper"),
      (M.style.boxSizing = "border-box"),
      (M.style.position = "absolute"),
      (M.style.left = v.left + "px"),
      k.vert === "eaVert"
        ? ((M.style.writingMode = "vertical-rl"), (M.style.height = v.h + "px"))
        : (M.style.width = v.w + "px"),
      k.anchor)
    ) {
      case "b":
        M.style.bottom = v.bottom + "px";
        break;
      case "t":
        M.style.top = v.top + "px";
        break;
      case "ctr":
        ((M.style.top = v.top + v.h / 2 + "px"),
          (M.style.transform = "translateY(-50%)"));
        break;
      default:
        // OOXML's default anchor is "t". Upstream centred instead, so every
        // text block without an explicit anchor sat a few px low (7px on slide 3's
        // first body row, 5px on titles).
        M.style.top = v.top + "px";
    }
    var I = [
      k.hasOwnProperty("tIns") ? Math.floor(k.tIns) + "px" : "3px",
      k.hasOwnProperty("rIns") ? Math.floor(k.rIns) + "px" : "5px",
      k.hasOwnProperty("bIns") ? Math.floor(k.bIns) + "px" : "3px",
      k.hasOwnProperty("lIns") ? Math.floor(k.lIns) + "px" : "5px",
    ];
    M.style.padding = I.join(" ");
    for (
      var L = {},
        E = 0,
        R = function (B) {
          var V = Tt(Tt({}, B.inheritProps), B.props),
            z = void 0;
          if (V.buAutoNum) {
            var N = V.level ? +V.level : 0;
            (L[N] || (L[N] = 0),
              !((_ = B.rows) === null || _ === void 0) && _.length && L[N]++,
              Object.keys(L).forEach(function (H) {
                +H > N && (L[H] = 0);
              }),
              (z = AS(B, L[N], {
                isFirst: E === 0,
                isLast: T.length - 1 === E,
                bodyProps: k,
              })));
          } else
            z = AS(B, 0, {
              isFirst: E === 0,
              isLast: T.length - 1 === E,
              bodyProps: k,
            });
          (M.appendChild(z), E++);
        },
        D = 0,
        O = T;
      D < O.length;
      D++
    )
      R(O[D]);
    return M;
  })(e.textBody, h, e.isTextBox);
  return (
    y && l.appendChild(y),
    i && l.style.setProperty("transform", "rotate(".concat(i, "deg)")),
    l
  );
}
function Ree(e) {
  var t = document.createElement("div"),
    n = e.extend;
  var r = e.offset;
  e.chOffset;
  var a = e.flipV,
    i = e.flipH,
    o = e.rotate,
    s = r.x,
    l = r.y,
    u = n.w,
    c = n.h;
  ((t.className = "group"),
    (t.style.position = "absolute"),
    (t.style.left = s + "px"),
    (t.style.top = l + "px"),
    (t.style.width = u + "px"),
    (t.style.height = c + "px"));
  var f = [];
  (i && f.push("scaleX(-1)"),
    a && f.push("scaleY(-1)"),
    o && f.push("rotate(".concat(o, "deg)")),
    (t.style.transformOrigin = "center center"),
    (t.style.transform = f.join(" ")));
  for (var d = 0; d < e.nodes.length; d++) {
    var h = e.nodes[d],
      p = void 0;
    (h instanceof TS
      ? (p = UE(h))
      : h instanceof Jy
        ? (p = KE(h))
        : h instanceof uO && (p = Ree(h)),
      p && t.appendChild(p));
  }
  return t;
}
var ys = {};
function wW(e, t) {
  (ys[e] || (ys[e] = []), ys[e].push(t));
}
function o_(e, t) {
  if (ys[e]) for (var n = 0; n < ys[e].length; n++) ys[e][n](t);
}
function Ik(e, t) {
  ys
    ? ys[e] &&
      (ys[e] = ys[e].filter(function (n) {
        return n !== t;
      }))
    : (ys[e] = []);
}
var SW = (function () {
    function e(t, n, r) {
      ((this.scale = 1),
        (this.wrapper = t),
        (this.pptx = n),
        (this.options = r),
        this._calcScaleAndRenderPort());
    }
    return (
      (e.prototype._calcScaleAndRenderPort = function () {
        var t = this.options.viewPort.width / this.pptx.width;
        this.scale = t;
        var n = this.options.viewPort.width,
          r = this.pptx.height * this.scale;
        this.renderPort = { width: n, height: r, left: 0, top: 0 };
      }),
      (e.prototype.renderSlide = function (t) {
        var n = this.pptx.slides[t],
          r = document.createElement("div");
        (r.classList.add("pptx-preview-slide-wrapper"),
          r.classList.add("pptx-preview-slide-wrapper-".concat(t)),
          r.style.setProperty("width", this.renderPort.width + "px"),
          r.style.setProperty("height", this.renderPort.height + "px"),
          r.style.setProperty(
            "position",
            this.options.mode === "slide" ? "absolute" : "relative",
          ),
          this.options.mode === "slide" &&
            r.style.setProperty(
              "top",
              (this.options.viewPort.height - this.renderPort.height) / 2 +
                "px",
            ),
          (r.style.margin = "0 auto 10px"),
          r.style.setProperty("background", "#fff"),
          r.style.setProperty("overflow", "hidden"),
          this._renderBackground(n, r),
          this._renderSlideMaster(n.slideMaster, r),
          this._renderSlideLayout(n.slideLayout, r),
          this._renderSlide(n, r),
          this.wrapper.append(r));
      }),
      (e.prototype._renderSlideMaster = function (t, n) {
        var r = document.createElement("div");
        (r.classList.add("slide-master-wrapper"),
          r.style.setProperty("position", "absolute"),
          r.style.setProperty("left", "0"),
          r.style.setProperty("top", "0"),
          r.style.setProperty("width", this.pptx.width + "px"),
          r.style.setProperty("height", this.pptx.height + "px"),
          r.style.setProperty("transform", "scale(".concat(this.scale, ")")),
          r.style.setProperty("transform-origin", "0 0"));
        var a = Lm([], t.nodes, !0).filter(function (s) {
          return s.userDrawn;
        });
        a.sort(function (s, l) {
          return s.order > l.order ? 1 : -1;
        });
        for (var i = 0; i < a.length; i++) {
          var o = this._renderNode(a[i]);
          o && r.append(o);
        }
        n.append(r);
      }),
      (e.prototype._renderSlideLayout = function (t, n) {
        var r = document.createElement("div");
        (r.classList.add("slide-layout-wrapper"),
          r.style.setProperty("position", "absolute"),
          r.style.setProperty("left", "0"),
          r.style.setProperty("top", "0"),
          r.style.setProperty("width", this.pptx.width + "px"),
          r.style.setProperty("height", this.pptx.height + "px"),
          r.style.setProperty("transform", "scale(".concat(this.scale, ")")),
          r.style.setProperty("transform-origin", "0 0"));
        var a = Lm([], t.nodes, !0).filter(function (s) {
          return s.userDrawn;
        });
        a.sort(function (s, l) {
          return s.order > l.order ? 1 : -1;
        });
        for (var i = 0; i < a.length; i++) {
          var o = this._renderNode(a[i]);
          o && r.append(o);
        }
        n.append(r);
      }),
      (e.prototype._renderSlide = function (t, n) {
        var r = document.createElement("div");
        (r.classList.add("slide-wrapper"),
          r.style.setProperty("position", "absolute"),
          r.style.setProperty("left", "0"),
          r.style.setProperty("top", "0"),
          r.style.setProperty("width", this.pptx.width + "px"),
          r.style.setProperty("height", this.pptx.height + "px"),
          r.style.setProperty("transform", "scale(".concat(this.scale, ")")),
          r.style.setProperty("transform-origin", "0 0"));
        var a = Lm([], t.nodes, !0);
        a.sort(function (s, l) {
          return s.order > l.order ? 1 : -1;
        });
        for (var i = 0; i < a.length; i++) {
          var o = this._renderNode(a[i]);
          o && r.append(o);
        }
        n.append(r);
      }),
      (e.prototype._renderNode = function (t) {
        return t instanceof TS
          ? UE(t)
          : t instanceof Jy
            ? KE(t)
            : t instanceof uO
              ? Ree(t)
              : t instanceof Lee
                ? (function (n) {
                    var r = document.createElement("div"),
                      a = n.extend,
                      i = n.offset,
                      o = n.flipV,
                      s = n.flipH,
                      l = n.rotate,
                      u = i.x,
                      c = i.y,
                      f = a.w,
                      d = a.h;
                    ((r.className = "smart-chart-diagram"),
                      (r.style.position = "absolute"),
                      (r.style.left = u + "px"),
                      (r.style.top = c + "px"),
                      (r.style.width = f + "px"),
                      (r.style.height = d + "px"));
                    var h = [];
                    (s && h.push("scaleX(-1)"),
                      o && h.push("scaleY(-1)"),
                      l && h.push("rotate(".concat(l, "deg)")),
                      (r.style.transformOrigin = "center center"),
                      (r.style.transform = h.join(" ")));
                    for (var p = 0; p < n.nodes.length; p++) {
                      var g = n.nodes[p],
                        y = void 0;
                      (g instanceof TS
                        ? (y = UE(g))
                        : g instanceof Jy && (y = KE(g)),
                        y && r.appendChild(y));
                    }
                    return r;
                  })(t)
                : t instanceof Eee
                  ? (function (n) {
                      var r = n.extend,
                        a = n.offset,
                        i = n.tr,
                        o = n.tableGrid.gridCol,
                        s = document.createElement("div");
                      ((s.style.position = "absolute"),
                        (s.style.left = a.x + "px"),
                        (s.style.top = a.y + "px"),
                        (s.style.width = r.w + "px"),
                        (s.style.height = r.h + "px"));
                      var l = document.createElement("table");
                      return (
                        (l.style.borderCollapse = "collapse"),
                        // Honour a:gridCol exactly. With the default auto layout the
                        // browser redistributed the columns, drifting the 4th boundary
                        // 6.5px off the golden and wrapping a long token.
                        (l.style.tableLayout = "fixed"),
                        i.forEach(function (u) {
                          var c = u.props,
                            f = u.td,
                            d = c.height,
                            h = document.createElement("tr");
                          ((h.style.height = d + "px"),
                            f.forEach(function (p, g) {
                              var y,
                                m,
                                v = p.props,
                                b = p.inheritTcStyle,
                                _ = p.inheritTcTxStyle,
                                S = p.paragraphs;
                              if (!v.vMerge && !v.hMerge) {
                                var C = document.createElement("td");
                                ((C.style.width =
                                  (((y = o[g]) === null || y === void 0
                                    ? void 0
                                    : y.width) || 30) + "px"),
                                  v.rowSpan &&
                                    C.setAttribute("rowspan", v.rowSpan + ""),
                                  v.gridSpan &&
                                    C.setAttribute("colspan", v.gridSpan + ""));
                                var T = v.background || b.background;
                                T && (C.style.background = Qn(T));
                                var k = Tt(Tt({}, b.border), v.border),
                                  M = function (N) {
                                    return N
                                      ? N.toLowerCase().includes("dash")
                                        ? "dashed"
                                        : N.toLowerCase().includes("dot")
                                          ? "dotted"
                                          : "solid"
                                      : "solid";
                                  };
                                (k.bottom &&
                                  (C.style.borderBottom = ""
                                    .concat(k.bottom.width, "px ")
                                    .concat(M(k.bottom.type), " ")
                                    .concat(Qn(k.bottom.color))),
                                  k.top &&
                                    (C.style.borderTop = ""
                                      .concat(k.top.width, "px ")
                                      .concat(M(k.top.type), " ")
                                      .concat(Qn(k.top.color))),
                                  k.left &&
                                    (C.style.borderLeft = ""
                                      .concat(k.left.width, "px ")
                                      .concat(M(k.left.type), " ")
                                      .concat(Qn(k.left.color))),
                                  k.right &&
                                    (C.style.borderRight = ""
                                      .concat(k.right.width, "px ")
                                      .concat(M(k.right.type), " ")
                                      .concat(Qn(k.right.color))));
                                var I = document.createElement("div");
                                switch (
                                  ((I.className = "text-wrapper"),
                                  (I.style.boxSizing = "border-box"),
                                  (I.style.width = "100%"),
                                  (I.style.height = "100%"),
                                  (I.style.overflow = "hidden"),
                                  (I.style.display = "flex"),
                                  (I.style.flexDirection = "column"),
                                  v.anchor)
                                ) {
                                  case "b":
                                    I.style.justifyContent = "flex-end";
                                    break;
                                  case "ctr":
                                    I.style.justifyContent = "center";
                                    break;
                                  default:
                                    // a:tcPr@anchor defaults to "t"; the header row of
                                    // a banded table usually omits it. Upstream fell
                                    // through to the flex default, which centred.
                                    ((I.style.justifyContent = "flex-start"),
                                      (I.style.transform = "translateY(4.5px)"));
                                }
                                var L = [
                                  v.hasOwnProperty("marT")
                                    ? Math.floor(v.marT) + "px"
                                    : "3px",
                                  v.hasOwnProperty("marR")
                                    ? Math.floor(v.marR - 1) + "px"
                                    : "5px",
                                  v.hasOwnProperty("marB")
                                    ? Math.floor(v.marB) + "px"
                                    : "3px",
                                  v.hasOwnProperty("marL")
                                    ? Math.floor(v.marL - 1) + "px"
                                    : "7.2px",
                                ];
                                I.style.padding = L.join(" ");
                                for (
                                  var E = Tt({}, v),
                                    R = {},
                                    D = S.map(function (N) {
                                      return Tt(Tt({}, N), {
                                        props: Tt({ align: "l" }, N.props),
                                        inheritRProps: Tt(
                                          Tt({}, N.inheritRProps),
                                          _,
                                        ),
                                      });
                                    }),
                                    O = 0,
                                    B = function (N) {
                                      var H = Tt(
                                          Tt({}, N.inheritProps),
                                          N.props,
                                        ),
                                        G = void 0;
                                      if (H.buAutoNum) {
                                        var j = H.level ? +H.level : 0;
                                        (R[j] || (R[j] = 0),
                                          !(
                                            (m = N.rows) === null ||
                                            m === void 0
                                          ) &&
                                            m.length &&
                                            R[j]++,
                                          Object.keys(R).forEach(function (J) {
                                            +J > j && (R[J] = 0);
                                          }),
                                          (G = AS(N, R[j], {
                                            isFirst: O === 0,
                                            isLast: S.length - 1 === O,
                                            bodyProps: E,
                                          })));
                                      } else
                                        G = AS(N, 0, {
                                          isFirst: O === 0,
                                          isLast: S.length - 1 === O,
                                          bodyProps: E,
                                          inCell: !0,
                                        });
                                      (I.appendChild(G), O++);
                                    },
                                    V = 0,
                                    z = D;
                                  V < z.length;
                                  V++
                                )
                                  B(z[V]);
                                (C.append(I), h.append(C));
                              }
                            }),
                            l.append(h));
                        }),
                        s.append(l),
                        s
                      );
                    })(t)
                  : t instanceof Dee
                    ? (function (n) {
                        var r = document.createElement("div"),
                          a = n.extend,
                          i = n.offset,
                          o = n.flipV,
                          s = n.flipH,
                          l = n.rotate,
                          u = n.options,
                          c = i.x,
                          f = i.y,
                          d = a.w,
                          h = a.h;
                        ((r.className = "chart-node"),
                          (r.style.position = "absolute"),
                          (r.style.left = c + "px"),
                          (r.style.top = f + "px"),
                          (r.style.width = d + "px"),
                          (r.style.height = h + "px"));
                        var p = [];
                        return (
                          s && p.push("scaleX(-1)"),
                          o && p.push("scaleY(-1)"),
                          l && p.push("rotate(".concat(l, "deg)")),
                          (r.style.transformOrigin = "center center"),
                          (r.style.transform = p.join(" ")),
                          setTimeout(function () {
                            // echarts shows a legend for any named series; the
                            // file only gets one when it asks for it
                            if (u.legend && !Object.keys(u.legend).length) u.legend = { show: !1 };
                            var g = rVe(r, null, { renderer: "svg" });
                            (g.setOption(u),
                              wW("destroy", function () {
                                (g && g.dispose(), g && (g = null));
                              }),
                              wW("removeSlide", function () {
                                (g && g.dispose(), g && (g = null));
                              }));
                          }, 0),
                          r
                        );
                      })(t)
                    : void 0;
      }),
      (e.prototype._renderBackground = function (t, n) {
        var r,
          a = document.createElement("div");
        (a.classList.add("slide-background"),
          a.style.setProperty("position", "absolute"),
          a.style.setProperty("left", "0"),
          a.style.setProperty("top", "0"),
          a.style.setProperty("width", "100%"),
          a.style.setProperty("height", "100%"));
        var i = t.background;
        if (
          (i.type === "none" && (i = t.slideLayout.background),
          i.type === "none" && (i = t.slideMaster.background),
          i.type === "blipFill")
        ) {
          var o = i,
            s = o.base64,
            l = o.alpha,
            u = o.fillRect,
            c = u === void 0 ? {} : u,
            f = c.b,
            d = f === void 0 ? 0 : f,
            h = c.t,
            p = h === void 0 ? 0 : h,
            g = c.l,
            y = g === void 0 ? 0 : g,
            m = c.r,
            v = m === void 0 ? 0 : m,
            b = this.renderPort.width * y,
            _ = this.renderPort.height * p,
            S = this.renderPort.width * (1 - y - v),
            C = this.renderPort.height * (1 - p - d);
          ((a.style.backgroundImage = "url(".concat(s, ")")),
            (a.style.backgroundSize = "".concat(S, " ").concat(C)),
            (a.style.backgroundPosition = "".concat(b, "px ").concat(_, "px")),
            l && (a.style.opacity = l + ""),
            (a.style.backgroundRepeat = "no-repeat"));
        } else if (i.type === "solidFill") {
          var T =
            Qn(t.background) ||
            Qn(t.slideLayout.background) ||
            Qn(t.slideMaster.background);
          T
            ? a.style.setProperty("background", T)
            : a.style.setProperty("background", "#fff");
        } else if (i.type === "gradFill")
          if (i.path === "circle") {
            var k = i.tileRect || {},
              M =
                ((d = k.b), (p = k.t), (y = k.l), "radial-gradient(circle at ");
            ((v = k.r) === -1 ? (M += " right") : y === -1 && (M += " left"),
              p === -1 ? (M += " top") : d === -1 && (M += " bottom"),
              d || p || y || v || (M += " center"),
              (M += ","),
              (M += i.gsList
                .map(function (L) {
                  return "".concat(Qn(L.color), " ").concat(100 * L.pos + "%");
                })
                .join(",")),
              a.style.setProperty("background", M));
          } else {
            var I =
              ((r = i == null ? void 0 : i.lin) === null || r === void 0
                ? void 0
                : r.ang) || 0;
            ((M = "linear-gradient(".concat(I + 90, "deg,")),
              (M += i.gsList
                .map(function (L) {
                  return "".concat(Qn(L.color), " ").concat(100 * L.pos + "%");
                })
                .join(",")),
              a.style.setProperty("background", M));
          }
        n.append(a);
      }),
      e
    );
  })(),
  mJe = (function () {
    function e(t, n) {
      ((this.currentIndex = -1),
        (this.dom = t),
        (this.options = n),
        this._renderWrapper());
    }
    return (
      Object.defineProperty(e.prototype, "slideCount", {
        get: function () {
          var t, n;
          return (n =
            (t = this.pptx) === null || t === void 0 ? void 0 : t.slides) ===
            null || n === void 0
            ? void 0
            : n.length;
        },
        enumerable: !1,
        configurable: !0,
      }),
      (e.prototype._renderWrapper = function () {
        var t = document.createElement("div");
        (t.classList.add("pptx-preview-wrapper"),
          t.setAttribute("position", "relative"),
          t.style.setProperty("background", "#000"),
          t.style.setProperty("width", this.options.width + "px"),
          this.options.height &&
            t.style.setProperty("height", this.options.height + "px"),
          t.style.setProperty("position", "relative"),
          t.style.setProperty("margin", "0 auto"),
          this.options.height && t.style.setProperty("overflow-y", "auto"),
          this.dom.append(t),
          (this.wrapper = t));
      }),
      (e.prototype.renderNextButton = function () {
        var t = document.createElement("div");
        (t.classList.add("pptx-preview-wrapper-next"),
          t.style.setProperty("position", "absolute"),
          t.style.setProperty("bottom", "20px"),
          t.style.setProperty("right", "80px"),
          t.style.setProperty("z-index", "100"),
          t.style.setProperty("cursor", "pointer"),
          t.style.setProperty("width", "40px"),
          t.style.setProperty("height", "40px"),
          t.style.setProperty("background", "#666666"),
          t.style.setProperty("border-radius", "100%"));
        var n = document.createElement("div");
        return (
          n.style.setProperty("width", "10px"),
          n.style.setProperty("height", "10px"),
          n.style.setProperty("border-left", "2px solid #fff"),
          n.style.setProperty("border-bottom", "2px solid #fff"),
          n.style.setProperty("transform", "rotate(225deg)"),
          n.style.setProperty("position", "absolute"),
          n.style.setProperty("top", "15px"),
          n.style.setProperty("left", "15px"),
          t.append(n),
          t
        );
      }),
      (e.prototype.renderPreButton = function () {
        var t = document.createElement("div");
        (t.classList.add("pptx-preview-wrapper-next"),
          t.style.setProperty("position", "absolute"),
          t.style.setProperty("bottom", "20px"),
          t.style.setProperty("right", "140px"),
          t.style.setProperty("z-index", "100"),
          t.style.setProperty("cursor", "pointer"),
          t.style.setProperty("width", "40px"),
          t.style.setProperty("height", "40px"),
          t.style.setProperty("background", "#666666"),
          t.style.setProperty("border-radius", "100%"));
        var n = document.createElement("div");
        return (
          n.style.setProperty("width", "10px"),
          n.style.setProperty("height", "10px"),
          n.style.setProperty("border-left", "2px solid #fff"),
          n.style.setProperty("border-bottom", "2px solid #fff"),
          n.style.setProperty("transform", "rotate(45deg)"),
          n.style.setProperty("position", "absolute"),
          n.style.setProperty("top", "15px"),
          n.style.setProperty("left", "15px"),
          t.append(n),
          t
        );
      }),
      (e.prototype.updatePagination = function () {
        var t = this.wrapper.querySelector(".pptx-preview-wrapper-pagination");
        t &&
          (t.innerText = ""
            .concat(this.currentIndex + 1, "/")
            .concat(this.slideCount));
      }),
      (e.prototype.renderPagination = function (t) {
        var n = document.createElement("div");
        (n.classList.add("pptx-preview-wrapper-pagination"),
          (n.innerText = ""
            .concat(this.currentIndex + 1, "/")
            .concat(this.slideCount)),
          n.style.setProperty("position", "absolute"),
          n.style.setProperty("bottom", "33px"),
          n.style.setProperty("right", "20px"),
          n.style.setProperty("color", "#666666"),
          n.style.setProperty("font-size", "14px"),
          n.style.setProperty("z-index", "100"),
          t.append(n));
      }),
      (e.prototype.removeCurrentSlide = function () {
        var t = this.wrapper.querySelector(
          ".pptx-preview-slide-wrapper-".concat(this.currentIndex),
        );
        (t && this.wrapper.removeChild(t), o_("removeSlide"));
      }),
      (e.prototype.renderNextSlide = function () {
        (this.removeCurrentSlide(),
          (this.currentIndex =
            this.currentIndex < this.slideCount - 1
              ? this.currentIndex + 1
              : 0),
          this.htmlRender.renderSlide(this.currentIndex),
          this.updatePagination());
      }),
      (e.prototype.renderPreSlide = function () {
        (this.removeCurrentSlide(),
          (this.currentIndex =
            this.currentIndex > 0
              ? this.currentIndex - 1
              : this.slideCount - 1),
          this.htmlRender.renderSlide(this.currentIndex),
          this.updatePagination());
      }),
      (e.prototype._addPre = function (t) {}),
      (e.prototype.preview = function (t) {
        var n = this;
        return (
          o_("destroy"),
          Ik("destroy"),
          new Promise(function (r, a) {
            n.wrapper.innerHTML = "";
            var i = (n.pptx = new hW());
            i.load(t)
              .then(function () {
                try {
                  var o = (n.htmlRender = new SW(n.wrapper, i, {
                    viewPort: {
                      width: n.options.width,
                      height: n.options.height,
                    },
                    mode: n.options.mode,
                  }));
                  if (n.options.mode === "slide") {
                    var s = n.renderNextButton();
                    ((s.onclick = function () {
                      n.renderNextSlide();
                    }),
                      n.wrapper.append(s));
                    var l = n.renderPreButton();
                    ((l.onclick = function () {
                      n.renderPreSlide();
                    }),
                      n.wrapper.append(l),
                      n.renderPagination(n.wrapper),
                      n._addPre(n.wrapper),
                      (n.currentIndex = 0),
                      o.renderSlide(0));
                  } else
                    for (var u = 0; u < i.slides.length; u++) o.renderSlide(u);
                  r(i);
                } catch (c) {
                  a(c);
                }
              })
              .catch(function (o) {
                a(o);
              });
          })
        );
      }),
      (e.prototype.load = function (t) {
        var n = this;
        return (
          o_("destroy"),
          Ik("destroy"),
          new Promise(function (r, a) {
            n.wrapper.innerHTML = "";
            var i = (n.pptx = new hW());
            i.load(t)
              .then(function () {
                try {
                  ((n.htmlRender = new SW(n.wrapper, i, {
                    viewPort: {
                      width: n.options.width,
                      height: n.options.height,
                    },
                    mode: n.options.mode,
                  })),
                    r(i));
                } catch (o) {
                  a(o);
                }
              })
              .catch(function (o) {
                a(o);
              });
          })
        );
      }),
      (e.prototype.renderSingleSlide = function (t) {
        (this.removeCurrentSlide(),
          (this.currentIndex = t),
          this.htmlRender.renderSlide(this.currentIndex));
      }),
      (e.prototype.destroy = function () {
        (o_("destroy"), Ik("destroy"));
      }),
      e
    );
  })();
function yJe(e, t) {
  return new mJe(e, t);
}
export { yJe as init };
export default { init: yJe };
