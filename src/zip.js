/* A zip reader for the OOXML package formats: it walks the central directory
   and inflates the entries with the platform's DecompressionStream, so the
   extension carries no third-party zip library.

   Only the subset a .docx / .xlsx / .pptx needs is implemented: stored entries
   and deflate, with the same `file(name).async("string" | "base64" |
   "uint8array")` shape the readers use. */
(function (root) {
  "use strict";

  var EOCD = 0x06054b50;
  var CENTRAL = 0x02014b50;
  var LOCAL = 0x04034b50;

  function u16(b, i) {
    return b[i] | (b[i + 1] << 8);
  }

  function u32(b, i) {
    return (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16)) + b[i + 3] * 0x1000000;
  }

  /* the end of central directory record sits at the tail, after the comment */
  function findEocd(bytes) {
    var min = Math.max(0, bytes.length - 65557);
    for (var i = bytes.length - 22; i >= min; i--) {
      if (u32(bytes, i) === EOCD) return i;
    }
    return -1;
  }

  async function inflate(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("this browser cannot inflate a zip entry");
    }
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function base64(bytes) {
    var out = "";
    for (var i = 0; i < bytes.length; i += 0x8000) {
      out += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(out);
  }

  function decode(bytes, type) {
    if (type === "base64") return base64(bytes);
    if (type === "uint8array" || type === "arraybuffer") return bytes;
    return new TextDecoder("utf-8").decode(bytes);
  }

  async function load(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var eocd = findEocd(bytes);
    if (eocd < 0) throw new Error("not a zip file");

    var count = u16(bytes, eocd + 10);
    var offset = u32(bytes, eocd + 16);
    var entries = {};
    var order = [];

    for (var n = 0; n < count; n++) {
      if (u32(bytes, offset) !== CENTRAL) break;
      var method = u16(bytes, offset + 10);
      var compressed = u32(bytes, offset + 20);
      var size = u32(bytes, offset + 24);
      var nameLen = u16(bytes, offset + 28);
      var extraLen = u16(bytes, offset + 30);
      var commentLen = u16(bytes, offset + 32);
      var localAt = u32(bytes, offset + 42);
      var name = new TextDecoder("utf-8").decode(bytes.subarray(offset + 46, offset + 46 + nameLen));

      entries[name] = { name: name, method: method, compressed: compressed, size: size, localAt: localAt };
      order.push(name);
      offset += 46 + nameLen + extraLen + commentLen;
    }

    var cache = {};

    async function bytesOf(name) {
      if (cache[name]) return cache[name];
      var entry = entries[name];
      if (!entry) throw new Error("no such part: " + name);
      var at = entry.localAt;
      if (u32(bytes, at) !== LOCAL) throw new Error("bad local header: " + name);
      var start = at + 30 + u16(bytes, at + 26) + u16(bytes, at + 28);
      var raw = bytes.subarray(start, start + entry.compressed);
      cache[name] = entry.method === 0 ? raw : await inflate(raw);
      return cache[name];
    }

    /* the same object serves `zip.file(name).async(...)` and `zip.files[name].async(...)` */
    var files = {};
    order.forEach(function (name) {
      files[name] = {
        name: name,
        dir: false,
        size: entries[name].size,
        async: function (type) {
          return bytesOf(name).then(function (data) {
            return decode(data, type);
          });
        },
      };
    });

    return {
      files: files,
      names: order,
      file: function (name) {
        return files[name] || null;
      },
      async: function (name, type) {
        return bytesOf(name).then(function (data) {
          return decode(data, type || "string");
        });
      },
    };
  }

  root.HFPZip = { load: load };
})(typeof globalThis !== "undefined" ? globalThis : this);
