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

  async function inflate(bytes, maxOutput, account) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("this browser cannot inflate a zip entry");
    }
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    var reader = new Response(stream).body.getReader();
    var chunks = [];
    var total = 0;
    try {
      while (true) {
        var result = await reader.read();
        if (result.done) break;
        var chunk = result.value;
        if (total + chunk.length > maxOutput || !account(chunk.length)) {
          try { await reader.cancel(); } catch (_) {}
          throw new Error("zip entry exceeds size limit");
        }
        total += chunk.length;
        chunks.push(chunk);
      }
      var output = new Uint8Array(total);
      var offset = 0;
      for (var i = 0; i < chunks.length; i++) {
        output.set(chunks[i], offset);
        offset += chunks[i].length;
      }
      return output;
    } catch (e) {
      account(-total);
      throw e;
    }
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

  async function load(buffer, options) {
    options = options || {};
    var maxEntry = options.maxEntry === undefined ? 64 * 1024 * 1024 : options.maxEntry;
    var maxTotal = options.maxTotal === undefined ? 256 * 1024 * 1024 : options.maxTotal;
    if (typeof maxEntry !== "number" || !isFinite(maxEntry) || maxEntry < 0 ||
        typeof maxTotal !== "number" || !isFinite(maxTotal) || maxTotal < 0) {
      throw new Error("invalid zip size limit");
    }
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var eocd = findEocd(bytes);
    if (eocd < 0) throw new Error("not a zip file");

    var count = u16(bytes, eocd + 10);
    var offset = u32(bytes, eocd + 16);
    if (offset >= bytes.length) throw new Error("bad central directory offset");
    var entries = Object.create(null);
    var order = [];
    var declaredTotal = 0;

    for (var n = 0; n < count; n++) {
      if (offset + 46 > bytes.length || u32(bytes, offset) !== CENTRAL) {
        throw new Error("bad central directory");
      }
      var method = u16(bytes, offset + 10);
      var compressed = u32(bytes, offset + 20);
      var size = u32(bytes, offset + 24);
      var nameLen = u16(bytes, offset + 28);
      var extraLen = u16(bytes, offset + 30);
      var commentLen = u16(bytes, offset + 32);
      var localAt = u32(bytes, offset + 42);
      var nextOffset = offset + 46 + nameLen + extraLen + commentLen;
      if (nextOffset > bytes.length) throw new Error("bad central directory entry");
      if (method !== 0 && method !== 8) throw new Error("unsupported zip compression method: " + method);
      if (size > maxEntry) throw new Error("zip entry exceeds size limit");
      declaredTotal += size;
      if (declaredTotal > maxTotal) throw new Error("zip exceeds total size limit");
      if (localAt + 30 > bytes.length) throw new Error("bad local offset");
      var dataAt = localAt + 30 + u16(bytes, localAt + 26) + u16(bytes, localAt + 28);
      if (dataAt > bytes.length || dataAt + compressed > bytes.length) {
        throw new Error("bad local data offset");
      }
      var name = new TextDecoder("utf-8").decode(bytes.subarray(offset + 46, offset + 46 + nameLen));

      entries[name] = { name: name, method: method, compressed: compressed, size: size, localAt: localAt };
      order.push(name);
      offset = nextOffset;
    }

    var cache = Object.create(null);
    var actualTotal = 0;

    async function bytesOf(name) {
      if (cache[name]) return cache[name];
      var entry = entries[name];
      if (!entry) throw new Error("no such part: " + name);
      cache[name] = (async function () {
        var at = entry.localAt;
        if (at + 30 > bytes.length || u32(bytes, at) !== LOCAL) throw new Error("bad local header: " + name);
        var start = at + 30 + u16(bytes, at + 26) + u16(bytes, at + 28);
        var end = start + entry.compressed;
        if (start > bytes.length || end > bytes.length) throw new Error("bad local data offset: " + name);
        var raw = bytes.subarray(start, end);
        if (entry.method === 0) {
          if (raw.length > maxEntry || actualTotal + raw.length > maxTotal) {
            throw new Error("zip entry exceeds size limit");
          }
          actualTotal += raw.length;
          return raw;
        }
        return inflate(raw, maxEntry, function (length) {
          if (actualTotal + length > maxTotal) return false;
          actualTotal += length;
          return true;
        });
      })();
      return cache[name];
    }

    /* the same object serves `zip.file(name).async(...)` and `zip.files[name].async(...)` */
    var files = Object.create(null);
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
