// Minimal ZIP reader/writer for Office files (docx/xlsx/pptx are ZIPs).
// Reading uses the platform's native DecompressionStream("deflate-raw"),
// writing stores files uncompressed (fast, no CPU-heavy compression).

const enc = new TextEncoder();

let CRC_TABLE;
function crc32(bytes) {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * files: [{ name, data: string | Uint8Array }]  ->  Uint8Array (zip)
 */
export function createZip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;
  // fixed timestamp 2026-01-01 00:00
  const time = 0, date = ((2026 - 1980) << 9) | (1 << 5) | 1;
  for (const f of files) {
    const name = enc.encode(f.name);
    const data = typeof f.data === "string" ? enc.encode(f.data) : f.data;
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // UTF-8 names
    local.setUint16(8, 0, true);      // stored
    local.setUint16(10, time, true);
    local.setUint16(12, date, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true);
    local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);

    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true);
    c.setUint16(4, 20, true);
    c.setUint16(6, 20, true);
    c.setUint16(8, 0x0800, true);
    c.setUint16(10, 0, true);
    c.setUint16(12, time, true);
    c.setUint16(14, date, true);
    c.setUint32(16, crc, true);
    c.setUint32(20, data.length, true);
    c.setUint32(24, data.length, true);
    c.setUint16(28, name.length, true);
    c.setUint32(42, offset, true);
    central.push(new Uint8Array(c.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const centralSize = central.reduce((n, a) => n + a.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);
  const all = [...chunks, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((n, a) => n + a.length, 0));
  let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

async function inflateRaw(bytes, maxSize) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const reader = stream.getReader();
  const parts = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > maxSize) { await reader.cancel(); throw new Error("zip_entry_too_large"); }
    parts.push(value);
  }
  const out = new Uint8Array(total);
  let p = 0;
  for (const v of parts) { out.set(v, p); p += v.length; }
  return out;
}

/**
 * Lists entries of a zip. Returns { names: string[], read(name) -> Promise<Uint8Array|null> }.
 * maxEntrySize protects against zip bombs.
 */
export function openZip(buffer, { maxEntrySize = 30_000_000 } = {}) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65_557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not_a_zip");
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const entries = new Map();
  for (let k = 0; k < count && p + 46 <= bytes.length; k++) {
    if (view.getUint32(p, true) !== 0x02014b50) break;
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const size = view.getUint32(p + 24, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    entries.set(name, { method, compSize, size, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return {
    names: [...entries.keys()],
    async read(name) {
      const e = entries.get(name);
      if (!e) return null;
      if (e.size > maxEntrySize) throw new Error("zip_entry_too_large");
      const lo = e.localOffset;
      if (view.getUint32(lo, true) !== 0x04034b50) throw new Error("bad_zip");
      const start = lo + 30 + view.getUint16(lo + 26, true) + view.getUint16(lo + 28, true);
      const data = bytes.subarray(start, start + e.compSize);
      if (e.method === 0) return data;
      if (e.method === 8) return inflateRaw(data, maxEntrySize);
      throw new Error("unsupported_zip_method");
    },
    async text(name) {
      const d = await this.read(name);
      return d ? dec.decode(d) : null;
    },
  };
}
