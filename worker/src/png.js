// Pure JS PNG encoder — no external dependencies
// Produces uncompressed (DEFLATE stored) RGB PNG

function buildCRCTable() {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
}
const CRC_TABLE = buildCRCTable();

function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xFF];
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function adler32(data) {
  let a = 1, b = 0;
  for (let i = 0; i < data.length; i++) {
    a = (a + data[i]) % 65521;
    b = (b + a) % 65521;
  }
  // b can reach 65520; (b << 16) overflows into a negative signed int.
  // >>> 0 reinterprets as unsigned 32-bit before returning.
  return ((b << 16) | a) >>> 0;
}

function u32be(buf, off, v) {
  buf[off] = (v >>> 24) & 0xFF;
  buf[off + 1] = (v >>> 16) & 0xFF;
  buf[off + 2] = (v >>> 8) & 0xFF;
  buf[off + 3] = v & 0xFF;
}

function concat(arrays) {
  const total = arrays.reduce((n, a) => n + a.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const a of arrays) { out.set(a, pos); pos += a.length; }
  return out;
}

function makeChunk(type, data) {
  const enc = new TextEncoder();
  const tb = enc.encode(type);
  const chunk = new Uint8Array(4 + 4 + data.length + 4);
  u32be(chunk, 0, data.length);
  chunk.set(tb, 4);
  chunk.set(data, 8);
  const crcInput = new Uint8Array(4 + data.length);
  crcInput.set(tb, 0);
  crcInput.set(data, 4);
  u32be(chunk, 8 + data.length, crc32(crcInput));
  return chunk;
}

function deflateStore(raw) {
  const MAX = 65535;
  const blocks = [];
  let pos = 0;
  while (pos < raw.length || pos === 0) {
    const end = Math.min(pos + MAX, raw.length);
    const len = end - pos;
    const isFinal = end >= raw.length ? 1 : 0;
    const blk = new Uint8Array(5 + len);
    blk[0] = isFinal;
    blk[1] = len & 0xFF;
    blk[2] = (len >>> 8) & 0xFF;
    blk[3] = (~len) & 0xFF;
    blk[4] = (~len >>> 8) & 0xFF;
    blk.set(raw.subarray(pos, end), 5);
    blocks.push(blk);
    pos = end;
    if (pos >= raw.length) break;
  }
  return concat(blocks);
}

/**
 * Encode a 2D bitmap as PNG.
 * @param {number} width
 * @param {number} height
 * @param {(x:number,y:number)=>boolean} isDark - returns true for dark (black) pixel
 * @param {number} [scale=1] - pixel scale factor
 */
export function encodePNG(width, height, isDark, scale = 1) {
  const W = width * scale;
  const H = height * scale;
  const stride = W * 3;

  // Build raw scanlines with filter byte 0 prefix
  const raw = new Uint8Array(H * (1 + stride));
  for (let py = 0; py < H; py++) {
    raw[py * (1 + stride)] = 0; // filter: None
    const sy = Math.floor(py / scale);
    for (let px = 0; px < W; px++) {
      const sx = Math.floor(px / scale);
      const v = isDark(sx, sy) ? 0 : 255;
      const off = py * (1 + stride) + 1 + px * 3;
      raw[off] = raw[off + 1] = raw[off + 2] = v;
    }
  }

  const deflated = deflateStore(raw);
  const checksum = adler32(raw);

  const idat = new Uint8Array(2 + deflated.length + 4);
  idat[0] = 0x78; idat[1] = 0x9C; // zlib header (valid: 0x789C % 31 === 0)
  idat.set(deflated, 2);
  u32be(idat, 2 + deflated.length, checksum);

  const ihdrData = new Uint8Array(13);
  u32be(ihdrData, 0, W);
  u32be(ihdrData, 4, H);
  ihdrData[8] = 8; // bit depth
  ihdrData[9] = 2; // color type: RGB

  const sig = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  return concat([
    sig,
    makeChunk('IHDR', ihdrData),
    makeChunk('IDAT', idat),
    makeChunk('IEND', new Uint8Array(0)),
  ]);
}
