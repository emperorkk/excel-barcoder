// QR Code generator — pure JS, no dependencies
// Supports versions 1-10, error correction level M, byte mode

// ---------------------------------------------------------------------------
// GF(256) arithmetic for Reed-Solomon
// ---------------------------------------------------------------------------
const GF_EXP = new Uint8Array(512);
const GF_LOG = new Uint8Array(256);
(function initGF() {
  let x = 1;
  for (let i = 0; i < 255; i++) {
    GF_EXP[i] = x;
    GF_LOG[x] = i;
    x = x & 0x80 ? (x << 1) ^ 0x11D : x << 1;
    x &= 0xFF;
  }
  for (let i = 255; i < 512; i++) GF_EXP[i] = GF_EXP[i - 255];
})();

const gfMul = (a, b) => a && b ? GF_EXP[(GF_LOG[a] + GF_LOG[b]) % 255] : 0;

function gfPolyMul(p, q) {
  const r = new Uint8Array(p.length + q.length - 1);
  for (let i = 0; i < p.length; i++)
    for (let j = 0; j < q.length; j++)
      r[i + j] ^= gfMul(p[i], q[j]);
  return r;
}

function rsGenPoly(n) {
  let g = new Uint8Array([1]);
  for (let i = 0; i < n; i++) g = gfPolyMul(g, new Uint8Array([1, GF_EXP[i]]));
  return g;
}

function rsEncode(data, nec) {
  const g = rsGenPoly(nec);
  const msg = new Uint8Array(data.length + nec);
  msg.set(data);
  for (let i = 0; i < data.length; i++) {
    const c = msg[i];
    if (c) for (let j = 0; j < g.length; j++) msg[i + j] ^= gfMul(g[j], c);
  }
  return msg.subarray(data.length);
}

// ---------------------------------------------------------------------------
// Block specification for level M — [dataBytes, ecBytes] per block
// Source: ISO/IEC 18004 Annex I
// ---------------------------------------------------------------------------
const BLOCK_SPEC = {
  1:  [[16,10]],
  2:  [[28,16]],
  3:  [[22,13],[22,13]],
  4:  [[32,18],[32,18]],
  5:  [[43,24],[43,24]],
  6:  [[27,16],[27,16],[27,16],[27,16]],
  7:  [[31,18],[31,18],[31,18],[31,18]],
  8:  [[38,22],[38,22],[39,22],[39,22]],
  9:  [[36,22],[36,22],[36,22],[37,22],[37,22]],
  10: [[43,26],[43,26],[43,26],[43,26],[44,26]],
};
const DATA_CAP = Object.fromEntries(
  Object.entries(BLOCK_SPEC).map(([v, bs]) => [+v, bs.reduce((s, [d]) => s + d, 0)])
);

// ---------------------------------------------------------------------------
// Byte-mode data encoding
// ---------------------------------------------------------------------------
function encodeData(text, version) {
  const bytes = new TextEncoder().encode(text);
  const bits = [0,1,0,0]; // mode = byte
  const len = bytes.length;
  for (let i = 7; i >= 0; i--) bits.push((len >> i) & 1); // 8-bit count
  for (const b of bytes) for (let i = 7; i >= 0; i--) bits.push((b >> i) & 1);
  const cap = DATA_CAP[version] * 8;
  // Terminator
  for (let i = 0; i < 4 && bits.length < cap; i++) bits.push(0);
  while (bits.length % 8) bits.push(0);
  // Pad bytes 0xEC/0x11
  const pads = [0xEC, 0x11];
  let pi = 0;
  while (bits.length < cap) {
    const p = pads[pi++ % 2];
    for (let i = 7; i >= 0; i--) bits.push((p >> i) & 1);
  }
  const out = new Uint8Array(bits.length / 8);
  for (let i = 0; i < out.length; i++) {
    let b = 0;
    for (let j = 0; j < 8; j++) b = (b << 1) | bits[i * 8 + j];
    out[i] = b;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Interleave data and error-correction codewords
// ---------------------------------------------------------------------------
function buildCodewords(data, version) {
  const blocks = BLOCK_SPEC[version];
  const dataBlocks = [], ecBlocks = [];
  let pos = 0;
  for (const [dc, ec] of blocks) {
    const d = data.subarray(pos, pos + dc);
    dataBlocks.push(d);
    ecBlocks.push(rsEncode(d, ec));
    pos += dc;
  }
  const cw = [];
  const maxD = Math.max(...dataBlocks.map(b => b.length));
  for (let i = 0; i < maxD; i++) for (const b of dataBlocks) if (i < b.length) cw.push(b[i]);
  const maxE = Math.max(...ecBlocks.map(b => b.length));
  for (let i = 0; i < maxE; i++) for (const b of ecBlocks) if (i < b.length) cw.push(b[i]);
  return cw;
}

// ---------------------------------------------------------------------------
// Matrix construction
// ---------------------------------------------------------------------------
const UNSET = -1;

function newMatrix(n) {
  return Array.from({ length: n }, () => new Int8Array(n).fill(UNSET));
}

function placeFinderPattern(m, r, c) {
  const n = m.length;
  for (let dr = -1; dr <= 7; dr++) for (let dc = -1; dc <= 7; dc++) {
    const row = r + dr, col = c + dc;
    if (row < 0 || row >= n || col < 0 || col >= n) continue;
    if (dr < 0 || dr > 6 || dc < 0 || dc > 6) { m[row][col] = 0; continue; } // separator
    const isBox = (dr >= 2 && dr <= 4 && dc >= 2 && dc <= 4);
    const isBorder = dr === 0 || dr === 6 || dc === 0 || dc === 6;
    m[row][col] = (isBorder || isBox) ? 1 : 0;
  }
}

// Alignment pattern centre positions (from ISO 18004 Table E.1)
const ALIGN_CENTERS = [
  [], [], [18], [22], [26], [30], [34],
  [22, 38], [24, 42], [26, 46], [28, 50],
];

function placeAlignmentPatterns(m, version) {
  const positions = ALIGN_CENTERS[version];
  if (!positions || positions.length === 0) return;
  for (const r of positions) for (const c of positions) {
    // Skip if any module in the 5×5 area is already set (overlap with finder)
    let overlap = false;
    for (let dr = -2; dr <= 2 && !overlap; dr++)
      for (let dc = -2; dc <= 2 && !overlap; dc++)
        if (m[r + dr][c + dc] !== UNSET) overlap = true;
    if (overlap) continue;
    for (let dr = -2; dr <= 2; dr++) for (let dc = -2; dc <= 2; dc++) {
      const edge = Math.abs(dr) === 2 || Math.abs(dc) === 2;
      m[r + dr][c + dc] = (edge || (dr === 0 && dc === 0)) ? 1 : 0;
    }
  }
}

// Compute 15-bit format information (EC level M, given mask pattern)
const FMT_MASK = 0b101010000010010;
function formatBits(maskNum) {
  // EC level M = 0b00 in format info (ISO 18004 Table 12), mask = 3-bit maskNum
  // Format data = (ec_indicator << 3) | mask; EC M indicator = 00
  const data5 = (0b00 << 3) | maskNum;
  let d = data5 << 10;
  const g = 0b10100110111; // generator x^10+x^8+x^5+x^4+x^2+x+1
  for (let i = 14; i >= 10; i--) if (d & (1 << i)) d ^= g << (i - 10);
  return ((data5 << 10) | (d & 0x3FF)) ^ FMT_MASK;
}

function placeFormatInfo(m, maskNum) {
  const n = m.length;
  const fmt = formatBits(maskNum);

  // Primary copy: row 8 (horizontal) and column 8 (vertical)
  const PRIMARY = [
    [8,0],[8,1],[8,2],[8,3],[8,4],[8,5],[8,7],[8,8],
    [7,8],[5,8],[4,8],[3,8],[2,8],[1,8],[0,8],
  ];
  for (let i = 0; i < 15; i++) {
    const b = (fmt >> (14 - i)) & 1;
    m[PRIMARY[i][0]][PRIMARY[i][1]] = b;
  }

  // Top-right copy: row 8, cols n-8 to n-2, bits 8-14 (left to right).
  // 7 modules only — col n-1 is part of the finder separator, never written here.
  for (let i = 0; i < 7; i++) {
    m[8][n - 8 + i] = (fmt >> i) & 1;
  }

  // Bottom-left copy: column 8, rows n-7 to n-1 (bits 7-14)
  for (let i = 7; i < 15; i++) {
    m[n - 15 + i][8] = (fmt >> (14 - i)) & 1;
  }

  // Dark module (always 1)
  m[n - 8][8] = 1;
}

function applyMask(mask, r, c, bit) {
  let flip = false;
  switch (mask) {
    case 0: flip = (r + c) % 2 === 0; break;
    case 1: flip = r % 2 === 0; break;
    case 2: flip = c % 3 === 0; break;
    case 3: flip = (r + c) % 3 === 0; break;
    case 4: flip = (Math.floor(r/2) + Math.floor(c/3)) % 2 === 0; break;
    case 5: flip = (r*c)%2 + (r*c)%3 === 0; break;
    case 6: flip = ((r*c)%2 + (r*c)%3) % 2 === 0; break;
    case 7: flip = ((r+c)%2 + (r*c)%3) % 2 === 0; break;
  }
  return flip ? bit ^ 1 : bit;
}

function buildMatrix(version, codewords, maskNum) {
  const n = version * 4 + 17;
  const m = newMatrix(n);

  placeFinderPattern(m, 0, 0);
  placeFinderPattern(m, 0, n - 7);
  placeFinderPattern(m, n - 7, 0);

  // Timing patterns
  for (let i = 8; i < n - 8; i++) {
    if (m[6][i] === UNSET) m[6][i] = i % 2 === 0 ? 1 : 0;
    if (m[i][6] === UNSET) m[i][6] = i % 2 === 0 ? 1 : 0;
  }

  placeAlignmentPatterns(m, version);

  // Reserve format info areas (written after masking)
  for (let i = 0; i <= 8; i++) {
    if (m[8][i] === UNSET) m[8][i] = 0;
    if (m[i][8] === UNSET) m[i][8] = 0;
  }
  for (let i = n - 8; i < n; i++) {
    if (m[8][i] === UNSET) m[8][i] = 0;
    if (m[i][8] === UNSET) m[i][8] = 0;
  }
  m[n - 8][8] = 1; // dark module

  // Place data in zigzag pattern
  let di = 0, bit = 7;
  let cwByte = codewords[0] || 0;
  const nextBit = () => {
    const b = (cwByte >> bit) & 1;
    if (--bit < 0) { bit = 7; cwByte = codewords[++di] || 0; }
    return b;
  };

  let goUp = true;
  for (let col = n - 1; col >= 1; col -= 2) {
    if (col === 6) col = 5; // skip timing column
    for (let ri = 0; ri < n; ri++) {
      const row = goUp ? n - 1 - ri : ri;
      for (let dc = 0; dc < 2; dc++) {
        const c = col - dc;
        if (m[row][c] === UNSET) m[row][c] = applyMask(maskNum, row, c, nextBit());
      }
    }
    goUp = !goUp;
  }

  placeFormatInfo(m, maskNum);
  return m;
}

// ---------------------------------------------------------------------------
// Mask penalty scoring
// ---------------------------------------------------------------------------
function penaltyScore(m) {
  const n = m.length;
  let score = 0;
  for (let r = 0; r < n; r++) {
    for (let c = 0; c <= n - 5; c++) {
      let run = 1;
      while (c + run < n && m[r][c + run] === m[r][c]) run++;
      if (run >= 5) score += 3 + (run - 5);
      c += run - 1;
    }
  }
  for (let c = 0; c < n; c++) {
    for (let r = 0; r <= n - 5; r++) {
      let run = 1;
      while (r + run < n && m[r + run][c] === m[r][c]) run++;
      if (run >= 5) score += 3 + (run - 5);
      r += run - 1;
    }
  }
  for (let r = 0; r < n - 1; r++)
    for (let c = 0; c < n - 1; c++)
      if (m[r][c] === m[r+1][c] && m[r][c] === m[r][c+1] && m[r][c] === m[r+1][c+1])
        score += 3;
  return score;
}

// ---------------------------------------------------------------------------
// Public renderer
// ---------------------------------------------------------------------------
export function getRenderer(value, reqWidth, reqHeight) {
  const text = String(value);
  const byteLen = new TextEncoder().encode(text).length;

  let version = 1;
  while (version <= 10 && DATA_CAP[version] < byteLen + 2) version++;
  if (version > 10) throw new Error('QR: input too long (max ~200 chars)');

  const data = encodeData(text, version);
  const cw   = buildCodewords(data, version);

  let best = null, bestScore = Infinity;
  for (let mask = 0; mask < 8; mask++) {
    const mat = buildMatrix(version, [...cw], mask);
    const s = penaltyScore(mat);
    if (s < bestScore) { bestScore = s; best = mat; }
  }

  const N    = best.length;
  const dim  = Math.max(reqWidth, reqHeight, 180);
  const cell = Math.max(2, Math.floor(dim / (N + 8)));
  const quiet = 4 * cell;
  const width  = N * cell + quiet * 2;
  const height = N * cell + quiet * 2;

  return {
    width, height,
    isDark: (x, y) => {
      const col = Math.floor((x - quiet) / cell);
      const row = Math.floor((y - quiet) / cell);
      if (col < 0 || col >= N || row < 0 || row >= N) return false;
      return best[row][col] === 1;
    },
  };
}
