// CODE128 barcode generator — pure JS, no dependencies
// Subset B (printable ASCII 32-127), auto-switches to Subset C for digit pairs

// ---------------------------------------------------------------------------
// Symbol table — defined as bar/space width strings, converted to binary
// Format: 6 chars alternating bar-space widths (each 1-4), total must = 11
// Source: ISO/IEC 15417 (CODE 128)
// ---------------------------------------------------------------------------
const WIDTHS = [
  '212222','222122','222221','121223','121322','131222','122213','122312', // 0-7
  '132212','221213','221312','231211','112232','122132','122231','113222', // 8-15
  '123122','123221','213221','221132','221231','213212','223112','312131', // 16-23
  '311222','321122','321221','312212','322112','322211','212123','212321', // 24-31
  '232121','111323','131123','131321','112313','132113','132311','211313', // 32-39
  '231113','231311','112133','112331','132131','113123','113321','133121', // 40-47
  '313121','211331','231131','213113','213311','213131','311123','311321', // 48-55
  '331121','312113','312311','332111','314111','221411','431111','111224', // 56-63
  '111422','121124','121421','141122','141221','112214','112412','122114', // 64-71
  '122411','142112','142211','241211','221114','413111','241112','134111', // 72-79
  '111242','121142','121241','114212','124112','124211','411212','421112', // 80-87
  '421211','212141','214121','412121','111143','111341','131141','114113', // 88-95
  '114311','411113','411311','113141','114131','311141','411131','211412', // 96-103
  '211214','211232',                                                        // 104=StartB, 105=StartC
];
// Stop (106): 13-module special pattern including termination bar
const STOP_BINARY = '1100011101011';

function widthsToBinary(w) {
  let b = '';
  for (let i = 0; i < w.length; i++) b += (i % 2 === 0 ? '1' : '0').repeat(+w[i]);
  return b;
}
const SYM = WIDTHS.map(widthsToBinary);

// ---------------------------------------------------------------------------
// Encoding
// ---------------------------------------------------------------------------
function encodeSubsetB(text) {
  const syms = [];
  syms.push(104); // Start B
  let check = 104;
  for (let i = 0; i < text.length; i++) {
    const v = text.charCodeAt(i) - 32;
    if (v < 0 || v > 95) throw new Error(`Code128: unsupported char U+${text.charCodeAt(i).toString(16)}`);
    syms.push(v);
    check += v * (i + 1);
  }
  syms.push(check % 103); // checksum
  return syms;
}

function encodeSubsetC(text) {
  // Only usable for even-length all-digit strings
  const syms = [];
  syms.push(105); // Start C
  let check = 105;
  for (let i = 0; i < text.length; i += 2) {
    const v = parseInt(text.substring(i, i + 2), 10);
    syms.push(v);
    check += v * (i / 2 + 1);
  }
  syms.push(check % 103);
  return syms;
}

function encode(text) {
  const allDigits = /^\d+$/.test(text) && text.length % 2 === 0 && text.length >= 4;
  return allDigits ? encodeSubsetC(text) : encodeSubsetB(text);
}

// ---------------------------------------------------------------------------
// Bitmap font 5×7 for label text (subset of printable ASCII)
// Each glyph is a 35-bit string (row-major, top-left first)
// ---------------------------------------------------------------------------
const FONT5 = {
  ' ':'00000000000000000000000000000000000',
  '0':'01110100011000110001100011000101110',
  '1':'00100011000100001000010000100001110',
  '2':'01110100010000100010001000100011111',
  '3':'11111000010001000010000011000101110',
  '4':'00010001100101010001111110001000010',
  '5':'11111100001111000001000011000101110',
  '6':'00110010001000011110100011000101110',
  '7':'11111000010001000100010000100001000',
  '8':'01110100011000101110100011000101110',
  '9':'01110100011000101111000010001001100',
  'A':'00100010101000110001111111000110001',
  'B':'11110100011000111110100011000111110',
  'C':'01110100011000010000100001000101110',
  'D':'11100100101000110001100011001011100',
  'E':'11111100001000011110100001000011111',
  'F':'11111100001000011110100001000010000',
  'G':'01110100011000010111100011000101111',
  'H':'10001100011000111111100011000110001',
  'I':'01110001000010000100001000010001110',
  'J':'00111000100001000010000110001001100',
  'K':'10001100101010011000101001001010001',
  'L':'10000100001000010000100001000011111',
  'M':'10001110111010110101100011000110001',
  'N':'10001100011100110101100111000110001',
  'O':'01110100011000110001100011000101110',
  'P':'11110100011000111110100001000010000',
  'Q':'01110100011000110001100011010101101',
  'R':'11110100011000111110100101001010001',
  'S':'01111100001000001110000010000111110',
  'T':'11111001000010000100001000010000100',
  'U':'10001100011000110001100011000101110',
  'V':'10001100011000110001010100010000100',
  'W':'10001100011000110001101011010100010',
  'X':'10001100010101000100010101000110001',
  'Y':'10001100010101000100001000010000100',
  'Z':'11111000010001000100010001000011111',
  '!':'00100001000010000100001000000000100',
  '-':'00000000000000011111000000000000000',
  '_':'00000000000000000000000001111111111',
  '.':'00000000000000000000000000000000100',
  '/':'00001000010001000100010001000000000',
  ':':'00000001000010000000001000010000000',
  '@':'01110100011011110101101111000001110',
};

function charBits(ch) {
  return FONT5[ch] || FONT5[' '];
}

// ---------------------------------------------------------------------------
// Renderer
// ---------------------------------------------------------------------------
export function getRenderer(value, reqWidth, reqHeight) {
  const text = String(value);
  let syms;
  try { syms = encode(text); } catch(e) { throw new Error(e.message); }

  // Build module array
  const modules = [];
  const QUIET = 10;
  for (let i = 0; i < QUIET; i++) modules.push(0);
  for (const s of syms) {
    for (const b of SYM[s]) modules.push(+b);
  }
  for (const b of STOP_BINARY) modules.push(+b); // stop (13 modules)
  for (let i = 0; i < QUIET; i++) modules.push(0);

  const N = modules.length;
  const width  = Math.max(reqWidth,  180, N * 2); // at least 2px/module
  const height = Math.max(reqHeight,  90);
  const modW   = width / N;

  const textH  = Math.max(10, height * 0.15);
  const barH   = height - textH;
  const textCY = barH + textH * 0.52;

  // Label: truncate if needed
  const label  = text.length > 28 ? text.substring(0, 27) + '…' : text;
  const lScale = Math.max(1, Math.floor(Math.min(textH * 0.8, (width * 0.85) / (label.length * 6)) ));
  const cW = 5 * lScale, cH = 7 * lScale;
  const totalLW = label.length * (cW + lScale);
  const lStartX = Math.round((width - totalLW) / 2);
  const lStartY = Math.round(textCY - cH / 2);

  function isBarDark(x, y) {
    if (y >= barH) return false;
    const idx = Math.floor(x / modW);
    return idx >= 0 && idx < N && modules[idx] === 1;
  }

  function isLabelDark(x, y) {
    const ci = Math.floor((x - lStartX) / (cW + lScale));
    if (ci < 0 || ci >= label.length) return false;
    const bits = charBits(label[ci].toUpperCase());
    const col = Math.floor((x - lStartX - ci * (cW + lScale)) / lScale);
    const row = Math.floor((y - lStartY) / lScale);
    if (col < 0 || col >= 5 || row < 0 || row >= 7) return false;
    return bits[row * 5 + col] === '1';
  }

  return {
    width, height,
    isDark: (x, y) => isBarDark(x, y) || isLabelDark(x, y),
  };
}
