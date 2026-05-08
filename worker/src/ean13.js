// EAN-13 barcode generator — pure JS, no dependencies

const L = ['0001101','0011001','0010011','0111101','0100011','0110001','0101111','0111011','0110111','0001011'];
const G = ['0100111','0110011','0011011','0100001','0011101','0111001','0000101','0010001','0001001','0010111'];
const R = ['1110010','1100110','1101100','1000010','1011100','1001110','1010000','1000100','1001000','1110100'];
const PARITY = ['LLLLLL','LLGLGG','LLGGLG','LLGGGL','LGLLGG','LGGLLG','LGGGLL','LGLGLG','LGLGGL','LGGLGL'];

function checkDigit(d) {
  let s = 0;
  for (let i = 0; i < 12; i++) s += +d[i] * (i % 2 ? 3 : 1);
  return (10 - s % 10) % 10;
}

/**
 * Build EAN-13 module array (113 modules including quiet zones).
 * Throws if value is not 12 or 13 digits.
 */
export function buildModules(value) {
  let t = value.toString().replace(/\D/g, '');
  if (t.length === 12) t += checkDigit(t);
  if (t.length !== 13) throw new Error('EAN-13 requires 12 or 13 digits');
  const p = PARITY[+t[0]];
  const m = [];
  for (let i = 0; i < 11; i++) m.push(0);   // left quiet zone
  m.push(1, 0, 1);                            // left guard
  for (let i = 0; i < 6; i++) {
    const code = p[i] === 'L' ? L[+t[i+1]] : G[+t[i+1]];
    for (const b of code) m.push(+b);
  }
  m.push(0, 1, 0, 1, 0);                     // center guard
  for (let i = 0; i < 6; i++) {
    for (const b of R[+t[i+7]]) m.push(+b);
  }
  m.push(1, 0, 1);                            // right guard
  for (let i = 0; i < 7; i++) m.push(0);     // right quiet zone
  return { modules: m, digits: t };           // m.length === 113
}

// Guard module indices (bars that extend below data area)
const GUARD_IDX = new Set([11, 13, 57, 59, 103, 105]);

/**
 * Returns isDark(x, y) for a pixel grid of the given dimensions.
 * minWidth/minHeight are enforced to ensure readability.
 */
export function getRenderer(value, reqWidth, reqHeight) {
  const width  = Math.max(reqWidth,  180);
  const height = Math.max(reqHeight,  90);
  const { modules, digits } = buildModules(value);
  const N = modules.length; // 113
  const modW = width / N;

  // Text area is bottom 15% of height, guard bars extend into it
  const textH   = Math.max(10, height * 0.15);
  const dataBarH = height - textH;
  const extH    = textH * 0.45;
  const fontSize = Math.max(7, textH * 0.72);

  function isDark(x, y) {
    const modIdx = Math.floor(x / modW);
    if (modIdx < 0 || modIdx >= N) return false;
    if (modules[modIdx] === 0) return false;

    const isGuard = GUARD_IDX.has(modIdx);
    const barBottom = isGuard ? dataBarH + extH : dataBarH;
    return y < barBottom;
  }

  // Digit label pixel renderer using a 5×7 bitmap font
  const FONT = {
    '0':'0111010001100011000110001011100',
    '1':'0010001100001000010000100111000',
    '2':'0111010001000100010001000111110',
    '3':'1111000010011100000110001111100',
    '4':'0001000110010101111100010000100',
    '5':'1111110000111100000110001111100',
    '6':'0011010000111101000110001011100',
    '7':'1111100010001000100010000100000',
    '8':'0111010001011100100110001011100',
    '9':'0111010001011110000100010011100',
  };

  // Precompute digit label positions
  // Digit 0: x-center = 5.5 modules, y-center = dataBarH + textH/2
  // Left digits 1-6: centered in their 7-module group starting at module 14
  // Right digits 7-12: starting at module 61
  const labelScale = Math.max(1, Math.floor(fontSize / 7));
  const charW = 5 * labelScale;
  const charH = 7 * labelScale;

  const labels = [];
  const addLabel = (digit, cx, cy) => {
    const bits = FONT[digit];
    if (!bits) return;
    const lx = Math.round(cx - charW / 2);
    const ly = Math.round(cy - charH / 2);
    labels.push({ bits, lx, ly, scale: labelScale });
  };

  const textCY = dataBarH + textH * 0.55;
  addLabel(digits[0], 5.5 * modW, textCY);
  for (let i = 0; i < 6; i++) addLabel(digits[i+1], (14 + i*7 + 3.5) * modW, textCY);
  for (let i = 0; i < 6; i++) addLabel(digits[i+7], (61 + i*7 + 3.5) * modW, textCY);

  function isLabelDark(x, y) {
    for (const { bits, lx, ly, scale } of labels) {
      const col = Math.floor((x - lx) / scale);
      const row = Math.floor((y - ly) / scale);
      if (col < 0 || col >= 5 || row < 0 || row >= 7) continue;
      if (bits[row * 5 + col] === '1') return true;
    }
    return false;
  }

  return {
    width, height,
    isDark: (x, y) => isDark(x, y) || isLabelDark(x, y),
  };
}
