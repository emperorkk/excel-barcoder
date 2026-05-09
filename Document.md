# Excel Barcode Inserter — Technical Documentation

## Table of Contents

1. [Overview](#1-overview)
2. [Repository Layout](#2-repository-layout)
3. [Runtime & Infrastructure](#3-runtime--infrastructure)
4. [Worker Entry Point — `src/index.js`](#4-worker-entry-point--srcindexjs)
5. [Barcode Renderers](#5-barcode-renderers)
   - [EAN-13 — `src/ean13.js`](#51-ean-13--srcean13js)
   - [CODE128 — `src/code128.js`](#52-code128--srccode128js)
   - [QR Code — `src/qrcode.js`](#53-qr-code--srcqrcodejs)
6. [PNG Encoder — `src/png.js`](#6-png-encoder--srcpngjs)
7. [Office Add-in (Worker-served)](#7-office-add-in-worker-served)
   - [Dynamic Manifest](#71-dynamic-manifest)
   - [Task Pane UI — `public/taskpane.html`](#72-task-pane-ui--publictaskpanehtml)
   - [Task Pane Logic — `public/taskpane.js`](#73-task-pane-logic--publictaskpanejs)
   - [Commands Stub — `public/commands.html`](#74-commands-stub--publiccommandshtml)
8. [Standalone Add-in Package — `excel-addin/`](#8-standalone-add-in-package--excel-addin)
9. [Installation Guides](#9-installation-guides)
10. [REST API Reference](#10-rest-api-reference)
11. [LAMBDA Formula](#11-lambda-formula)
12. [Dark Mode — CSS Audit](#12-dark-mode--css-audit)
13. [Configuration Files](#13-configuration-files)
14. [Data Flow Diagrams](#14-data-flow-diagrams)
15. [Security & Privacy](#15-security--privacy)
16. [Known Limitations](#16-known-limitations)

---

## 1. Overview

**Excel Barcode Inserter** is a Cloudflare Worker that serves two concerns from a
single deployed URL:

- A **REST barcode API** (`GET /barcode`) that returns a PNG image for EAN-13,
  CODE128, or QR Code input, rendered entirely in JavaScript with no npm runtime
  dependencies.
- A **Microsoft Office Add-in** task pane that calls that same API and embeds the
  returned PNG as a Base64 shape object directly inside an `.xlsx` workbook via the
  Excel JavaScript API (`sheet.shapes.addImage()`).

The worker also serves bilingual installation guides (English + Greek) as static HTML
with responsive layout and WCAG-compliant dark mode.

---

## 2. Repository Layout

```
excel-barcoder/
│
├── src/                          Cloudflare Worker source (ES modules)
│   ├── index.js                  Entry point: routing, CORS, manifest generator
│   ├── ean13.js                  EAN-13 barcode encoder + rasteriser
│   ├── code128.js                CODE128 encoder + rasteriser
│   ├── qrcode.js                 QR Code encoder (GF(256) Reed-Solomon) + rasteriser
│   └── png.js                    Uncompressed PNG encoder
│
├── public/                       Static files — served via Cloudflare ASSETS binding
│   ├── taskpane.html             Add-in task pane (full feature set)
│   ├── taskpane.js               Task pane JS (Office.js / Excel API)
│   ├── commands.html             Add-in FunctionFile stub
│   ├── guide.html                Installation guide (English)
│   └── guide-el.html             Installation guide (Greek / Ελληνικά)
│
├── excel-addin/                  Standalone add-in for manual deployment
│   ├── manifest.xml              Static manifest (WORKER_URL placeholder)
│   └── src/taskpane/
│       ├── taskpane.html         Simplified task pane (no bulk, no size modes)
│       └── taskpane.js           Task pane JS with hardcoded WORKER_URL constant
│
├── wrangler.toml                 Primary Cloudflare Worker config
├── wrangler.jsonc                Alternate/legacy config (not primary)
└── package.json                  { wrangler ^3 as devDependency }
```

---

## 3. Runtime & Infrastructure

| Property | Value |
|----------|-------|
| Runtime | Cloudflare Workers (V8 isolate, ES modules mode) |
| Compatibility date | `2024-09-23` |
| Worker name | `excel-barcoder` |
| Entry module | `src/index.js` |
| Static assets | `./public/` served via `[assets]` binding named `ASSETS` |
| Runtime npm dependencies | **None** |
| Dev dependency | `wrangler ^3` |
| Deploy command | `wrangler deploy` (or `npx wrangler deploy`) |

All four `src/` modules are ES modules bundled together by Wrangler at deploy time.
The `public/` directory is uploaded to Cloudflare's asset storage and served with
automatic MIME types, ETags, and cache headers.

---

## 4. Worker Entry Point — `src/index.js`

### 4.1 CORS headers

Every dynamic response includes:

```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, OPTIONS
Access-Control-Allow-Headers: Content-Type
```

`OPTIONS` requests return `204 No Content` immediately.

### 4.2 Route table

| Condition | Action |
|-----------|--------|
| `OPTIONS *` | 204 CORS preflight |
| `GET /barcode` | Barcode API (→ §10) |
| `GET /manifest.xml` | Dynamic manifest XML (→ §7.1) |
| `GET /assets/icon-*` | 1×1 blue placeholder PNG |
| `GET /` or `GET /guide` | 302 → `/guide.html` |
| `GET /guide-el` | 302 → `/guide-el.html` |
| everything else | `_env.ASSETS.fetch(request)` |

### 4.3 Type normalisation

Before dispatching to a renderer, the `type` query parameter is uppercased and
stripped of `-`, `_`, and spaces:

| Normalised | Maps to |
|-----------|---------|
| `EAN13` | EAN-13 renderer |
| `CODE128`, `CODE128B`, `C128` | CODE128 renderer |
| `QR`, `QRCODE` | QR Code renderer |

Unknown types return a 400 JSON error.

### 4.4 Barcode response

On success, the worker calls `encodePNG(renderer.width, renderer.height, renderer.isDark)`
and returns:

```
HTTP/1.1 200 OK
Content-Type: image/png
Cache-Control: public, max-age=86400
Access-Control-Allow-Origin: *
X-Barcode-Width: <actual px>
X-Barcode-Height: <actual px>
```

### 4.5 Placeholder icon

A hardcoded 1×1 blue PNG (Base64 constant `ICON_B64`) is decoded on each `/assets/icon-*`
request and returned as `image/png`. This satisfies the manifest's icon resource
references without requiring real icon files in the repository.

---

## 5. Barcode Renderers

All three renderers share the same interface:

```ts
getRenderer(value: string, reqWidth: number, reqHeight: number): {
  width:  number;                          // actual image width (≥ minimum)
  height: number;                          // actual image height (≥ minimum)
  isDark: (x: number, y: number) => boolean;  // pixel colour callback
}
```

`encodePNG` calls `isDark(x, y)` for every pixel coordinate. `true` → black `(0,0,0)`;
`false` → white `(255,255,255)`.

---

### 5.1 EAN-13 — `src/ean13.js`

#### Symbol structure

EAN-13 encodes 13 decimal digits. The first (number system) digit selects one of 10
parity patterns that control how the six left-group digits are encoded (L-code vs
G-code). The right six digits always use R-code.

| Table | Description |
|-------|-------------|
| `L[]` | 7-bit bar patterns for left-group digits (normal parity) |
| `G[]` | 7-bit patterns with inverted parity for left-group |
| `R[]` | 7-bit patterns for right-group digits |
| `PARITY[]` | 10 strings of `'L'`/`'G'` — one per first digit |

Module sequence (113 modules total, including quiet zones):

```
[11 quiet] [101 left-guard] [6×7 left data] [01010 centre-guard] [6×7 right data] [101 right-guard] [7 quiet]
```

#### Check digit

Computed per ISO/IEC 15420 (GTIN-13): alternating weights 1 and 3 from left
(1-indexed), modulo 10 complement. If 12 digits are supplied the check digit is
appended automatically; if 13 are supplied it is validated.

#### Rasterisation

- Minimum dimensions enforced: **180 × 90 px**.
- Bottom **15 %** of height is the text area; guard bar extensions reach **45 %** of
  `textH` into it.
- Digit labels rendered using a built-in **5×7 monochrome bitmap font** (35-bit strings,
  row-major).
- Label scale: `max(1, floor(fontSize / 7))` — labels grow proportionally with image
  height.
- Digit 0 centred at module 5.5; digits 1–6 centred in their 7-module groups starting
  at module 14; digits 7–12 starting at module 61.

---

### 5.2 CODE128 — `src/code128.js`

#### Symbol table

106 entries (indices 0–105) defined as 6-character bar/space width strings (each digit
1–4, total must equal 11). Converted to binary at module load time via `widthsToBinary`.

Special symbols: Start B = 104, Start C = 105. Stop pattern = 13-module
`1100011101011` (fixed binary, not in the table).

#### Encoding subsets

| Subset | Input | Start symbol |
|--------|-------|-------------|
| B | Any printable ASCII (codes 32–127) | 104 |
| C | Even-length all-digit strings, length ≥ 4 | 105 |

Auto-selects Subset C when `text` matches `/^\d+$/` and has even length ≥ 4, producing
a narrower barcode. Otherwise Subset B.

**Checksum:** `(startValue + Σ symbolValue × position) % 103`, where position starts
at 1 for the first data symbol.

Module array:
```
[10 quiet] [start] [data symbols] [checksum symbol] [stop (13 bits)] [10 quiet]
```

Minimum width: `max(reqWidth, 180, N × 2)` — enforces at least 2 px per module.

#### Label rendering

- Truncated to 28 characters with `…` if longer.
- 5×7 bitmap font covers uppercase A–Z, digits 0–9, and common punctuation
  (space, `!`, `-`, `_`, `.`, `/`, `:`, `@`).
- Label scale: `max(1, floor(min(textH × 0.8, width × 0.85 / (chars × 6))))`.
- Horizontally centred beneath the bars.

---

### 5.3 QR Code — `src/qrcode.js`

#### GF(256) arithmetic

Primitive polynomial: `x⁸ + x⁴ + x³ + x² + 1` (0x11D).  
`GF_EXP[512]` and `GF_LOG[256]` are computed once at module load using a standard
generator `x = 2`.

#### Block specification (level M)

Lookup table `BLOCK_SPEC` maps version (1–10) to an array of `[dataBytes, ecBytes]`
blocks per ISO/IEC 18004 Annex I. `DATA_CAP[version]` holds the total data byte
capacity for each version.

#### Encoding pipeline

1. **Version selection:** smallest version where `DATA_CAP[v] ≥ UTF-8 byte count + 2`.
   Throws if input exceeds version 10 capacity.
2. **Data encoding (`encodeData`):**  
   Mode indicator `0100` (byte mode) + 8-bit byte count + UTF-8 bytes, padded to
   capacity with terminator, byte-alignment zeros, then alternating `0xEC`/`0x11`.
3. **RS interleaving (`buildCodewords`):**  
   Data split across blocks per `BLOCK_SPEC`; each block independently Reed-Solomon
   encoded; data codewords interleaved column-wise, then EC codewords interleaved.
4. **Matrix construction (`buildMatrix`):**  
   Finder patterns (top-left, top-right, bottom-left) with separators, timing
   patterns, alignment patterns (version-dependent centres from `ALIGN_CENTERS`),
   format information areas reserved. Data placed in standard two-column up/down
   zigzag, skipping column 6 (timing). Mask applied per-module.
5. **Mask selection:**  
   All 8 masks tried; penalty scored per ISO/IEC 18004 §7.8.3 (P1 consecutive runs,
   P2 2×2 blocks, P3 finder-like patterns); lowest penalty wins.
6. **Format information (`formatBits`):**  
   EC level M indicator `0b00`, BCH(15,5) error correction, XOR-masked with
   `0b101010000010010`. Written to two copies (row 8/column 8 primary;
   top-right + bottom-left secondary).

#### Rasterisation

- Output is always square: `dim = max(reqWidth, reqHeight, 180)`.
- `cell = max(2, floor(dim / (N + 8)))` px per module; `quiet = 4 × cell` px quiet
  zone on each side.
- Actual image size: `N × cell + 2 × quiet` for both width and height.
- `isDark(x, y)` maps pixel coordinates back to matrix row/column via integer
  division.

---

## 6. PNG Encoder — `src/png.js`

A self-contained PNG encoder; no zlib or compression library used.

### Algorithm

1. **Scanlines:** for each of `H` rows, prepend filter byte `0x00` (None), then
   `W × 3` bytes of RGB (black = `0x000000`, white = `0xFFFFFF`). Total raw buffer:
   `H × (1 + W×3)` bytes.
2. **DEFLATE (stored):** raw buffer split into blocks of ≤ 65535 bytes.
   Each block header: `[isFinal, len_lo, len_hi, ~len_lo, ~len_hi]`.
   No LZ77 or Huffman coding — maximum simplicity, correct output.
3. **Zlib wrapper:** two-byte header `0x78 0x9C` (valid: `0x789C % 31 === 0`),
   DEFLATE blocks, then Adler-32 checksum of the raw scanline data.
4. **PNG chunks:**
   - `IHDR`: width (4 B), height (4 B), bit depth 8, colour type 2 (RGB truecolour),
     5 zero bytes (compression/filter/interlace).
   - `IDAT`: the zlib-wrapped stream.
   - `IEND`: empty.
5. **CRC-32:** computed per chunk over `[4-byte type | data]` using reflected
   polynomial `0xEDB88320` (CRC-32/ISO-HDLC).

### Exported function

```ts
encodePNG(
  width:  number,
  height: number,
  isDark: (x: number, y: number) => boolean,
  scale?: number   // pixel scale factor, default 1
): Uint8Array      // complete PNG file bytes
```

---

## 7. Office Add-in (Worker-served)

### 7.1 Dynamic Manifest

`serveManifest(origin)` in `src/index.js` builds the manifest XML at request time,
substituting `${origin}` for every URL resource. This means the manifest is always
self-consistent regardless of which Cloudflare subdomain or custom domain the worker
is deployed to.

| Manifest field | Value |
|----------------|-------|
| `<Id>` | `A1B2C3D4-E5F6-7890-ABCD-EF1234567890` |
| `<Version>` | `1.0.0` |
| `<ProviderName>` | Excel Barcode |
| `<DisplayName>` | Barcode Inserter |
| `<Host Name>` | Workbook |
| `<Permissions>` | ReadWriteDocument |
| Ribbon group | "Barcodes" on `TabHome` |
| Ribbon button | "Insert Barcode" → `ShowTaskpane` |
| `<SourceLocation>` | `{origin}/taskpane.html` |
| `<FunctionFile>` | `{origin}/commands.html` |
| Icons | `{origin}/assets/icon-{16,32,80}.png` |

---

### 7.2 Task Pane UI — `public/taskpane.html`

Inline CSS styled to match Microsoft Fluent UI (Segoe UI, blue `#106ebe`, 2 px
border-radius). No external stylesheet.

#### Controls

| Element | ID | Description |
|---------|----|-------------|
| `<input type="text">` | `barcodeValue` | Barcode content |
| `<select>` | `barcodeType` | EAN13 / CODE128 / QR |
| `<select>` | `sizeMode` | resize / custom / fitcell |
| `<input type="number">` | `barcodeWidth` | Width px, default 300, min 180 |
| `<input type="number">` | `barcodeHeight` | Height px, default 150, min 90 |
| `<input type="checkbox">` | `useSelectedCell` | Auto-read selected cell (default: checked) |
| `<div>` | `bulkBanner` | Blue info banner when multiple rows selected |
| `<div>` | `preview` | Barcode preview image container |
| `<button>` | `previewBtn` | Fetch and display barcode preview |
| `<button>` | `insertBtn` | Insert single barcode |
| `<button>` | `bulkBtn` | Insert for all selected rows (hidden until multi-row) |
| `<div>` | `status` | Success (green) / error (red) message |
| `<p class="hint">` | `valueHint` | Dynamic per-type hint text |

#### CSS tokens (Fluent palette)

| Role | Value |
|------|-------|
| Primary blue | `#106ebe` |
| Button hover / dark blue | `#005a9e` |
| Bulk button green | `#107c10` |
| Disabled / muted | `#a19f9d` |
| Error red | `#a4262c` |
| Border | `#c8c6c4` |
| Body text | `#323130` |

> The task pane runs inside an Office iframe — no dark mode is applied (Office
> manages its own chrome).

---

### 7.3 Task Pane Logic — `public/taskpane.js`

Depends on `Office.js` (loaded from Microsoft CDN) and the Excel JavaScript API.

#### Initialisation

`Office.onReady` guards execution: only proceeds when `host === Office.HostType.Excel`.
Registers a `DocumentSelectionChanged` handler and immediately calls
`syncFromSelection()`.

#### Selection sync (`syncFromSelection`)

```
Excel.run → getSelectedRange().load(['values','rowCount']) → ctx.sync()
  if rowCount > 1 → show bulkBanner + bulkBtn
  if useSelectedCell.checked → valueInput.value = values[0][0]
```

#### URL builder

```js
const API_BASE = '';   // empty string → relative URL → same worker origin
function barcodeURL(value, type, w, h) {
  return `${API_BASE}/barcode?${new URLSearchParams({ value, type, width: w, height: h })}`;
}
```

#### PNG fetcher (`fetchPNG`)

1. `fetch(url)` → on non-200, parse JSON error body.
2. `resp.arrayBuffer()` → `Uint8Array` → binary string → `btoa()` → Base64.

#### Barcode placement (`placeBarcode`)

Called with a `Range` object already loaded with `left`, `top`, `width`, `height`,
`format.columnWidth`:

| Size mode | Image dimensions | Cell resize? |
|-----------|-----------------|-------------|
| `resize` | `reqW` × `reqH` | Yes — `rowHeight = reqH`, `columnWidth` scaled proportionally |
| `custom` | `reqW` × `reqH` | No |
| `fitcell` | `max(180, cell.width)` × `max(90, cell.height)` | No |

After dimensions are resolved: `fetchPNG` → `sheet.shapes.addImage(b64)` →
set `shape.left/top/width/height/name/lockAspectRatio` → `ctx.sync()`.

Shape name format: `` `Barcode_${type}_${value.substring(0,20)}_${Date.now()}` ``.

#### Bulk insert

Iterates rows **sequentially** (one `ctx.sync` per row — parallel calls would require
pre-loading all cell positions, which complicates the resize logic). Skips empty cells
and cells whose value stringifies as `'false'`. Updates status div with live progress
count.

---

### 7.4 Commands Stub — `public/commands.html`

Required by the manifest `<FunctionFile>` element. Contains only the Office.js script
tag and an empty `Office.onReady(() => {})` callback. No commands are currently
registered.

---

## 8. Standalone Add-in Package — `excel-addin/`

This directory is an alternative to the auto-manifest route — useful when:
- Deploying the add-in independently of the Worker.
- Needing a static XML file for manual editing or IT-managed sideloading.

### `excel-addin/manifest.xml`

Identical structure to the dynamic manifest but with `WORKER_URL` placeholder strings
instead of the auto-filled origin. Differences from the dynamic version:

| Feature | Dynamic (`/manifest.xml`) | Standalone (`excel-addin/manifest.xml`) |
|---------|--------------------------|----------------------------------------|
| Origin URL | Auto-filled from `request.url` | Requires manual `WORKER_URL` substitution |
| GUID | Fixed `A1B2C3D4-…` | Fixed `A1B2C3D4-…` (comment suggests replacing) |
| `<GetStarted>` block | Absent | Present (links to GitHub) |
| Group id | `BarcodeGroup` | `CommandsGroup` |
| Taskpane id | `BarcodePane` | `ButtonId1` |

### `excel-addin/src/taskpane/taskpane.js`

A simplified version of the main task pane — differences from `public/taskpane.js`:

| Feature | `public/taskpane.js` | `excel-addin/taskpane.js` |
|---------|---------------------|--------------------------|
| API base URL | `''` (relative) | `WORKER_URL` constant (hardcoded) |
| Size mode selector | Yes (resize / custom / fitcell) | No — always inserts at specified w×h |
| Bulk insert | Yes | No |
| Bulk banner | Yes | No |
| Cell resize on insert | Yes (resize mode) | No |
| Shape positioning | `cell.left / cell.top` | `range.left / range.top` |

---

## 9. Installation Guides

Two static HTML pages, `public/guide.html` (English) and `public/guide-el.html`
(Greek), sharing identical CSS and structure.

### URL aliases

| URL | Serves |
|-----|--------|
| `/` | 302 → `/guide.html` |
| `/guide` | 302 → `/guide.html` |
| `/guide.html` | English guide (direct) |
| `/guide-el` | 302 → `/guide-el.html` |
| `/guide-el.html` | Greek guide (direct) |

### Page structure

- **Sticky header** — title, badges (Microsoft 365, Excel 2016+, no plug-in), language switcher badge.
- **Sidebar nav** (desktop) / **horizontal nav** (mobile ≤ 700 px) — anchors to all sections.
- **YouTube video embed** — 16:9 responsive wrapper above main content.
- **Main content sections:** Requirements, Install (options A/B), First Barcode,
  Task Pane Guide, Barcode Types, Sizes & Pixels, LAMBDA Formula, REST API,
  Troubleshooting, FAQ.
- **Sticky footer** — links (guide, GitHub, kourentzes.com/konstantinos, onecode.gr,
  manifest.xml) + PayPal donate button.

### Inline JavaScript

Small `<script>` blocks in both guides populate code examples with the live server
origin at runtime:

```js
document.getElementById('manifestUrl').textContent = location.origin + '/manifest.xml';
document.getElementById('lambdaFormula').textContent =
  '=LAMBDA(val,type,w,h,IMAGE("' + location.origin + '/barcode?value="&val&...))';
document.getElementById('ex1').textContent = location.origin + '/barcode?value=…';
```

---

## 10. REST API Reference

### Endpoint

```
GET /barcode
```

### Parameters

| Name | Type | Required | Default | Constraints |
|------|------|----------|---------|-------------|
| `value` | string | yes | — | URL-encoded barcode content |
| `type` | string | no | `EAN13` | Normalised to: `EAN13`, `CODE128`, `QR` |
| `width` | integer | no | `300` | Clamped `[1, 2000]`; per-type minimum applied |
| `height` | integer | no | `150` | Clamped `[1, 1000]`; per-type minimum applied |

### Per-type minimums

| Type | Min width | Min height | Notes |
|------|-----------|------------|-------|
| EAN-13 | 180 px | 90 px | Hard-coded in `ean13.js` |
| CODE128 | `max(180, N×2)` | 90 px | `N` = module count; ensures ≥ 2 px/module |
| QR Code | 180 px (effective) | same as width | Always square; `cell = max(2, floor(dim/(N+8)))` |

### Response headers (200 OK)

```
Content-Type: image/png
Cache-Control: public, max-age=86400
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, OPTIONS
Access-Control-Allow-Headers: Content-Type
X-Barcode-Width: <actual pixel width>
X-Barcode-Height: <actual pixel height>
```

### Error response (400 Bad Request)

```json
{ "error": "EAN-13 requires 12 or 13 digits" }
```

---

## 11. LAMBDA Formula

For Excel 365 with `IMAGE()` function support:

```
=LAMBDA(val,type,w,h,
  IMAGE("https://your-worker.workers.dev/barcode?value="&val
        &"&type="&type&"&width="&w&"&height="&h))
```

Register in **Formulas → Name Manager** as `ADDBARCODE` (Workbook scope).

| Argument | Type | Example |
|----------|------|---------|
| `val` | Cell ref or string | `A1`, `"012345678901"` |
| `type` | String | `"EAN13"`, `"CODE128"`, `"QR"` |
| `w` | Number | `300` |
| `h` | Number | `150` |

**Limitations:**
- Stores a URL, not pixel data — requires internet on every workbook open.
- LAMBDA definition is per-workbook — must be repeated or delivered via a `.xltx`
  template.
- Not available in Excel 2019 or earlier (requires `IMAGE()` function, Excel 365 only).

---

## 12. Dark Mode — CSS Audit

Both guide pages implement `@media (prefers-color-scheme: dark)`.

### CSS variable overrides

| Variable | Light value | Dark value |
|----------|------------|------------|
| `--blue` | `#106ebe` | `#2d9cdb` |
| `--blue-dark` | `#005a9e` | `#1a6fa8` |
| `--green` | `#107c10` | `#27ae60` |
| `--warn-bg` | `#fff4ce` | `#2c1e00` |
| `--note-bg` | `#eff6fc` | `#0d2035` |
| `--border` | `#c8c6c4` | `#3a3a3a` |
| `--text` | `#201f1e` | `#d4d4d4` |
| `--muted` | `#605e5c` | `#9a9a9a` |
| `--bg` | `#faf9f8` | `#1a1a1a` |
| `--white` | `#ffffff` | `#252525` |
| `--code-bg` | `#f3f2f1` | `#2a2a2a` |

### Explicit selector overrides (dark mode block)

| Selector | Override |
|----------|----------|
| `th` | `background: #1a4f7a` |
| `tr:nth-child(even) td` | `background: #222` |
| `.mockup-input`, `.mockup-select` | `color: var(--text)` |
| `footer` | `background: #111` |
| `.tip` | `background: #0a2a15` |
| `.warn` | `background: #2c1e00; color: var(--text)` |
| `h2` | `color: #5ab0d9` |

### Rationale for explicit overrides

| Selector | Problem without override | Fix |
|----------|------------------------|-----|
| `.tip` | Hardcoded `background: #e6f4ea` (light green) — `#d4d4d4` text on near-white = ~1.2:1 contrast | `#0a2a15` dark green → 10.5:1 |
| `.warn` | `--warn-bg: #fff4ce` inherited fine, but explicit `color` needed for text | `color: var(--text)` ensures `#d4d4d4` on `#2c1e00` |
| `h2` | `--blue-dark: #1a6fa8` on `#1a1a1a` = 3.2:1 (borderline large-text only) | `#5ab0d9` → 7.2:1 (WCAG AA+) |

### Hardcoded colours outside variable/dark-mode blocks

| Colour | Usage | Dark mode safe? |
|--------|-------|----------------|
| `#e6f4ea` | `.tip` light-mode background | Yes — overridden to `#0a2a15` |
| `#ffc439` | PayPal donate button background | Yes — intentional brand yellow on dark footer |
| `#f0b429` | Donate button hover | Yes — same context |
| `#003087` | Donate button text | Yes — on `#ffc439` yellow, context-independent |
| `#000` | Video wrapper background | Yes — black bg for video |
| `#fff` | Various `color: #fff` (header, buttons, footer) | Yes — always on dark-coloured backgrounds |

### WCAG 2.1 contrast ratios (dark mode, background `#1a1a1a`)

| Element | Foreground | Background | Ratio | Level |
|---------|-----------|------------|-------|-------|
| Body text | `#d4d4d4` | `#1a1a1a` | 11.7:1 | **AAA** |
| Muted / nav | `#9a9a9a` | `#1a1a1a` | 6.2:1 | **AA** |
| `h2` heading | `#5ab0d9` | `#1a1a1a` | 7.2:1 | **AA** |
| Links / `--blue` | `#2d9cdb` | `#1a1a1a` | 5.7:1 | **AA** |
| `.note` text | `#d4d4d4` | `#0d2035` | 11.1:1 | **AAA** |
| `.tip` text | `#d4d4d4` | `#0a2a15` | 10.5:1 | **AAA** |
| `.warn` text | `#d4d4d4` | `#2c1e00` | 11.0:1 | **AAA** |
| Code / pre text | `#d4d4d4` | `#2a2a2a` | 9.7:1 | **AAA** |
| Table `th` | `#ffffff` | `#1a4f7a` | 8.6:1 | **AAA** |
| Even table row | `#d4d4d4` | `#222222` | 10.7:1 | **AAA** |
| Button text | `#ffffff` | `#1a6fa8` | 5.4:1 | **AA** |
| Footer links | `#ffffff` | `#111111` | 18.9:1 | **AAA** |
| Donate btn text | `#003087` | `#ffc439` | 7.0:1 | **AA** |
| Warn border | `#e0a800` | `#2c1e00` | 7.6:1 | **AA** |

All pairs pass WCAG 2.1 AA. Every element that carries meaningful text passes at
least 4.5:1 (AA normal text) or 3.0:1 (AA large/bold text ≥ 18pt or ≥ 14pt bold).

---

## 13. Configuration Files

### `wrangler.toml` (primary)

```toml
name = "excel-barcoder"
main = "src/index.js"
compatibility_date = "2024-09-23"

[assets]
directory = "./public"
binding   = "ASSETS"
```

`ASSETS` is a Cloudflare `KVNamespace`-like binding exposing a `fetch(request)`
method. It serves files from `./public/` with automatic MIME types, ETags, and
`Cache-Control` headers. When no dynamic route matches, `_env.ASSETS.fetch(request)`
is called as a catch-all.

### `package.json`

```json
{
  "name": "excel-barcode-worker",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "dev":    "wrangler dev",
    "deploy": "wrangler deploy"
  },
  "devDependencies": {
    "wrangler": "^3.0.0"
  }
}
```

---

## 14. Data Flow Diagrams

### Single barcode insert (task pane → workbook)

```
User clicks "Insert into Excel"
    │
    ▼
taskpane.js: insertBtn handler
    │  reads: value, type, w, h from form
    ▼
Excel.run(ctx)
    │  getSelectedRange().load(['left','top','width','height'])
    │  range.format.load('columnWidth')
    ▼
placeBarcode(ctx, sheet, cell, value, type, w, h)
    │
    ▼
fetch( GET /barcode?value=…&type=…&width=…&height=… )
    │        ↑ same Cloudflare Worker origin
    │  Worker: selectRenderer → isDark callback
    │  Worker: encodePNG(width, height, isDark) → Uint8Array
    │  Worker: Response(png, { Content-Type: image/png, … })
    ▼
resp.arrayBuffer() → btoa() → base64 string
    │
    ▼
sheet.shapes.addImage(base64)
    │  shape.left/top/width/height/name set
    │  (resize mode: rowHeight + columnWidth adjusted)
    ▼
ctx.sync() → barcode embedded as shape in workbook
```

### Bulk insert (N rows)

```
User selects A1:A50, clicks "Insert for All Selected Cells"
    │
    ▼
Excel.run → range.load(['values','rowCount']) → ctx.sync()
    │
    for i = 0 … rowCount-1:
    │   skip if values[i][0] is empty or 'false'
    │   range.getCell(i,0).load([…]) → ctx.sync()
    │   placeBarcode(…)  ← one HTTP call per row
    │   showStatus("Processing… N/total")
    ▼
showStatus("Done: N barcodes inserted, M skipped")
```

### REST API direct call (external client)

```
Client: GET /barcode?value=ABC-123&type=CODE128&width=350&height=120
    │
    ▼
Worker fetch handler
    │  selectRenderer('CODE128', 'ABC-123', 350, 120)
    ▼
code128.js: getRenderer('ABC-123', 350, 120)
    │  encode('ABC-123') → Subset B symbol array
    │  build modules: [10 quiet][start B][data][check][stop][10 quiet]
    │  returns { width: 350, height: 120, isDark(x,y) }
    ▼
png.js: encodePNG(350, 120, isDark)
    │  build raw scanlines H×(1 + W×3) bytes
    │  DEFLATE stored blocks
    │  zlib wrap (0x789C + Adler-32)
    │  IHDR + IDAT + IEND chunks with CRC-32
    │  returns Uint8Array
    ▼
Response(Uint8Array, {
  Content-Type: image/png,
  Cache-Control: public, max-age=86400,
  X-Barcode-Width: 350,
  X-Barcode-Height: 120,
  Access-Control-Allow-Origin: *
})
```

---

## 15. Security & Privacy

| Concern | Detail |
|---------|--------|
| **Data persistence** | None. The worker is stateless; barcode values in query parameters are processed in memory and not logged or stored. |
| **CORS policy** | `Access-Control-Allow-Origin: *` — necessary for Excel Online's sandboxed iframe to call the same-origin API. |
| **Authentication** | None. The API is public. Do not encode confidential data (passwords, PII) in barcodes intended for public display — the value is visible in plain text in the request URL and browser network logs. |
| **Edge processing** | Cloudflare Workers run at the nearest PoP. No backend servers, no databases, no persistent storage. |
| **Office permissions** | `ReadWriteDocument` — required by `sheet.shapes.addImage()`. The add-in reads cell values and writes image shapes only; it makes no external network requests beyond calling the worker's own `/barcode` endpoint (same origin). |
| **Content Security** | `excel-addin/taskpane.js` uses a hardcoded `WORKER_URL` constant — operators should verify this points to their own worker before distributing. |

---

## 16. Known Limitations

| Limitation | Detail |
|------------|--------|
| QR version cap | Versions 1–10 only (level M). Maximum ~200 UTF-8 bytes of input. Inputs above this throw `'QR: input too long (max ~200 chars)'`. |
| CODE128 charset | ASCII codes 32–127 only. Accented characters, emoji, or control codes throw `'Code128: unsupported char U+XXXX'`. |
| EAN-8 | Not supported. EAN-13 only (and UPC-A as 12-digit EAN-13 with leading zero). |
| Mobile Excel | `sheet.shapes.addImage()` is not available on Excel for iOS or Android. The task pane opens and shows a preview, but the Insert button has no effect on mobile platforms. |
| PNG file size | DEFLATE stored blocks (no compression). A 400×200 EAN-13 PNG is ~240 KB vs ~20 KB compressed. Not a concern at Cloudflare edge scale, but large workbooks with many embedded barcodes grow accordingly. |
| Bulk insert speed | Rows processed sequentially — one `ctx.sync()` round-trip per row. For 100+ rows this may take 30–60 seconds depending on network latency. |
| LAMBDA per-workbook | The `ADDBARCODE` LAMBDA must be defined in each new workbook's Name Manager. It cannot be deployed globally without an Excel `.xltx` template or M365 admin policy. |
| LAMBDA offline | The `=ADDBARCODE()` formula embeds a URL, not pixel data. The barcode does not display without an internet connection or if the worker URL changes. |
| Standalone manifest | `excel-addin/manifest.xml` requires manual substitution of every `WORKER_URL` placeholder before use — forgetting one results in a broken add-in. |
