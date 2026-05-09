# Excel Barcode Inserter — Technical Documentation

## Table of Contents

1. [Overview](#overview)
2. [Repository Layout](#repository-layout)
3. [Runtime & Infrastructure](#runtime--infrastructure)
4. [Worker Entry Point — `src/index.js`](#worker-entry-point--srcindexjs)
5. [Barcode Renderers](#barcode-renderers)
   - [EAN-13 — `src/ean13.js`](#ean-13--srcean13js)
   - [CODE128 — `src/code128.js`](#code128--srccode128js)
   - [QR Code — `src/qrcode.js`](#qr-code--srcqrcodejs)
6. [PNG Encoder — `src/png.js`](#png-encoder--srcpngjs)
7. [Office Add-in](#office-add-in)
   - [Manifest](#manifest)
   - [Task Pane UI — `public/taskpane.html`](#task-pane-ui--publictaskpanehtml)
   - [Task Pane Logic — `public/taskpane.js`](#task-pane-logic--publictaskpanejs)
8. [Installation Guides](#installation-guides)
9. [REST API Reference](#rest-api-reference)
10. [LAMBDA Formula](#lambda-formula)
11. [Dark Mode](#dark-mode)
12. [Configuration Files](#configuration-files)
13. [Data Flow Diagrams](#data-flow-diagrams)
14. [Security & Privacy](#security--privacy)
15. [Known Limitations](#known-limitations)

---

## Overview

**Excel Barcode Inserter** is a zero-dependency Cloudflare Worker that:

- Exposes a REST endpoint (`GET /barcode`) returning a PNG barcode image for any of three symbologies: **EAN-13**, **CODE128**, and **QR Code**.
- Serves a **Microsoft Office Add-in** task pane that fetches barcodes from the same worker origin and embeds them as Base64 image objects directly inside `.xlsx` workbook files.
- Provides bilingual installation guides (English and Greek) as static HTML pages with dark-mode support.

The entire barcode rendering pipeline — symbol encoding, pixel rasterisation, and PNG encoding — runs in pure JavaScript inside the Cloudflare Worker runtime with **no npm runtime dependencies**.

---

## Repository Layout

```
excel-barcoder/
│
├── src/                        ← Cloudflare Worker source
│   ├── index.js                ← Entry point: routing, manifest generation, CORS
│   ├── ean13.js                ← EAN-13 encoder + rasteriser
│   ├── code128.js              ← CODE128 encoder + rasteriser
│   ├── qrcode.js               ← QR Code encoder (Reed-Solomon ECC) + rasteriser
│   └── png.js                  ← PNG encoder (no zlib library)
│
├── public/                     ← Static assets served via Cloudflare ASSETS binding
│   ├── taskpane.html           ← Office Add-in task pane HTML + CSS
│   ├── taskpane.js             ← Office Add-in task pane JavaScript
│   ├── commands.html           ← Add-in command stub (required by manifest)
│   ├── guide.html              ← Installation guide (English)
│   └── guide-el.html           ← Installation guide (Greek / Ελληνικά)
│
├── excel-addin/                ← Standalone add-in package (legacy / alternate)
│   ├── manifest.xml
│   └── src/taskpane/
│       ├── taskpane.html
│       └── taskpane.js
│
├── wrangler.toml               ← Cloudflare Worker & Assets configuration
├── wrangler.jsonc              ← (alternate config, not primary)
└── package.json                ← devDependency: wrangler ^3
```

---

## Runtime & Infrastructure

| Aspect | Details |
|--------|---------|
| **Runtime** | Cloudflare Workers (V8 isolate, ES modules) |
| **Compatibility date** | `2024-09-23` |
| **Worker name** | `excel-barcoder` |
| **Static assets** | Served via `[assets]` binding (`ASSETS`) from `./public/` |
| **npm runtime deps** | **None** |
| **Dev dependency** | `wrangler ^3` (CLI for local dev & deploy) |
| **Entry module** | `src/index.js` |

The worker uses ES module syntax (`import`/`export`). All four source files under `src/` are standard ES modules linked together at bundle time by Wrangler.

---

## Worker Entry Point — `src/index.js`

### Routing

The `fetch` handler inspects `url.pathname` and dispatches accordingly:

| Pathname | Handler |
|----------|---------|
| `OPTIONS *` | Returns 204 with CORS preflight headers |
| `/barcode` | Barcode API — returns PNG |
| `/manifest.xml` | Generates and returns the Office Add-in manifest XML |
| `/assets/icon-*` | Returns a 1×1 placeholder PNG icon |
| `/` or `/guide` | 302 redirect → `/guide.html` |
| `/guide-el` | 302 redirect → `/guide-el.html` |
| *(everything else)* | Falls through to `_env.ASSETS.fetch(request)` |

### CORS Headers

All API responses include:
```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, OPTIONS
Access-Control-Allow-Headers: Content-Type
```

### Barcode API Handler

1. Reads `value`, `type`, `width`, `height` from query parameters.
2. Clamps `width` to `[1, 2000]` and `height` to `[1, 1000]`.
3. Calls `selectRenderer(type, value, width, height)` which normalises the `type` string and delegates to the appropriate renderer module.
4. Passes the renderer's `{ width, height, isDark }` to `encodePNG()`.
5. Returns the PNG bytes with `Content-Type: image/png`, `Cache-Control: public, max-age=86400`, and two informational headers `X-Barcode-Width` / `X-Barcode-Height`.

### Manifest Generator

`serveManifest(origin)` builds an Office Add-in XML manifest string at request time with the current worker origin baked into every URL resource. This means the manifest is always correct regardless of the deployment URL — no manual editing required.

The manifest registers:
- `ProviderName`: Excel Barcode
- `DisplayName`: Barcode Inserter
- `Host`: Workbook
- A single command group **"Barcodes"** on `TabHome` with button **"Insert Barcode"** that opens the task pane
- Resources: three icon sizes (16, 32, 80 px), `taskpane.html`, `commands.html`

---

## Barcode Renderers

All renderers share the same interface:

```ts
getRenderer(value: string, reqWidth: number, reqHeight: number): {
  width:  number;   // actual pixel width (enforced minimum applied)
  height: number;   // actual pixel height
  isDark: (x: number, y: number) => boolean;  // pixel colour callback
}
```

`encodePNG` calls `isDark(x, y)` for every pixel. Dark → black (0, 0, 0), light → white (255, 255, 255).

---

### EAN-13 — `src/ean13.js`

#### Encoding

EAN-13 encodes 13 decimal digits. The first digit determines the **parity pattern** for the six left-side digits (choosing between L-code and G-code tables). The right six digits always use R-code.

| Table | Meaning |
|-------|---------|
| L (Left) | 7-bit patterns for left-group digits in normal parity |
| G (G-odd) | 7-bit patterns with inverted parity for left group |
| R (Right) | 7-bit patterns for right-group digits |
| PARITY | 6-char string per first digit (e.g. `'LLGLGG'`) selecting L or G per position |

Structure of the 113-module symbol (including quiet zones):
```
[11 quiet] [101 guard] [6×7 left data] [01010 centre] [6×7 right data] [101 guard] [7 quiet]
```

**Check digit** is computed using ISO/IEC 15420 (GTIN-13): alternating weights 1 and 3 from left, modulo 10.

#### Rasterisation

- Minimum size: 180 × 90 px.
- Bottom 15% of height is the **text area**; guard bars extend 45% of textH into it.
- Digit labels are rendered using a built-in **5×7 bitmap font** (35-bit strings per glyph).
- Label scale factor = `floor(fontSize / 7)`, so labels scale with image height.
- Digit positions:
  - Digit 0: centre of module 5.5 (before left guard)
  - Digits 1–6: centred in their 7-module group starting at module 14
  - Digits 7–12: centred in their 7-module group starting at module 61

---

### CODE128 — `src/code128.js`

#### Encoding

Implements **Subset B** (printable ASCII, codes 32–127) and **Subset C** (pairs of decimal digits, compact encoding for all-digit even-length strings ≥ 4).

Auto-selects Subset C when the input consists entirely of an even number of decimal digits (≥ 4).

**Symbol table**: 106 entries (0–103 data symbols + Start B (104), Start C (105)) each defined as 6-character bar/space width strings (values 1–4), totalling 11 modules per symbol. These are converted to binary at module load time.

**Checksum**: weighted sum modulo 103.

**Stop pattern**: 13-module fixed binary `1100011101011`.

Module array structure:
```
[10 quiet] [start symbol] [data symbols] [checksum] [stop] [10 quiet]
```

#### Label Rendering

- The full input text is rendered below the bars using a built-in **5×7 bitmap font** covering uppercase A–Z, digits 0–9, and common punctuation.
- Truncated to 28 characters with `…` if longer.
- Label scale and horizontal position are computed to fill ~85% of image width.
- Minimum width enforced to `max(reqWidth, 180, N * 2)` (at least 2 px per module).

---

### QR Code — `src/qrcode.js`

#### Encoding Pipeline

1. **Input encoding**: UTF-8 byte mode (mode indicator `0100`).
2. **Version selection**: versions 1–10 are supported; the smallest version whose data capacity (error correction level M) fits the UTF-8-encoded byte count is chosen automatically.
3. **Reed-Solomon ECC**: GF(256) arithmetic over the primitive polynomial `x⁸ + x⁴ + x³ + x² + 1` (0x11D). Generator polynomials are computed on demand for each EC block size.
4. **Data placement**: function patterns (finder, timing, alignment, format, dark module) are placed first; data bits interleaved across EC blocks are placed in the remaining modules using the standard two-column up/down zigzag.
5. **Masking**: all 8 mask patterns are evaluated; the one with the lowest penalty score (ISO/IEC 18004 penalties P1–P4) is applied.
6. **Format information**: error correction level M (bits `01`) combined with the mask pattern index, encoded with BCH(15,5) and XOR-masked with `101010000010010`.

#### Block Specification

A lookup table (versions 1–10, level M) provides `[dataBytes, ecBytes]` per block, matching Annex I of ISO/IEC 18004.

#### Rasterisation

- Output is always square: `side = max(reqWidth, reqHeight)`.
- Module size = `floor(side / qrSize)`, with uniform quiet zone.
- The renderer returns an `isDark(x, y)` function that maps pixel coordinates to module coordinates.

---

## PNG Encoder — `src/png.js`

A self-contained PNG encoder with no library dependencies.

### Algorithm

1. **Scanline construction**: for each row, a filter byte `0x00` (None) is prepended, followed by `width × 3` RGB bytes (each pixel is either `0x000000` black or `0xFFFFFF` white).
2. **DEFLATE (stored)**: the raw scanline buffer is split into 65535-byte non-compressed DEFLATE blocks (`BTYPE=00`). No LZ77 or Huffman coding is applied — the goal is simplicity and correctness, not compression ratio.
3. **Zlib wrapper**: a two-byte zlib header `0x78 0x9C` (deflate, default compression) precedes the DEFLATE blocks; an Adler-32 checksum of the raw data follows.
4. **PNG chunks**:
   - `IHDR`: width, height, bit depth 8, colour type 2 (RGB truecolour)
   - `IDAT`: the zlib-wrapped DEFLATE stream
   - `IEND`: empty trailer
5. **CRC-32**: computed per chunk over `[type bytes | data bytes]` using the standard polynomial `0xEDB88320` (reflected CRC-32/ISO-HDLC).

### Output

`encodePNG(width, height, isDark, scale=1)` returns a `Uint8Array` containing a valid PNG file.

---

## Office Add-in

### Manifest

The manifest is generated dynamically by `serveManifest(origin)` in `src/index.js`. It is a standard Office Add-in 1.1 XML manifest (`TaskPaneApp` type).

Key manifest fields:

| Field | Value |
|-------|-------|
| Id | `A1B2C3D4-E5F6-7890-ABCD-EF1234567890` |
| Version | `1.0.0` |
| ProviderName | Excel Barcode |
| DisplayName | Barcode Inserter |
| Host | Workbook |
| Permissions | ReadWriteDocument |
| SourceLocation | `{origin}/taskpane.html` |
| FunctionFile | `{origin}/commands.html` |

The ribbon button is placed in a custom group **"Barcodes"** on the built-in **Home** tab (`TabHome`) and opens the task pane via `ShowTaskpane` action.

---

### Task Pane UI — `public/taskpane.html`

A self-contained HTML page (inline CSS, no external stylesheet) styled to match Microsoft's Fluent UI design language.

#### Form Controls

| Control | ID | Purpose |
|---------|----|---------|
| Text input | `barcodeValue` | Barcode content |
| Select | `barcodeType` | EAN13 / CODE128 / QR |
| Select | `sizeMode` | resize / custom / fitcell |
| Number input | `barcodeWidth` | Width in px (default 300) |
| Number input | `barcodeHeight` | Height in px (default 150) |
| Checkbox | `useSelectedCell` | Auto-read from selected cell |
| Button | `previewBtn` | Fetch & show preview |
| Button | `insertBtn` | Insert single barcode |
| Button | `bulkBtn` | Insert for all selected rows |
| Div | `preview` | Preview image container |
| Div | `status` | Success / error message |
| Div | `bulkBanner` | Bulk-mode row count banner |

#### CSS Design Tokens (Fluent palette)

| Variable-equivalent | Light value | Dark value (not used — no dark mode in task pane) |
|---------------------|-------------|--------------------------------------------------|
| Primary blue | `#106ebe` | — |
| Primary dark blue | `#005a9e` | — |
| Success green | `#107c10` | — |
| Error red | `#a4262c` | — |
| Border | `#c8c6c4` | — |
| Body text | `#323130` | — |
| Background | `#ffffff` | — |

> The task pane itself runs inside an Office iframe and does not apply a dark mode — Office controls its own chrome separately.

---

### Task Pane Logic — `public/taskpane.js`

Depends on `Office.js` (loaded from Microsoft CDN) and the Excel JavaScript API.

#### Initialisation (`Office.onReady`)

Runs only when `host === Office.HostType.Excel`. Registers:
- `DocumentSelectionChanged` event handler → `syncFromSelection()`
- Initial call to `syncFromSelection()` to pre-populate the value field

#### Selection Sync (`syncFromSelection`)

Uses `Excel.run` to:
1. Load the selected range's `values` and `rowCount`.
2. If `rowCount > 1`, show the bulk banner and the bulk button.
3. If `useSelectedCell` is checked, copy `values[0][0]` to the value input.

#### Barcode URL Builder

```js
function barcodeURL(value, type, w, h) {
  return `/barcode?${new URLSearchParams({ value, type, width: w, height: h })}`;
}
```

`API_BASE` is an empty string, so the URL is relative to the same worker origin that served `taskpane.html`.

#### PNG Fetcher (`fetchPNG`)

1. Fetches the barcode URL.
2. On error, parses the JSON error body.
3. On success, reads the response as `ArrayBuffer`, converts to binary string, and returns `btoa(...)` (Base64).

#### Barcode Placement (`placeBarcode`)

Called with an already-loaded `Range` cell object:

1. **fitcell mode**: uses `cell.width` and `cell.height` as the image dimensions (clamped to minimums).
2. **Other modes**: uses the user-supplied `reqW` / `reqH`.
3. Calls `fetchPNG` → `sheet.shapes.addImage(b64)`.
4. Sets `shape.left`, `shape.top`, `shape.width`, `shape.height`, `shape.name`, `shape.lockAspectRatio = false`.
5. **resize mode**: additionally sets `cell.format.rowHeight` and scales `cell.format.columnWidth` proportionally.

#### Bulk Insert

Iterates each row of the selected range sequentially (not in parallel, because each `placeBarcode` call is async and requires `ctx.sync()`). Empty/false values are skipped. Progress is displayed in the status div.

---

## Installation Guides

Two static HTML pages with identical structure, identical CSS, and full dark-mode support:

| File | Language | URL path |
|------|----------|----------|
| `public/guide.html` | English | `/guide.html` (also `/`, `/guide`) |
| `public/guide-el.html` | Greek | `/guide-el.html` (also `/guide-el`) |

Both pages include:
- A **sticky header** with version badges and a language switcher
- A **sticky sidebar navigation** (desktop) / **horizontal nav** (mobile ≤700px)
- A **YouTube video embed** walkthrough
- Sections: Requirements, Install, First Barcode, Task Pane Guide, Barcode Types, Sizes & Pixels, LAMBDA Formula, REST API, Troubleshooting, FAQ
- A **sticky footer** with links and a PayPal donate button
- Inline JavaScript to dynamically fill in the current server's `location.origin` into code examples

---

## REST API Reference

### Endpoint

```
GET /barcode
```

### Query Parameters

| Parameter | Type | Required | Default | Constraints |
|-----------|------|----------|---------|-------------|
| `value` | string | Yes | — | URL-encoded barcode content |
| `type` | string | No | `EAN13` | `EAN13`, `CODE128`, `QR`, `QRCODE`, `CODE128B`, `C128` (normalised) |
| `width` | integer | No | `300` | Clamped to `[1, 2000]`; minimum enforced per type |
| `height` | integer | No | `150` | Clamped to `[1, 1000]`; minimum enforced per type |

### Type normalisation

Before dispatch, `type` is uppercased and stripped of `-`, `_`, and spaces:
- `EAN13` → EAN-13 renderer
- `CODE128`, `CODE128B`, `C128` → CODE128 renderer
- `QR`, `QRCODE` → QR Code renderer

### Responses

| Status | Content-Type | Body |
|--------|-------------|------|
| 200 | `image/png` | PNG binary |
| 400 | `application/json` | `{ "error": "..." }` |
| 204 | — | CORS preflight (OPTIONS) |

### Response Headers (200)

```
Content-Type: image/png
Cache-Control: public, max-age=86400
Access-Control-Allow-Origin: *
X-Barcode-Width: <actual width>
X-Barcode-Height: <actual height>
```

---

## LAMBDA Formula

For Excel 365 users who prefer a cell formula over the task pane, a custom LAMBDA can be defined in Name Manager:

**Name**: `ADDBARCODE`  
**Refers to**:
```
=LAMBDA(val,type,w,h,IMAGE("https://your-worker.workers.dev/barcode?value="&val&"&type="&type&"&width="&w&"&height="&h))
```

**Signature**: `=ADDBARCODE(val, type, w, h)`

| Argument | Description | Example |
|----------|-------------|---------|
| `val` | Cell reference or literal string | `A1`, `"012345678901"` |
| `type` | Barcode type string | `"EAN13"`, `"CODE128"`, `"QR"` |
| `w` | Width in pixels | `300` |
| `h` | Height in pixels | `150` |

> **Limitation**: Uses `IMAGE()` which stores a URL, not pixel data. The barcode loads from the server on every file open. Not suitable for offline-embedded barcodes.

---

## Dark Mode

Both guide pages implement CSS `@media (prefers-color-scheme: dark)`.

### CSS Variable Overrides

| Variable | Light | Dark |
|----------|-------|------|
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

### Additional Dark Overrides

| Selector | Dark value |
|----------|------------|
| `th` | `background: #1a4f7a` |
| `tr:nth-child(even) td` | `background: #222` |
| `.mockup-input`, `.mockup-select` | `color: var(--text)` |
| `footer` | `background: #111` |
| `.tip` | `background: #0a2a15` |
| `.warn` | `background: #2c1e00; color: var(--text)` |
| `h2` | `color: #5ab0d9` |

> **Note on overrides**: `.tip` used hardcoded `#e6f4ea` (light green) not overridden in dark mode — causing `#d4d4d4` text on a near-white background. `.warn` had the same issue with `#fff4ce`. `h2` used `--blue-dark: #1a6fa8` which gives only 3.2:1 contrast on `#1a1a1a`; the explicit `#5ab0d9` override raises it to 7.2:1 (WCAG AA+).

### Verified WCAG Contrast Ratios (Dark Mode)

| Element | Foreground | Background | Ratio | Level |
|---------|-----------|------------|-------|-------|
| Body text | `#d4d4d4` | `#1a1a1a` | 11.7:1 | AAA |
| Muted / nav text | `#9a9a9a` | `#1a1a1a` | 6.2:1 | AA |
| h2 headings | `#5ab0d9` | `#1a1a1a` | 7.2:1 | AA |
| Links / `--blue` | `#2d9cdb` | `#1a1a1a` | 5.7:1 | AA |
| `.note` text | `#d4d4d4` | `#0d2035` | 11.1:1 | AAA |
| `.tip` text | `#d4d4d4` | `#0a2a15` | 10.5:1 | AAA |
| `.warn` text | `#d4d4d4` | `#2c1e00` | 11.0:1 | AAA |
| Code text | `#d4d4d4` | `#2a2a2a` | 9.7:1 | AAA |
| Table `th` text | `#ffffff` | `#1a4f7a` | 8.6:1 | AAA |
| Even row text | `#d4d4d4` | `#222222` | 10.7:1 | AAA |
| Button text | `#ffffff` | `#1a6fa8` | 5.4:1 | AA |
| Footer links | `#ffffff` | `#111111` | 18.9:1 | AAA |

---

## Configuration Files

### `wrangler.toml`

```toml
name = "excel-barcoder"
main = "src/index.js"
compatibility_date = "2024-09-23"

[assets]
directory = "./public"
binding = "ASSETS"
```

The `ASSETS` binding exposes a `fetch(request)` method that serves files from `./public/` with correct MIME types, ETags, and cache headers automatically.

### `package.json`

```json
{
  "name": "excel-barcode-worker",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "dev": "wrangler dev",
    "deploy": "wrangler deploy"
  },
  "devDependencies": {
    "wrangler": "^3.0.0"
  }
}
```

---

## Data Flow Diagrams

### Single barcode insert (task pane)

```
User clicks "Insert into Excel"
        │
        ▼
taskpane.js: insertBtn handler
        │  reads: value, type, width, height from form
        │
        ▼
Excel.run(ctx)
  getSelectedRange() → load left, top, width, height
        │
        ▼
placeBarcode(ctx, sheet, cell, value, type, w, h)
        │
        ▼
fetchPNG( GET /barcode?value=...&type=...&width=...&height=... )
        │  ← Cloudflare Worker processes request
        │  ← ean13/code128/qrcode renderer builds pixel callback
        │  ← png.js encodes callback to PNG bytes
        │  → returns 200 image/png
        │
        ▼
btoa(PNG bytes) → Base64 string
        │
        ▼
sheet.shapes.addImage(base64)
  shape.left / top / width / height set
  (resize mode: cell.format.rowHeight / columnWidth adjusted)
        │
        ▼
ctx.sync() → barcode embedded in workbook
```

### Barcode API (external caller)

```
Client: GET /barcode?value=ABC&type=CODE128&width=350&height=120
        │
        ▼
Worker fetch handler
  selectRenderer("CODE128", "ABC", 350, 120)
        │
        ▼
code128.js: getRenderer("ABC", 350, 120)
  encode("ABC") → symbol array (Subset B)
  build modules array (quiet + start + data + check + stop + quiet)
  returns { width, height, isDark(x,y) }
        │
        ▼
png.js: encodePNG(width, height, isDark)
  build raw scanlines (H rows × [filter + W×3 bytes])
  DEFLATE stored blocks
  zlib wrap (0x789C header + Adler-32)
  PNG chunks: IHDR + IDAT + IEND
  returns Uint8Array
        │
        ▼
Response(png, { Content-Type: image/png, Cache-Control: public, max-age=86400, ... })
```

---

## Security & Privacy

- **No data persistence**: the worker is stateless. Barcode values passed as query parameters are processed in-memory and discarded. Nothing is logged or stored.
- **CORS open**: `Access-Control-Allow-Origin: *` allows the API to be called from any origin, including Excel Online's sandboxed iframe.
- **No authentication**: the API is public. Do not encode confidential data (passwords, PII) in barcodes intended for public display — the value is visible in the request URL.
- **Cloudflare edge**: requests are handled at the nearest Cloudflare PoP. No backend servers or databases are involved.
- **Office permissions**: the manifest requests `ReadWriteDocument` (required by `sheet.shapes.addImage`). The add-in reads cell values and writes image shapes; it does not access the network on behalf of the user beyond calling the worker's own `/barcode` endpoint.

---

## Known Limitations

| Limitation | Details |
|------------|---------|
| QR version cap | Versions 1–10 only (level M). Maximum ~200 bytes of UTF-8 input. |
| CODE128 charset | ASCII 32–127 only. Accented characters, emoji, and control characters are rejected. |
| EAN-8 | Not supported. Only EAN-13 (and UPC-A as a 12-digit EAN-13 with leading zero). |
| Mobile Excel | `sheet.shapes.addImage()` is not available on Excel for iOS/Android. The task pane shows a preview but cannot insert. |
| PNG compression | The encoder uses DEFLATE stored (uncompressed) blocks. For large barcodes, file sizes are larger than a compressed PNG would be. Performance is not a concern at Cloudflare edge scale. |
| LAMBDA per-workbook | The `ADDBARCODE` LAMBDA must be defined in each workbook's Name Manager. It cannot be shared globally without an Excel template or admin deployment. |
| Bulk insert speed | Rows are processed sequentially (one `ctx.sync()` per row). For large ranges (100+ rows) this may take tens of seconds. |
