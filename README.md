# Excel Barcode Inserter

> *Sheets filled with order,*
> *Konstantinos weaves the code —*
> *Barcodes speak in light.*

A single **Cloudflare Worker** that:

1. **Generates barcode PNG images** (EAN-13, CODE128, QR Code) via a REST API
2. **Serves a Microsoft Excel Office Add-in** that embeds those barcodes directly inside `.xlsx` workbooks — the image data is stored in the file, no external URL is saved

All barcode rendering is pure JavaScript — **zero runtime npm dependencies**.

---

## Architecture

```
https://your-worker.workers.dev/
│
├── GET /barcode?value=…&type=EAN13&width=300&height=150
│                                              → PNG image (barcode API)
├── GET /manifest.xml                          → Office Add-in manifest (origin auto-filled)
├── GET /taskpane.html                         → Add-in task pane UI
├── GET /taskpane.js                           → Add-in task pane logic
├── GET /commands.html                         → Add-in commands stub
├── GET /assets/icon-{16,32,80}.png            → Placeholder add-in icons
├── GET /guide.html  (also / and /guide)       → Installation guide (English)
└── GET /guide-el.html  (also /guide-el)       → Installation guide (Greek / Ελληνικά)
```

Dynamic routes (`/barcode`, `/manifest.xml`, `/assets/icon-*`, redirects) are handled in
`src/index.js`. Everything else falls through to Cloudflare's **ASSETS** binding which
serves the `public/` directory.

---

## Project Structure

```
excel-barcoder/
│
├── src/
│   ├── index.js          Worker entry point — routing, manifest generation, CORS
│   ├── ean13.js          EAN-13 encoder + pixel renderer (pure JS)
│   ├── code128.js        CODE128 encoder + pixel renderer (pure JS)
│   ├── qrcode.js         QR Code encoder with Reed-Solomon ECC (pure JS)
│   └── png.js            PNG encoder — no external library
│
├── public/               Static assets served via Cloudflare ASSETS binding
│   ├── taskpane.html     Office Add-in task pane — HTML + inline CSS
│   ├── taskpane.js       Task pane logic (Office.js / Excel JavaScript API)
│   ├── commands.html     Add-in commands stub (required by manifest spec)
│   ├── guide.html        Installation & usage guide (English)
│   └── guide-el.html     Installation & usage guide (Greek)
│
├── excel-addin/          Standalone add-in package for manual / local deployment
│   ├── manifest.xml      Static manifest — edit WORKER_URL before uploading
│   └── src/taskpane/
│       ├── taskpane.html Simplified task pane (no bulk insert, no size modes)
│       └── taskpane.js   Task pane logic — uses hardcoded WORKER_URL constant
│
├── wrangler.toml         Cloudflare Worker + Assets configuration (primary)
└── package.json          devDependency: wrangler ^3
```

---

## Deploy to Cloudflare

### Option A — GitHub → Cloudflare auto-deploy (recommended)

1. Push this repo to GitHub.
2. **Cloudflare Dashboard → Workers & Pages → Create → Connect to Git**
   - Select this repository
   - Build command: *(leave empty)*
   - Deploy command: `npx wrangler deploy`
   - Root directory: `/`
3. Cloudflare provides a URL such as `https://excel-barcoder.YOUR_SUBDOMAIN.workers.dev`.

No secrets or environment variables are required.

### Option B — CLI deploy

```bash
npm install
npm run deploy   # runs: wrangler deploy
```

---

## Barcode API

### `GET /barcode`

| Parameter | Required | Default | Constraints |
|-----------|----------|---------|-------------|
| `value`   | yes      | —       | URL-encoded barcode content |
| `type`    | no       | `EAN13` | `EAN13`, `CODE128`, `QR` (also `C128`, `QRCODE`, `CODE128B`) |
| `width`   | no       | `300`   | Integer; clamped to 1–2000; type minimum enforced |
| `height`  | no       | `150`   | Integer; clamped to 1–1000; type minimum enforced |

**200 response:** `image/png` with `Cache-Control: public, max-age=86400` and CORS headers.  
**400 response:** `application/json` — `{ "error": "…" }`.

```
/barcode?value=012345678901&type=EAN13&width=400&height=200
/barcode?value=HELLO-WORLD&type=CODE128&width=350&height=120
/barcode?value=https://example.com&type=QR&width=300&height=300
```

---

## Barcode Types

| Type | Input format | Max capacity | Typical use |
|------|-------------|--------------|-------------|
| **EAN-13** | 12 or 13 digits | 13 digits | Retail labels, books (ISBN), supermarket items |
| **CODE128** | Printable ASCII (codes 32–127) | ~80 readable chars | Warehouse, logistics, serial numbers |
| **QR Code** | Any Unicode text | ~200 UTF-8 bytes (v10, level M) | URLs, WiFi, contact info, free-form text |

**EAN-13:** supply 12 digits → check digit calculated automatically; 13 digits → check digit validated.  
**CODE128:** auto-selects Subset C (compact) for even-length all-digit strings ≥ 4; otherwise Subset B.  
**QR:** version 1–10 auto-selected by input length; error correction level M (~15 % recovery).

---

## Install the Excel Add-in

### Option A — Personal sideload (no admin rights needed)

1. Download `https://your-worker.workers.dev/manifest.xml`
   *(the origin URL is baked in automatically — no editing required)*
2. Open Excel → **Home → Add-ins → Manage your apps → Upload My Add-in**
3. Select the downloaded `manifest.xml` → confirm
4. A **"Barcodes"** group appears in the **Home** ribbon → click **Insert Barcode**

### Option B — Organisation-wide via Microsoft 365 Admin Centre

Sign in to `admin.microsoft.com` → **Settings → Integrated apps → Upload custom apps**
→ provide the manifest URL `https://your-worker.workers.dev/manifest.xml`.
Users see the add-in within 24 hours (usually faster).

### Option C — Standalone manifest (`excel-addin/manifest.xml`)

Use this if you need a static XML file independent of the auto-generation route.
Edit every `WORKER_URL` placeholder in `excel-addin/manifest.xml` with your deployed
worker URL, then sideload or upload that file.

---

## Using the Task Pane

1. Open **Home → Barcodes → Insert Barcode** to open the panel.
2. Click any cell — its value auto-fills the **Value** field.
3. Choose **Type**: EAN-13, CODE128, or QR Code.
4. Choose **Size & cell behaviour**:
   - *Custom size — resize cell to fit* (default): inserts at your pixel dimensions and auto-adjusts the row height and column width to match.
   - *Custom size — keep cell as-is*: inserts at your pixel dimensions without touching the cell.
   - *Fit barcode to current cell size*: reads the cell's current dimensions and generates a barcode that fills them exactly.
5. Set **Width** and **Height** in pixels (ignored in "fit" mode; minimums: 180 × 90 px).
6. Click **Preview** to verify appearance, then **Insert into Excel**.

The PNG is fetched once, Base64-encoded, and embedded via `sheet.shapes.addImage()`.
The barcode is **saved inside the `.xlsx` file** — no internet needed to view it later.

### Bulk Insert

Select a range of cells in a single column. The task pane shows a blue banner with the
row count and a green **"Insert for All Selected Cells"** button. Barcodes are placed
row-by-row; empty cells are skipped. The chosen size mode applies to every row.

---

## LAMBDA Formula — `=ADDBARCODE()`

For live-linked barcodes (Excel 365 only, requires `IMAGE()` function):

**One-time setup — Formulas → Name Manager → New:**

| Field | Value |
|-------|-------|
| Name | `ADDBARCODE` |
| Scope | Workbook |
| Refers to | `=LAMBDA(val,type,w,h,IMAGE("https://your-worker.workers.dev/barcode?value="&val&"&type="&type&"&width="&w&"&height="&h))` |

**Usage:**
```excel
=ADDBARCODE(A1, "EAN13",   300, 150)
=ADDBARCODE(B2, "CODE128", 400, 120)
=ADDBARCODE(C3, "QR",      250, 250)
```

> **Important:** The LAMBDA stores a URL in the cell, not pixel data. The barcode
> reloads from the server each time the workbook opens — internet access is required.
> For offline-embedded barcodes use the task pane **Insert** button instead.
> The LAMBDA definition is per-workbook; repeat the Name Manager setup in each new file
> or save a template (`.xltx`) with the LAMBDA pre-defined.

---

## Local Development

```bash
npm install
npm run dev          # wrangler dev — serves at http://localhost:8787
```

Test the API:
```
http://localhost:8787/barcode?value=012345678901&type=EAN13
http://localhost:8787/barcode?value=HELLO&type=CODE128
http://localhost:8787/barcode?value=https://example.com&type=QR
```

Point Excel at `http://localhost:8787/manifest.xml` to test the add-in locally.

---

## Links

- [Installation Guide (English)](https://excel-barcoder.workers.dev/guide.html)
- [Οδηγός Εγκατάστασης (Ελληνικά)](https://excel-barcoder.workers.dev/guide-el.html)
- [GitHub](https://github.com/emperorkk/excel-barcoder)
- [kourentzes.com/konstantinos](https://kourentzes.com/konstantinos)
- [onecode.gr](https://onecode.gr)
