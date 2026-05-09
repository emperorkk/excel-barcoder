# Excel Barcode Inserter

> *Sheets filled with order,*
> *Konstantinos weaves the code —*
> *Barcodes speak in light.*

A single **Cloudflare Worker** that does two things:

1. **Generates barcode PNG images** (EAN-13, CODE128, QR Code) via a REST API
2. **Serves an Excel Office Add-in** task pane that embeds barcodes directly into `.xlsx` workbooks — no external URL stored in the file

All barcode rendering is pure JavaScript — **zero runtime npm dependencies**.

---

## Architecture

```
https://your-worker.workers.dev/
│
├── GET /barcode?value=...&type=EAN13&width=300&height=150  → PNG image
├── GET /manifest.xml     → Office Add-in manifest (origin auto-filled)
├── GET /taskpane.html    → Add-in task pane UI
├── GET /taskpane.js      → Add-in task pane logic
├── GET /commands.html    → Add-in commands stub
├── GET /guide.html       → Installation guide (English)
├── GET /guide-el.html    → Installation guide (Greek / Ελληνικά)
└── GET /assets/icon-*   → Placeholder add-in icons
```

Static assets (`public/`) are served by Cloudflare's ASSETS binding.
Dynamic routes (barcode API, manifest) are handled in `src/index.js`.

---

## Deploy (GitHub → Cloudflare auto-build)

1. Push this repo to GitHub.

2. In **Cloudflare Dashboard → Workers & Pages → Create → Connect to Git**:
   - Select this repository
   - **Build command**: *(leave empty)*
   - **Deploy command**: `npx wrangler deploy`
   - Root directory: `/`

3. After the first deploy Cloudflare gives you a URL like:
   `https://excel-barcoder.YOUR_SUBDOMAIN.workers.dev`

That's it — no secrets or environment variables needed.

---

## API Reference

### `GET /barcode`

| Parameter | Required | Default  | Notes |
|-----------|----------|----------|-------|
| `value`   | ✓        | —        | Barcode content |
| `type`    |          | `EAN13`  | `EAN13`, `CODE128`, or `QR` |
| `width`   |          | `300`    | Output width in pixels (min 180) |
| `height`  |          | `150`    | Output height in pixels (min 90) |

**Examples:**
```
/barcode?value=012345678901&type=EAN13&width=400&height=200
/barcode?value=HELLO-WORLD&type=CODE128&width=350&height=120
/barcode?value=https://example.com&type=QR&width=300&height=300
```

Returns a `image/png` with `Cache-Control: public, max-age=86400` and full CORS headers.

---

## Barcode Types

| Type | Value format | Max capacity | Use case |
|------|-------------|--------------|----------|
| EAN-13 | 12 or 13 digits | 13 digits | Retail, books (ISBN), supermarket |
| CODE128 | Any printable ASCII (32–127) | ~80 chars | Warehouse, serial numbers, logistics |
| QR Code | Any text / URL | ~200 chars (v10, level M) | URLs, WiFi, contact info, free text |

---

## Install the Excel Add-in

### Option A — Personal sideload (any Excel, no admin needed)

1. Download `https://your-worker.workers.dev/manifest.xml`
2. In Excel: **Home → Add-ins → Manage your apps → Upload My Add-in**
3. Select the downloaded `manifest.xml`
4. A **"Barcodes"** group appears in the **Home** ribbon → click **Insert Barcode**

### Option B — Microsoft 365 Admin Centre (organisation-wide)

Upload the manifest URL (`/manifest.xml`) to the M365 admin centre
(**Settings → Integrated apps → Upload custom apps**) to deploy to all users.

---

## Using the Task Pane

1. Open the **Barcode Inserter** panel (Home → Barcodes → Insert Barcode)
2. Click any cell — its value auto-populates the **Value** field
3. Choose barcode type: **EAN-13**, **CODE128**, or **QR Code**
4. Choose a **Size & cell behaviour** mode:
   - *Custom size — resize cell to fit* (default) — adjusts row/column to match the barcode
   - *Custom size — keep cell as-is* — inserts at specified size, cell unchanged
   - *Fit barcode to current cell size* — generates a barcode sized to the cell
5. Set width/height in pixels (ignored in "fit" mode)
6. Click **Preview** to verify, then **Insert into Excel**

The barcode PNG is fetched, base64-encoded, and embedded via `sheet.shapes.addImage()` —
it is **saved inside the `.xlsx` file** and requires no internet connection to view later.

### Bulk Insert

Select a column range of cells containing barcode values. A blue banner appears with a
green **"Insert for All Selected Cells"** button. Clicking it iterates every row, placing
one barcode per non-empty cell using the chosen size mode.

---

## LAMBDA Formula — `=ADDBARCODE()`

Excel 365 supports a live-linked barcode via a custom LAMBDA + IMAGE() formula.

**One-time setup** in Formulas → Name Manager → New:
- **Name**: `ADDBARCODE`
- **Refers to**:
  ```
  =LAMBDA(val,type,w,h,IMAGE("https://your-worker.workers.dev/barcode?value="&val&"&type="&type&"&width="&w&"&height="&h))
  ```

**Usage**:
```
=ADDBARCODE(A1, "EAN13",  300, 150)
=ADDBARCODE(B2, "CODE128", 400, 120)
=ADDBARCODE(C3, "QR",      250, 250)
```

> **Note:** The LAMBDA formula shows the barcode live from the server — it is *not* embedded in the file. Internet access is required each time the workbook is opened. Use the task pane for offline-embedded barcodes.

---

## Local Development

```bash
npm install
npm run dev          # starts wrangler dev at http://localhost:8787
```

Test the API locally:
```
http://localhost:8787/barcode?value=012345678901&type=EAN13
http://localhost:8787/barcode?value=HELLO&type=CODE128
http://localhost:8787/barcode?value=https://example.com&type=QR
```

Point Excel at `http://localhost:8787/manifest.xml` to test the add-in locally.

---

## Project Structure

```
excel-barcoder/
├── src/
│   ├── index.js       — Cloudflare Worker entry point (routing + manifest)
│   ├── ean13.js       — EAN-13 barcode renderer (pure JS)
│   ├── code128.js     — CODE128 barcode renderer (pure JS)
│   ├── qrcode.js      — QR Code renderer with Reed-Solomon ECC (pure JS)
│   └── png.js         — PNG encoder (no dependencies)
├── public/
│   ├── taskpane.html  — Office Add-in task pane UI
│   ├── taskpane.js    — Task pane logic (Office.js / Excel API)
│   ├── commands.html  — Add-in commands stub
│   ├── guide.html     — Installation guide (English)
│   └── guide-el.html  — Installation guide (Greek)
├── wrangler.toml      — Cloudflare Worker config
└── package.json
```

---

## Links

- [Installation Guide (EN)](https://excel-barcoder.workers.dev/guide.html)
- [Οδηγός Εγκατάστασης (EL)](https://excel-barcoder.workers.dev/guide-el.html)
- [GitHub](https://github.com/emperorkk/excel-barcoder)
- [kourentzes.com/konstantinos](https://kourentzes.com/konstantinos)
- [onecode.gr](https://onecode.gr)
