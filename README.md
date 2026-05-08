# Excel Barcode Worker

A single Cloudflare Worker that:
1. **Generates barcode PNG images** (EAN-13, CODE128, QR Code) via a REST API
2. **Serves an Excel Office Add-in** task pane that embeds barcodes directly into `.xlsx` workbooks (no external URL stored in the file)

All barcode rendering is pure JavaScript — **zero runtime npm dependencies**.

---

## Architecture

```
https://your-worker.workers.dev/
│
├── GET /barcode?value=...&type=EAN13&width=300&height=150  → PNG image
├── GET /manifest.xml     → Office Add-in manifest (URL auto-filled)
├── GET /taskpane.html    → Add-in task pane UI
├── GET /taskpane.js      → Add-in task pane logic
└── GET /commands.html    → Add-in commands stub
```

---

## Deploy (GitHub → Cloudflare auto-build)

1. Push this repo to GitHub (already done on branch `claude/cloudflare-barcode-worker-c25en`)

2. In Cloudflare Dashboard → **Workers & Pages** → **Create** → **Connect to Git**
   - Select this repository
   - **Build command**: `cd worker && npm install`
   - **Deploy command**: `cd worker && npx wrangler deploy`
   - Root directory: `/`

3. After the first deploy, Cloudflare gives you a URL like:
   `https://excel-barcode-worker.YOUR_SUBDOMAIN.workers.dev`

That's it — no secrets or env vars needed.

---

## API Reference

### `GET /barcode`

| Parameter | Required | Default | Notes |
|-----------|----------|---------|-------|
| `value`   | ✓        | —       | Barcode content |
| `type`    |          | `EAN13` | `EAN13`, `CODE128`, or `QR` |
| `width`   |          | `300`   | Output width in pixels (min 180) |
| `height`  |          | `150`   | Output height in pixels (min 90) |

**Examples:**
```
/barcode?value=012345678901&type=EAN13&width=400&height=200
/barcode?value=HELLO-WORLD&type=CODE128&width=350&height=120
/barcode?value=https://example.com&type=QR&width=300&height=300
```

---

## Install the Excel Add-in

### Option A — Manual sideload (any Excel)

1. Download `https://your-worker.workers.dev/manifest.xml`
2. In Excel: **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Select the downloaded `manifest.xml`
4. A **"Barcodes"** group appears in the **Home** ribbon → click **Insert Barcode**

### Option B — Microsoft 365 Admin Centre (organisation-wide)

Upload the manifest to the M365 admin centre to deploy to all users.

---

## Using the Task Pane

1. Open the **Barcode Inserter** panel (Home → Barcodes → Insert Barcode)
2. Select a cell — its value auto-populates the **Value** field
3. Choose barcode type: **EAN-13**, **CODE128**, or **QR Code**
4. Set width/height in pixels
5. Click **Preview** to verify, then **Insert into Excel**

The barcode image is converted to base64 and embedded in the workbook via
`sheet.shapes.addImage()` — it is saved inside the `.xlsx` file and requires
no internet connection after insertion.

---

## Barcode types

| Type | Value format | Example |
|------|-------------|---------|
| EAN-13 | 12 or 13 digits (check digit auto-calculated) | `012345678901` |
| CODE128 | Any printable ASCII | `ABC-123 / lot#7` |
| QR Code | Any text or URL (≤ ~200 chars for v10) | `https://example.com` |

---

## Local development

```bash
cd worker
npm install
npm run dev          # starts wrangler dev server at http://localhost:8787
```

To test the barcode API locally:
```
http://localhost:8787/barcode?value=012345678901&type=EAN13
http://localhost:8787/barcode?value=HELLO&type=CODE128
http://localhost:8787/barcode?value=https://example.com&type=QR
```

For the add-in taskpane during local dev, use Microsoft's
[office-addin-dev-server](https://github.com/OfficeDev/Office-Addin-Dev-Server)
or simply point Excel at `http://localhost:8787/manifest.xml`.
