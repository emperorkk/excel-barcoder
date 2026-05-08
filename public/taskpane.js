/* globals Office, Excel */

const API_BASE = '';

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) return;

  const typeSelect  = document.getElementById('barcodeType');
  const valueInput  = document.getElementById('barcodeValue');
  const widthInput  = document.getElementById('barcodeWidth');
  const heightInput = document.getElementById('barcodeHeight');
  const useCell     = document.getElementById('useSelectedCell');
  const sizeMode    = document.getElementById('sizeMode');
  const sizeInputs  = document.getElementById('sizeInputs');
  const previewDiv  = document.getElementById('preview');
  const previewBtn  = document.getElementById('previewBtn');
  const insertBtn   = document.getElementById('insertBtn');
  const bulkBtn     = document.getElementById('bulkBtn');
  const statusDiv   = document.getElementById('status');
  const valueHint   = document.getElementById('valueHint');
  const bulkBanner  = document.getElementById('bulkBanner');

  const HINTS = {
    EAN13:   '12 or 13 digits (check digit auto-calculated)',
    CODE128: 'Any printable ASCII text',
    QR:      'Any text or URL (up to ~200 characters)',
  };

  typeSelect.addEventListener('change', () => {
    valueHint.textContent = HINTS[typeSelect.value];
  });

  sizeMode.addEventListener('change', () => {
    sizeInputs.classList.toggle('hidden', sizeMode.value === 'fitcell');
  });

  // ── Selection sync ──────────────────────────────────────────────────────
  async function syncFromSelection() {
    try {
      await Excel.run(async ctx => {
        const range = ctx.workbook.getSelectedRange();
        range.load(['values', 'rowCount']);
        await ctx.sync();

        const rows = range.rowCount;
        if (rows > 1) {
          bulkBanner.textContent =
            `${rows} rows selected — use "Insert for All Selected Cells" to generate one barcode per row.`;
          bulkBanner.style.display = 'block';
          bulkBtn.style.display = 'block';
        } else {
          bulkBanner.style.display = 'none';
          bulkBtn.style.display = 'none';
        }

        if (useCell.checked) {
          const val = range.values[0][0];
          if (val !== null && val !== undefined && val !== '') {
            valueInput.value = String(val);
          }
        }
      });
    } catch (_) {}
  }

  useCell.addEventListener('change', () => {
    if (useCell.checked) syncFromSelection();
  });

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => syncFromSelection()
  );

  syncFromSelection();

  // ── Helpers ─────────────────────────────────────────────────────────────
  function barcodeURL(value, type, w, h) {
    return `${API_BASE}/barcode?${new URLSearchParams({ value, type, width: w, height: h })}`;
  }

  function showStatus(msg, isError = false) {
    statusDiv.textContent = msg;
    statusDiv.className = isError ? 'error' : 'ok';
  }

  async function fetchPNG(url) {
    const resp = await fetch(url);
    if (!resp.ok) {
      const json = await resp.json().catch(() => ({}));
      throw new Error(json.error || `HTTP ${resp.status}`);
    }
    const buf = await resp.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  // Insert one barcode into `cell` (Range already loaded with left/top/width/height/format).
  async function placeBarcode(ctx, sheet, cell, value, type, reqW, reqH) {
    const mode = sizeMode.value;

    let imgW, imgH;
    if (mode === 'fitcell') {
      // Use the cell's current point dimensions as image pixel size.
      imgW = Math.max(180, Math.round(cell.width));
      imgH = Math.max(90,  Math.round(cell.height));
    } else {
      imgW = reqW;
      imgH = reqH;
    }

    const b64 = await fetchPNG(barcodeURL(value, type, imgW, imgH));

    const shape = sheet.shapes.addImage(b64);
    shape.left  = cell.left;
    shape.top   = cell.top;
    shape.width  = imgW;
    shape.height = imgH;
    shape.name   = `Barcode_${type}_${value.substring(0, 20)}_${Date.now()}`;
    shape.lockAspectRatio = false;

    if (mode === 'resize') {
      // rowHeight is in points; shape height is also set in points — they match.
      cell.format.rowHeight = imgH;
      // Scale column width proportionally using the current pixel width as reference.
      if (cell.width > 0) {
        cell.format.columnWidth =
          cell.format.columnWidth * (imgW / cell.width);
      }
    }

    await ctx.sync();
  }

  // ── Preview ─────────────────────────────────────────────────────────────
  previewBtn.addEventListener('click', async () => {
    const value = valueInput.value.trim();
    if (!value) { showStatus('Please enter a value.', true); return; }
    const type = typeSelect.value;
    const w = parseInt(widthInput.value, 10) || 300;
    const h = parseInt(heightInput.value, 10) || 150;

    previewDiv.innerHTML = '<span>Loading…</span>';
    try {
      const b64 = await fetchPNG(barcodeURL(value, type, w, h));
      previewDiv.innerHTML = `<img src="data:image/png;base64,${b64}" alt="barcode"/>`;
      statusDiv.className = '';
    } catch (e) {
      previewDiv.innerHTML = '<span>Preview failed</span>';
      showStatus(e.message, true);
    }
  });

  // ── Single insert ────────────────────────────────────────────────────────
  insertBtn.addEventListener('click', async () => {
    const value = valueInput.value.trim();
    if (!value) { showStatus('Please enter a value.', true); return; }
    const type = typeSelect.value;
    const w = parseInt(widthInput.value, 10) || 300;
    const h = parseInt(heightInput.value, 10) || 150;

    insertBtn.disabled = true;
    insertBtn.textContent = 'Generating…';
    try {
      await Excel.run(async ctx => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const cell  = ctx.workbook.getSelectedRange();
        cell.load(['left', 'top', 'width', 'height']);
        cell.format.load('columnWidth');
        await ctx.sync();
        await placeBarcode(ctx, sheet, cell, value, type, w, h);
      });
      showStatus(`${type} barcode inserted.`);
    } catch (e) {
      showStatus(e.message, true);
    } finally {
      insertBtn.disabled = false;
      insertBtn.textContent = 'Insert into Excel';
    }
  });

  // ── Bulk insert ──────────────────────────────────────────────────────────
  bulkBtn.addEventListener('click', async () => {
    const type = typeSelect.value;
    const w = parseInt(widthInput.value, 10) || 300;
    const h = parseInt(heightInput.value, 10) || 150;

    bulkBtn.disabled = true;
    bulkBtn.textContent = 'Processing…';
    let done = 0, skipped = 0;

    try {
      await Excel.run(async ctx => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = ctx.workbook.getSelectedRange();
        range.load(['values', 'rowCount']);
        await ctx.sync();

        const rows = range.rowCount;
        for (let i = 0; i < rows; i++) {
          const rawVal = range.values[i][0];
          const val = (rawVal === null || rawVal === undefined)
            ? '' : String(rawVal).trim();
          if (!val || val === 'false') { skipped++; continue; }

          const cell = range.getCell(i, 0);
          cell.load(['left', 'top', 'width', 'height']);
          cell.format.load('columnWidth');
          await ctx.sync();

          try {
            await placeBarcode(ctx, sheet, cell, val, type, w, h);
            done++;
            showStatus(`Processing… ${done} / ${rows - skipped}`);
          } catch (_) {
            skipped++;
          }
        }
      });
      showStatus(`Done: ${done} barcodes inserted${skipped ? `, ${skipped} skipped` : ''}.`);
    } catch (e) {
      showStatus(e.message, true);
    } finally {
      bulkBtn.disabled = false;
      bulkBtn.textContent = 'Insert for All Selected Cells';
    }
  });
});
