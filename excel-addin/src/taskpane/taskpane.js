/* globals Office, Excel */

// Replace this with your deployed Cloudflare Worker URL after deployment
const WORKER_URL = 'https://excel-barcode-worker.YOUR_SUBDOMAIN.workers.dev';

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) return;

  const typeSelect   = document.getElementById('barcodeType');
  const valueInput   = document.getElementById('barcodeValue');
  const widthInput   = document.getElementById('barcodeWidth');
  const heightInput  = document.getElementById('barcodeHeight');
  const useCell      = document.getElementById('useSelectedCell');
  const previewDiv   = document.getElementById('preview');
  const previewBtn   = document.getElementById('previewBtn');
  const insertBtn    = document.getElementById('insertBtn');
  const statusDiv    = document.getElementById('status');
  const valueHint    = document.getElementById('valueHint');

  const HINTS = {
    EAN13:   '12 or 13 digits (check digit auto-calculated)',
    CODE128: 'Any printable ASCII text',
    QR:      'Any text or URL (up to ~200 characters)',
  };

  typeSelect.addEventListener('change', () => {
    valueHint.textContent = HINTS[typeSelect.value];
  });

  // When "use selected cell" is ticked, auto-populate value field
  useCell.addEventListener('change', () => {
    if (useCell.checked) syncFromCell();
  });

  async function syncFromCell() {
    try {
      await Excel.run(async ctx => {
        const range = ctx.workbook.getSelectedRange();
        range.load('values');
        await ctx.sync();
        const val = range.values[0][0];
        if (val !== null && val !== undefined && val !== '') {
          valueInput.value = String(val);
        }
      });
    } catch (_) { /* ignore */ }
  }

  // Auto-sync when selection changes
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => { if (useCell.checked) syncFromCell(); }
  );

  syncFromCell(); // initial load

  function barcodeURL(value, type, width, height) {
    const p = new URLSearchParams({ value, type, width, height });
    return `${WORKER_URL}/barcode?${p}`;
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
    let binary = '';
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  }

  previewBtn.addEventListener('click', async () => {
    const value  = valueInput.value.trim();
    if (!value) { showStatus('Please enter a value.', true); return; }
    const type   = typeSelect.value;
    const width  = parseInt(widthInput.value, 10)  || 300;
    const height = parseInt(heightInput.value, 10) || 150;
    const url    = barcodeURL(value, type, width, height);

    previewDiv.innerHTML = '<span>Loading…</span>';
    try {
      const b64 = await fetchPNG(url);
      previewDiv.innerHTML = `<img src="data:image/png;base64,${b64}" alt="barcode"/>`;
      statusDiv.className = '';
    } catch (e) {
      previewDiv.innerHTML = '<span>Preview failed</span>';
      showStatus(e.message, true);
    }
  });

  insertBtn.addEventListener('click', async () => {
    const value  = valueInput.value.trim();
    if (!value) { showStatus('Please enter a value.', true); return; }
    const type   = typeSelect.value;
    const width  = parseInt(widthInput.value, 10)  || 300;
    const height = parseInt(heightInput.value, 10) || 150;
    const url    = barcodeURL(value, type, width, height);

    insertBtn.disabled = true;
    insertBtn.textContent = 'Generating…';

    try {
      const b64 = await fetchPNG(url);

      await Excel.run(async ctx => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        // Get the top-left corner of the current selection for positioning
        const range = ctx.workbook.getSelectedRange();
        range.load(['left', 'top']);
        await ctx.sync();

        const shape = sheet.shapes.addImage(b64);
        shape.left   = range.left;
        shape.top    = range.top;
        shape.width  = width;
        shape.height = height;
        shape.name   = `Barcode_${type}_${value.substring(0, 20)}_${Date.now()}`;

        // Lock aspect ratio and prevent accidental moves
        shape.lockAspectRatio = false;

        await ctx.sync();
      });

      showStatus(`${type} barcode embedded into the workbook.`);
    } catch (e) {
      showStatus(e.message, true);
    } finally {
      insertBtn.disabled = false;
      insertBtn.textContent = 'Insert into Excel';
    }
  });
});
