// Utility: simple DOM helpers
const $ = (sel) => document.querySelector(sel);
const logEl = $('#log');
const progressBar = $('#progressBar');
const statCounts = $('#statCounts');
const statStatus = $('#statStatus');
const fileInput = $('#fileInput');
const startBtn = $('#startBtn');
const sheetNameInput = $('#sheetName');
const concurrencyInput = $('#concurrency');

let parsedRows = []; // { code: string, links: string[] }
let totalImages = 0;

function log(msg, type = 'info') {
  const line = document.createElement('div');
  line.className = `log-line ${type}`;
  line.textContent = msg;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function setProgress(done, total, status = '') {
  const pct = total ? Math.round((done / total) * 100) : 0;
  progressBar.style.width = `${pct}%`;
  statCounts.textContent = `${done} / ${total}`;
  statStatus.textContent = status || `${pct}%`;
}

function sanitizeCode(str) {
  // Safe for filenames
  return String(str || '')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '_');
}

function splitLinks(raw) {
  if (!raw) return [];
  // Split by comma, semicolon, or newline
  const parts = String(raw).split(/[\n,;]+/);
  return [...new Set(parts.map(s => s.trim()).filter(Boolean))];
}

function detectColumns(headerRow) {
  // return { codeIdx, linksIdx }
  if (!headerRow) return { codeIdx: 0, linksIdx: 1 };
  const headers = headerRow.map(h => String(h || '').toLowerCase().trim());
  const codeIdx =
    headers.findIndex(h => /(code|sku|product)/.test(h)) !== -1
      ? headers.findIndex(h => /(code|sku|product)/.test(h))
      : 0;

  let linksIdx = headers.findIndex(h => /(links|images|image[_ ]?urls?)/.test(h));
  if (linksIdx === -1) {
    // Fallback: assume second column
    linksIdx = 1;
  }
  return { codeIdx, linksIdx };
}

function parseWorkbook(workbook, sheetName = '') {
  const targetSheet = sheetName && workbook.Sheets[sheetName]
    ? sheetName
    : workbook.SheetNames[0];

  const sheet = workbook.Sheets[targetSheet];
  if (!sheet) throw new Error('Sheet not found.');

  // SheetJS: array of arrays for best control
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (!rows.length) throw new Error('Sheet is empty.');

  // Detect header vs data
  const firstRow = rows[0] || [];
  const looksLikeHeader = firstRow.some(cell =>
    typeof cell === 'string' && /code|sku|product|links|images|image/i.test(cell)
  );

  const { codeIdx, linksIdx } = detectColumns(looksLikeHeader ? firstRow : null);

  const startIndex = looksLikeHeader ? 1 : 0;
  const out = [];

  for (let i = startIndex; i < rows.length; i++) {
    const row = rows[i] || [];
    const codeRaw = row[codeIdx];
    const linksRaw = row[linksIdx];

    const code = sanitizeCode(codeRaw);
    const links = splitLinks(linksRaw);

    if (!code || !links.length) continue; // skip incomplete
    out.push({ code, links });
  }

  return out;
}

function inferExtension(url, contentType) {
  // Try from content-type first
  if (contentType) {
    if (contentType.includes('jpeg')) return '.jpg';
    if (contentType.includes('png')) return '.png';
    if (contentType.includes('gif')) return '.gif';
    if (contentType.includes('webp')) return '.webp';
    if (contentType.includes('bmp')) return '.bmp';
    if (contentType.includes('svg')) return '.svg';
  }

  try {
    const u = new URL(url);
    const path = u.pathname.toLowerCase();
    const m = path.match(/\.(jpg|jpeg|png|gif|webp|bmp|svg)(?=$|\?)/i);
    if (m) return `.${m[1].toLowerCase()}`;
  } catch (_) {
    // ignore URL parsing errors
  }
  return ''; // unknown
}

async function fetchAsBlob(url) {
  // CORS must be allowed by the host
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const ct = res.headers.get('content-type') || '';
  const blob = await res.blob();
  return { blob, contentType: ct };
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

async function downloadAll(rows, maxConcurrency = 5) {
  const zip = new JSZip();

  // Build flat job list with target filenames
  const jobs = [];
  for (const { code, links } of rows) {
    links.forEach((url, idx) => {
      const seq = pad2(idx + 1);
      jobs.push({ code, url, seq });
    });
  }

  totalImages = jobs.length;
  setProgress(0, totalImages, 'Preparing…');

  let done = 0;
  const errors = [];

  // Simple worker pool
  let i = 0;
  async function worker() {
    while (i < jobs.length) {
      const job = jobs[i++];
      const { code, url, seq } = job;

      try {
        const { blob, contentType } = await fetchAsBlob(url);
        const ext = inferExtension(url, contentType) || '.bin';
        const fileName = `${code}_${seq}${ext}`;
        zip.file(fileName, blob);
        done++;
        setProgress(done, totalImages, `Downloaded ${fileName}`);
      } catch (err) {
        done++;
        setProgress(done, totalImages, `Failed: ${code}_${seq}`);
        const msg = `❌ ${code} [${seq}] ${url} → ${err.message}`;
        log(msg, 'error');
        errors.push(msg);
      }
    }
  }

  const workers = [];
  const pool = Math.max(1, Math.min(12, Number(maxConcurrency) || 5));
  for (let w = 0; w < pool; w++) workers.push(worker());
  await Promise.all(workers);

  const stamp = new Date();
  const zipName =
    `images_${stamp.getFullYear()}${pad2(stamp.getMonth() + 1)}${pad2(stamp.getDate())}_${pad2(stamp.getHours())}${pad2(stamp.getMinutes())}.zip`;

  setProgress(totalImages, totalImages, 'Zipping…');
  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, zipName);

  if (errors.length) {
    log(`Completed with ${errors.length} errors.`, 'warn');
  } else {
    log('All done ✔', 'success');
  }
  setProgress(totalImages, totalImages, 'Done');
}

fileInput.addEventListener('change', () => {
  logEl.innerHTML = '';
  parsedRows = [];
  totalImages = 0;
  setProgress(0, 0, 'Parsing file…');

  const file = fileInput.files?.[0];
  if (!file) {
    startBtn.disabled = true;
    statStatus.textContent = 'No file selected';
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      // XLSX can also parse CSV transparently
      const wb = XLSX.read(data, { type: 'array' });
      const requestedSheet = (sheetNameInput.value || '').trim();
      parsedRows = parseWorkbook(wb, requestedSheet);
      const totalLinks = parsedRows.reduce((acc, r) => acc + r.links.length, 0);

      if (!parsedRows.length) {
        setProgress(0, 0, 'No valid rows found');
        startBtn.disabled = true;
        log('No valid rows found (need product code and at least one link).', 'warn');
        return;
      }

      startBtn.disabled = false;
      setProgress(0, totalLinks, 'Ready');
      log(`Parsed ${parsedRows.length} products and ${totalLinks} links.`, 'success');
    } catch (err) {
      startBtn.disabled = true;
      setProgress(0, 0, 'Parse error');
      log(`Parse error: ${err.message}`, 'error');
    }
  };
  reader.onerror = () => {
    startBtn.disabled = true;
    setProgress(0, 0, 'Read error');
    log('Could not read file.', 'error');
  };
  reader.readAsArrayBuffer(file);
});

startBtn.addEventListener('click', async () => {
  startBtn.disabled = true;
  sheetNameInput.disabled = true;
  fileInput.disabled = true;

  try {
    await downloadAll(parsedRows, Number(concurrencyInput.value || 5));
  } catch (err) {
    log(`Unexpected error: ${err.message}`, 'error');
  } finally {
    startBtn.disabled = false;
    sheetNameInput.disabled = false;
    fileInput.disabled = false;
  }
});
