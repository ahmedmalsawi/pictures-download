// ===== DOM =====
const $ = (s) => document.querySelector(s);
const logEl = $('#log');
const progressBar = $('#progressBar');
const statCounts = $('#statCounts');
const statStatus = $('#statStatus');

const fileInput = $('#fileInput');
const fileLabel = $('#fileLabel');
const fileNameEl = $('#fileName');

const sheetNameInput = $('#sheetName');
const concurrencyInput = $('#concurrency');
const maxZipFilesInput = $('#maxZipFiles');
const maxZipMBInput = $('#maxZipMB');
const groupByProductInput = $('#groupByProduct');
const startBtn = $('#startBtn');
const estimateBtn = $('#estimateBtn');

// Dashboard KPIs
const kpiTotalProducts = $('#kpi-total-products');
const kpiTotalProductsNote = $('#kpi-total-products-note');
const kpiProductsWithLinks = $('#kpi-products-with-links');
const kpiCoverage = $('#kpi-coverage');
const kpiTotalLinks = $('#kpi-total-links');
const kpiExtTop = $('#kpi-ext-top');
const kpiEstParts = $('#kpi-est-parts');
const kpiZipLimits = $('#kpi-zip-limits');

// Chart
const extChartCanvas = $('#extChart');
let extChart = null;

// Size details
const knownCountEl = $('#known-count');
const knownTotalEl = $('#known-total');
const avgSizeEl = $('#avg-size');
const minMaxEl = $('#min-max');
const estTotalEl = $('#est-total');
const estZipEl = $('#est-zip');
const knownMeter = $('#known-meter');

// ===== State =====
let parsedRows = [];          // [{ code, links:[] }] products that HAVE links
let totalRowsInSheet = 0;     // all products (even without links)
let totalImages = 0;

// ===== Helpers =====
function log(msg, type='info'){
  const d = document.createElement('div');
  d.className = `log-line ${type}`;
  d.textContent = msg;
  logEl.appendChild(d);
  logEl.scrollTop = logEl.scrollHeight;
}
function setProgress(done, total, status=''){
  const pct = total ? Math.round((done/total)*100) : 0;
  progressBar.style.width = `${pct}%`;
  statCounts.textContent = `${done} / ${total}`;
  statStatus.textContent = status || `${pct}%`;
}
function sanitizeCode(s){
  return String(s||'').trim().replace(/[\\/:*?"<>|]/g,'-').replace(/\s+/g,'_');
}
function splitLinks(raw){
  if(!raw) return [];
  return [...new Set(String(raw).split(/[\n,;|،]+/).map(s=>s.trim()).filter(Boolean))];
}

// Column detection (Arabic + English + inference)
function detectColumns(headerRow, rows){
  if (headerRow){
    const h = headerRow.map(x => String(x||'').toLowerCase().trim());
    const codeIdx  = h.findIndex(v => /(code|sku|product|كود|رمز|الصنف|المنتج)/.test(v));
    let linksIdx = h.findIndex(v => /(links|images|image[_ ]?urls?|الرابط|الروابط|الصور|لينكات|لينك)/.test(v));
    return { codeIdx: codeIdx !== -1 ? codeIdx : 0, linksIdx: linksIdx !== -1 ? linksIdx : 1 };
  }
  const limit = Math.min(rows.length, 200);
  const colCount = Math.max(...rows.slice(0,limit).map(r=>r.length), 0);
  const httpScores = new Array(colCount).fill(0);
  const codeScores = new Array(colCount).fill(0);
  const codeSetPerCol = Array.from({length:colCount}, ()=>new Set());

  for (let i=0;i<limit;i++){
    const r = rows[i] || [];
    for(let c=0;c<colCount;c++){
      const v = String(r[c] ?? '').trim();
      if(!v) continue;
      if (/https?:\/\//i.test(v)) httpScores[c]++; else codeScores[c]++;
      codeSetPerCol[c].add(v);
    }
  }
  let linksIdx = httpScores.indexOf(Math.max(...httpScores));
  if (linksIdx < 0) linksIdx = 1;
  let bestCode=-1, bestUnique=-1;
  for(let c=0;c<colCount;c++){
    if (c===linksIdx) continue;
    const unique = codeSetPerCol[c].size + codeScores[c];
    if (unique>bestUnique){bestUnique=unique; bestCode=c;}
  }
  if (bestCode<0) bestCode=0;
  return { codeIdx: bestCode, linksIdx };
}

function parseWorkbook(wb, sheetName=''){
  const target = sheetName && wb.Sheets[sheetName] ? sheetName : wb.SheetNames[0];
  const sheet = wb.Sheets[target];
  if(!sheet) throw new Error('Sheet not found.');
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
  if(!rows.length) throw new Error('Sheet is empty.');

  const first = rows[0] || [];
  const looksHeader = first.some(cell => typeof cell === 'string' &&
    /(code|sku|product|links|images|image|كود|رمز|الصنف|المنتج|الرابط|الروابط|الصور|لينكات|لينك)/i.test(cell)
  );

  const body = looksHeader ? rows.slice(1) : rows;
  totalRowsInSheet = body.length;

  const { codeIdx, linksIdx } = detectColumns(looksHeader ? first : null, body);
  const out = [];
  for(const row of body){
    const code = sanitizeCode(row[codeIdx]);
    const links = splitLinks(row[linksIdx]);
    if(!code) continue;
    if(links.length) out.push({ code, links });
  }
  return out;
}

// Filetype + size helpers
function inferExtension(url, ct){
  if(ct){
    if(ct.includes('jpeg')) return '.jpg';
    if(ct.includes('png'))  return '.png';
    if(ct.includes('gif'))  return '.gif';
    if(ct.includes('webp')) return '.webp';
    if(ct.includes('bmp'))  return '.bmp';
    if(ct.includes('svg'))  return '.svg';
  }
  try{
    const u = new URL(url);
    const m = u.pathname.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp|bmp|svg)(?=$|\?)/i);
    if(m) return `.${m[1].toLowerCase()}`;
  }catch{}
  return '';
}
function summarizeExt(rows){
  const map = new Map();
  for(const {links} of rows){
    for(const url of links){
      const e = (inferExtension(url,'')||'(unknown)').toLowerCase();
      map.set(e, (map.get(e)||0)+1);
    }
  }
  return [...map.entries()].sort((a,b)=>b[1]-a[1]);
}
function prettyMB(bytes){ return (bytes/1024/1024).toFixed(2) + ' MB'; }
function pad2(n){ return String(n).padStart(2,'0'); }

async function headSize(url){
  try{
    const res = await fetch(url, { method:'HEAD', redirect:'follow' });
    if(!res.ok) throw new Error();
    const len = res.headers.get('content-length');
    const ct  = res.headers.get('content-type') || '';
    return { size: len ? Number(len) : null, contentType: ct, ok:true };
  }catch{
    return { size:null, contentType:'', ok:false };
  }
}

// ===== Dashboard renderers =====
function renderKpis(rows){
  const totalLinks = rows.reduce((a,r)=>a+r.links.length,0);
  const dist = summarizeExt(rows);
  const top = dist[0] ? `${dist[0][0]} (${dist[0][1]})` : '—';

  kpiTotalProducts.textContent = totalRowsInSheet.toLocaleString('en');
  kpiTotalProductsNote.textContent = 'يشمل المنتجات بدون روابط';

  kpiProductsWithLinks.textContent = rows.length.toLocaleString('en');
  const coverage = totalRowsInSheet ? Math.round((rows.length/totalRowsInSheet)*100) : 0;
  kpiCoverage.textContent = `تغطية ${coverage}%`;

  kpiTotalLinks.textContent = totalLinks.toLocaleString('en');
  kpiExtTop.textContent = `أكثر امتداد: ${top}`;

  const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
  const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
  kpiZipLimits.textContent = `حدود التقسيم: ${maxFiles} ملف / ${prettyMB(maxBytes)}`;

  // تقدير أجزاء أولي بالعدد فقط
  const partsByCount = Math.ceil(totalLinks / maxFiles);
  kpiEstParts.textContent = totalLinks ? partsByCount.toString() : '—';

  renderExtChart(dist.slice(0,8));
}

function renderExtChart(pairs){
  const labels = pairs.map(p=>p[0]);
  const data = pairs.map(p=>p[1]);

  // Destroy old
  if (extChart){ extChart.destroy(); extChart = null; }

  // NOTE: من غير تلوين مخصص — Chart.js هيختار ألوان افتراضية كويسة
  extChart = new Chart(extChartCanvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12 } },
        tooltip: { rtl: true, textDirection: 'rtl' }
      },
      cutout: '60%'
    }
  });
}

function renderSizeEstimates(info){
  const { totalLinks, knownCount, knownBytes, avg, minB, maxB, estTotal, estZip, knownRatio, partsByCount, partsBySize } = info;

  knownCountEl.textContent = `${knownCount} / ${totalLinks}`;
  knownTotalEl.textContent = knownBytes ? prettyMB(knownBytes) : '—';
  avgSizeEl.textContent = avg ? (avg/1024).toFixed(1)+' KB' : '—';
  minMaxEl.textContent = (minB===Infinity || maxB===0) ? '—' : `${(minB/1024).toFixed(1)} KB / ${(maxB/1024).toFixed(1)} KB`;
  estTotalEl.textContent = estTotal ? prettyMB(estTotal) : '—';
  estZipEl.textContent = estZip ? prettyMB(estZip) : '—';
  knownMeter.style.width = `${Math.round(knownRatio*100)}%`;

  const estParts = Math.max(partsByCount, partsBySize || 1);
  kpiEstParts.textContent = estParts ? estParts.toString() : kpiEstParts.textContent;
}

// ===== Estimate sizes =====
async function estimateSizes(rows){
  const jobs = [];
  for(const r of rows) for(const url of r.links) jobs.push(url);

  const concurrency = Math.max(2, Math.min(10, Number(concurrencyInput.value||6)));
  let i=0, knownBytes=0, knownCount=0, unknown=0, minB=Infinity, maxB=0;

  async function worker(){
    while(i<jobs.length){
      const url = jobs[i++];
      const info = await headSize(url);
      if(info.size!=null){
        knownBytes += info.size;
        knownCount += 1;
        if(info.size < minB) minB = info.size;
        if(info.size > maxB) maxB = info.size;
      } else {
        unknown += 1;
      }
    }
  }
  const workers=[]; for(let w=0; w<concurrency; w++) workers.push(worker());
  await Promise.all(workers);

  const totalLinks = jobs.length;
  const avg = knownCount ? Math.round(knownBytes/knownCount) : 0;
  const estTotal = knownBytes + (unknown * avg);
  const zipFactor = 1.02;
  const estZip = estTotal ? Math.round(estTotal * zipFactor) : 0;

  const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
  const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
  const partsByCount = Math.ceil(totalLinks / maxFiles);
  const partsBySize = estTotal ? Math.ceil(estTotal / maxBytes) : 0;

  renderSizeEstimates({
    totalLinks, knownCount, knownBytes, avg, minB, maxB,
    estTotal, estZip,
    knownRatio: totalLinks ? knownCount/totalLinks : 0,
    partsByCount, partsBySize
  });

  log(knownCount ? 'تم التقدير بنجاح.' : 'ماقدرناش نجيب Content-Length (CORS/إعدادات السيرفر).', knownCount ? 'success':'warn');
}

// ===== Download (batched ZIP) =====
async function fetchAsBlob(url){
  const res = await fetch(url, { redirect:'follow' });
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const ct = res.headers.get('content-type') || '';
  const blob = await res.blob();
  return { blob, contentType: ct };
}

async function downloadAll(rows, maxConcurrency, maxFilesPerZip, maxZipMB, groupByProduct){
  const jobs = [];
  for(const {code, links} of rows){
    links.forEach((url, idx)=>jobs.push({code, url, seq: pad2(idx+1)}));
  }
  totalImages = jobs.length;
  if(!totalImages) return;

  const pool = Math.max(1, Math.min(12, Number(maxConcurrency)||6));
  const maxFiles = Math.max(50, Number(maxZipFiles)||300);
  const maxBytes = Math.max(50, Number(maxZipMB)||200)*1024*1024;

  let done=0; setProgress(0,totalImages,'Starting…');

  let zip = new JSZip(), filesInZip=0, sizeInZip=0, part=1;
  let lock = Promise.resolve();
  const withLock = (fn)=> (lock = lock.then(fn).catch(e=>log(`Lock error: ${e.message}`,'error')));

  async function finalizeZip(){
    if(!filesInZip) return;
    setProgress(done,totalImages,`Zipping part ${part}…`);
    const blob = await zip.generateAsync({type:'blob', streamFiles:true, compression:'DEFLATE', compressionOptions:{level:6}});
    const d = new Date();
    const name = `images_part${String(part).padStart(2,'0')}_${d.getFullYear()}${pad2(d.getMonth()+1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}.zip`;
    saveAs(blob, name);
    log(`Saved ${name} (${filesInZip} files).`,'success');
    zip = new JSZip(); filesInZip=0; sizeInZip=0; part+=1;
  }

  let i=0;
  async function worker(){
    while(i<jobs.length){
      const job = jobs[i++];
      try{
        const { blob, contentType } = await fetchAsBlob(job.url);
        const ext = inferExtension(job.url, contentType) || '.bin';
        const baseName = `${job.code}_${job.seq}${ext}`;
        const filePath = groupByProduct ? `${job.code}/${baseName}` : baseName;

        await withLock(async ()=>{
          if(filesInZip >= maxFiles || (sizeInZip + blob.size) > maxBytes){
            await finalizeZip();
          }
          zip.file(filePath, blob);
          filesInZip += 1; sizeInZip += blob.size;
        });

        done++; setProgress(done,totalImages,`Downloaded ${baseName}`);
      }catch(err){
        done++; setProgress(done,totalImages,`Failed: ${job.code}_${job.seq}`);
        log(`❌ ${job.code} [${job.seq}] ${job.url} → ${err.message}`,'error');
      }
    }
  }

  const workers=[]; for(let w=0; w<pool; w++) workers.push(worker());
  await Promise.all(workers);
  await finalizeZip();
  setProgress(totalImages,totalImages,'Done');
  log('All done ✔','success');
}

// ===== File Upload UI (drag & drop) =====
['dragenter','dragover','dragleave','drop'].forEach(ev=>{
  fileLabel.addEventListener(ev, e=>{ e.preventDefault(); e.stopPropagation(); });
});
['dragenter','dragover'].forEach(ev=>{
  fileLabel.addEventListener(ev, ()=> fileLabel.classList.add('drag'));
});
;['dragleave','drop'].forEach(ev=>{
  fileLabel.addEventListener(ev, ()=> fileLabel.classList.remove('drag'));
});
fileLabel.addEventListener('drop', (e)=>{
  const dt = e.dataTransfer;
  if(!dt || !dt.files || !dt.files[0]) return;
  fileInput.files = dt.files;
  fileInput.dispatchEvent(new Event('change'));
});

fileInput.addEventListener('change', ()=>{
  const f = fileInput.files?.[0];
  fileNameEl.textContent = f ? f.name : 'لم يتم اختيار ملف';
});

// ===== Events =====
fileInput.addEventListener('change', ()=>{
  logEl.innerHTML=''; setProgress(0,0,'جاري قراءة الملف…');
  parsedRows=[]; totalImages=0;

  const f = fileInput.files?.[0];
  if(!f){ startBtn.disabled=true; estimateBtn.disabled=true; statStatus.textContent='لم يتم اختيار ملف'; return; }

  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type:'array' });
      const sheet = (sheetNameInput.value||'').trim();
      parsedRows = parseWorkbook(wb, sheet);

      const totalLinks = parsedRows.reduce((a,r)=>a+r.links.length,0);
      startBtn.disabled = parsedRows.length===0;
      estimateBtn.disabled = parsedRows.length===0;

      // Render dashboard KPIs/chart
      renderKpis(parsedRows);

      setProgress(0,totalLinks,'Ready');
      log(`Parsed ${parsedRows.length} products with links and ${totalLinks} total links.`,'success');
      log('نصيحة: خليك بين 4–8 تحميلات متوازية للاستقرار.','info');
    }catch(err){
      startBtn.disabled=true; estimateBtn.disabled=true;
      setProgress(0,0,'Parse error'); log(`Parse error: ${err.message}`,'error');
    }
  };
  reader.onerror = ()=>{ startBtn.disabled=true; estimateBtn.disabled=true; setProgress(0,0,'Read error'); log('Could not read file.','error'); };
  reader.readAsArrayBuffer(f);
});

estimateBtn.addEventListener('click', async ()=>{ await estimateSizes(parsedRows); });

startBtn.addEventListener('click', async ()=>{
  startBtn.disabled=true; estimateBtn.disabled=true; sheetNameInput.disabled=true; fileInput.disabled=true;
  try{
    await downloadAll(
      parsedRows,
      Number(concurrencyInput.value||6),
      Number(maxZipFilesInput.value||300),
      Number(maxZipMBInput.value||200),
      Boolean(groupByProductInput.checked)
    );
  }catch(err){ log(`Unexpected error: ${err.message}`,'error'); }
  finally{ startBtn.disabled=false; estimateBtn.disabled=false; sheetNameInput.disabled=false; fileInput.disabled=false; }
});
