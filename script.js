// ====== DOM ======
const $ = (s)=>document.querySelector(s);
const logEl = $('#log'), toasts = $('#toasts');
const progressBar = $('#progressBar'), statCounts = $('#statCounts'), statStatus = $('#statStatus');
const etaText = $('#etaText');

const fileInput = $('#fileInput'), fileLabel = $('#fileLabel'), fileNameEl = $('#fileName');
const sheetNameInput = $('#sheetName'), concurrencyInput = $('#concurrency'), perHostInput = $('#perHost');
const maxZipFilesInput = $('#maxZipFiles'), maxZipMBInput = $('#maxZipMB'), maxPerProductInput = $('#maxPerProduct');
const namingPatternInput = $('#namingPattern'), whitelistInput = $('#whitelist'), blacklistInput = $('#blacklist');
const groupByProductInput = $('#groupByProduct'), manualSelectInput = $('#manualSelect');
const previewBtn = $('#previewBtn'), estimateBtn = $('#estimateBtn'), startBtn = $('#startBtn');
const pauseBtn = $('#pauseBtn'), resumeBtn = $('#resumeBtn'), cancelBtn = $('#cancelBtn');
const exportReportBtn = $('#exportReportBtn'), exportFailuresBtn = $('#exportFailuresBtn');

// KPIs
const kpiTotalProducts = $('#kpi-total-products'), kpiTotalProductsNote = $('#kpi-total-products-note');
const kpiProductsWithLinks = $('#kpi-products-with-links'), kpiCoverage = $('#kpi-coverage');
const kpiTotalLinks = $('#kpi-total-links'), kpiExtTop = $('#kpi-ext-top'), kpiEstParts = $('#kpi-est-parts'), kpiZipLimits = $('#kpi-zip-limits');

// Charts
const extChartCanvas = $('#extChart'), hostChartCanvas = $('#hostChart'), sizeChartCanvas = $('#sizeChart');
let extChart=null, hostChart=null, sizeChart=null;

// Modal elements
const modal = $('#modal'), modalClose = $('#modalClose'), codeColSel = $('#codeCol'), linksColSel = $('#linksCol'), previewTable = $('#previewTable'), applyColsBtn = $('#applyColsBtn');

// ====== State ======
let parsedRows = [];            // [{code, links:[]}] with links
let totalRowsInSheet = 0;       // all products
let totalImages = 0;
let rawRows = [];               // 2D array for manual selection
let headerRow = null;           // array or null
let codeIdxAuto = 0, linksIdxAuto = 1;

let controller = null;          // AbortController for cancel
let paused = false;
let cancelRequested = false;

const reportRows = [];          // {code,url,status,http_status,size,filename,zip_part,error}
const failures = [];

// ====== Helpers ======
function toast(msg, type='info', timeout=3000){
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  toasts.appendChild(el);
  setTimeout(()=>{ el.classList.add('show'); },10);
  setTimeout(()=>{ el.classList.remove('show'); setTimeout(()=>toasts.removeChild(el),300); }, timeout);
}
function log(msg, type='info'){ const d=document.createElement('div'); d.className=`log-line ${type}`; d.textContent=msg; logEl.appendChild(d); logEl.scrollTop=logEl.scrollHeight; }
function setProgress(done, total, status=''){ const pct= total? Math.round((done/total)*100):0; progressBar.style.width = `${pct}%`; statCounts.textContent = `${done} / ${total}`; statStatus.textContent = status || `${pct}%`; }
function sanitizeCode(s){ return String(s||'').trim().replace(/[\\/:*?"<>|]/g,'-').replace(/\s+/g,'_'); }
function splitLinks(raw){ if(!raw) return []; return [...new Set(String(raw).split(/[\n,;|،]+/).map(s=>s.trim()).filter(Boolean))]; }
function prettyMB(bytes){ return (bytes/1024/1024).toFixed(2)+' MB'; }
function pad2(n){ return String(n).padStart(2,'0'); }
function hostOf(u){ try { return new URL(u).host; } catch { return ''; } }

// Column detection
function detectColumns(headerRow, rows){
  if (headerRow){
    const h = headerRow.map(x=>String(x||'').toLowerCase().trim());
    const codeIdx = h.findIndex(v=>/(code|sku|product|كود|رمز|الصنف|المنتج)/.test(v));
    const linksIdx = h.findIndex(v=>/(links|images|image[_ ]?urls?|الرابط|الروابط|الصور|لينكات|لينك)/.test(v));
    return { codeIdx: codeIdx !== -1 ? codeIdx : 0, linksIdx: linksIdx !== -1 ? linksIdx : 1 };
  }
  const limit = Math.min(rows.length, 200);
  const colCount = Math.max(...rows.slice(0,limit).map(r=>r.length), 0);
  const httpScores = new Array(colCount).fill(0);
  const codeScores = new Array(colCount).fill(0);
  const codeSetPerCol = Array.from({length:colCount}, ()=>new Set());
  for(let i=0;i<limit;i++){
    const r = rows[i] || [];
    for(let c=0;c<colCount;c++){
      const v = String(r[c] ?? '').trim(); if(!v) continue;
      if(/https?:\/\//i.test(v)) httpScores[c]++; else codeScores[c]++;
      codeSetPerCol[c].add(v);
    }
  }
  let linksIdx = httpScores.indexOf(Math.max(...httpScores)); if (linksIdx < 0) linksIdx = 1;
  let bestCode=-1, bestUnique=-1;
  for(let c=0;c<colCount;c++){
    if(c===linksIdx) continue; const unique = codeSetPerCol[c].size + codeScores[c];
    if(unique>bestUnique){bestUnique=unique; bestCode=c;}
  }
  if(bestCode<0) bestCode=0;
  return { codeIdx: bestCode, linksIdx };
}

// Parse workbook with optional manual selection
function parseWorkbook(wb, sheetName=''){
  const target = sheetName && wb.Sheets[sheetName] ? sheetName : wb.SheetNames[0];
  const sheet = wb.Sheets[target]; if(!sheet) throw new Error('Sheet not found.');
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
  if(!rows.length) throw new Error('Sheet is empty.');
  rawRows = rows;
  const first = rows[0] || [];
  const looksHeader = first.some(cell => typeof cell === 'string' &&
    /(code|sku|product|links|images|image|كود|رمز|الصنف|المنتج|الرابط|الروابط|الصور|لينكات|لينك)/i.test(cell));
  headerRow = looksHeader ? first : null;
  const body = looksHeader ? rows.slice(1) : rows;
  totalRowsInSheet = body.length;
  const det = detectColumns(headerRow, body); codeIdxAuto = det.codeIdx; linksIdxAuto = det.linksIdx;
  return buildRowsFrom(body, det.codeIdx, det.linksIdx);
}

function buildRowsFrom(bodyRows, codeIdx, linksIdx){
  const out=[];
  for(const row of bodyRows){
    const code = sanitizeCode(row[codeIdx]); const links = splitLinks(row[linksIdx]);
    if(!code) continue; if(links.length) out.push({ code, links });
  }
  return out;
}

// Charts & KPIs
function inferExt(url, ct){
  if(ct){
    if(ct.includes('jpeg')) return '.jpg'; if(ct.includes('png')) return '.png'; if(ct.includes('gif')) return '.gif';
    if(ct.includes('webp')) return '.webp'; if(ct.includes('bmp')) return '.bmp'; if(ct.includes('svg')) return '.svg';
  }
  try{ const u = new URL(url); const m = u.pathname.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp|bmp|svg)(?=$|\?)/i); if(m) return `.${m[1].toLowerCase()}`; }catch{}
  return '';
}
function summarizeExt(rows){
  const map = new Map();
  for(const {links} of rows){ for(const url of links){ const e = (inferExt(url,'')||'(unknown)').toLowerCase(); map.set(e,(map.get(e)||0)+1); } }
  return [...map.entries()].sort((a,b)=>b[1]-a[1]);
}
function summarizeHosts(rows){
  const map = new Map();
  for(const {links} of rows){ for(const url of links){ const h = hostOf(url)||'(invalid)'; map.set(h,(map.get(h)||0)+1); } }
  return [...map.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
}

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

  const partsByCount = Math.ceil(totalLinks / maxFiles);
  kpiEstParts.textContent = totalLinks ? partsByCount.toString() : '—';

  // Charts
  const labelsExt = dist.slice(0,8).map(p=>p[0]);
  const dataExt = dist.slice(0,8).map(p=>p[1]);
  if(extChart){ extChart.destroy(); } 
  extChart = new Chart(extChartCanvas, { type:'doughnut', data:{ labels:labelsExt, datasets:[{ data:dataExt }] }, options:{ plugins:{ legend:{ position:'bottom', labels:{boxWidth:12} }, tooltip:{rtl:true, textDirection:'rtl'} }, cutout:'60%' } });

  const topHosts = summarizeHosts(rows);
  if(hostChart){ hostChart.destroy(); }
  hostChart = new Chart(hostChartCanvas, { type:'bar', data:{ labels:topHosts.map(p=>p[0]), datasets:[{ data:topHosts.map(p=>p[1]) }] }, options:{ plugins:{ legend:{display:false} }, scales:{ x:{ ticks:{autoSkip:false} }, y:{ beginAtZero:true } } } });
}

// HEAD size estimation
async function headSize(url){
  try{
    const res = await fetch(url, { method:'HEAD', redirect:'follow' });
    if(!res.ok) throw new Error();
    const len = res.headers.get('content-length');
    const ct  = res.headers.get('content-type') || '';
    return { size: len ? Number(len) : null, contentType: ct, ok:true };
  }catch{ return { size:null, contentType:'', ok:false }; }
}

function renderSizeCharts(sizes){
  if(sizeChart){ sizeChart.destroy(); }
  if(!sizes.length){ return; }
  // build histogram (KB buckets)
  const kb = sizes.map(b=>Math.max(1, Math.round(b/1024)));
  kb.sort((a,b)=>a-b);
  const bucketSize = 64; // 64KB per bucket
  const maxKB = kb[kb.length-1];
  const buckets = [];
  for(let i=0;i<=maxKB;i+=bucketSize) buckets.push(i);
  const counts = new Array(buckets.length).fill(0);
  kb.forEach(v=>{
    const idx = Math.min(Math.floor(v/bucketSize), buckets.length-1);
    counts[idx]++;
  });
  const labels = buckets.map(v=>`${v}-${v+bucketSize}KB`);
  sizeChart = new Chart(sizeChartCanvas,{
    type:'bar',
    data:{ labels, datasets:[{ data:counts }] },
    options:{ plugins:{ legend:{display:false} }, scales:{ x:{ ticks:{maxRotation:0,minRotation:0, autoSkip:true} }, y:{ beginAtZero:true } } }
  });
}

async function estimateSizes(rows){
  const jobs=[]; for(const r of rows) for(const u of r.links) jobs.push(u);
  const concurrency = Math.max(2, Math.min(10, Number(concurrencyInput.value||6)));
  let i=0, knownBytes=0, knownCount=0, unknown=0, minB=Infinity, maxB=0;
  const sampleSizes = [];

  async function worker(){
    while(i<jobs.length){
      const url = jobs[i++];
      const info = await headSize(url);
      if(info.size!=null){
        knownBytes += info.size; knownCount += 1;
        sampleSizes.push(info.size);
        if(info.size < minB) minB = info.size;
        if(info.size > maxB) maxB = info.size;
      } else { unknown += 1; }
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

  // Update KPI area
  $('#known-count').textContent = `${knownCount} / ${totalLinks}`;
  $('#known-total').textContent = knownBytes ? prettyMB(knownBytes) : '—';
  $('#avg-size').textContent = avg ? (avg/1024).toFixed(1)+' KB' : '—';
  $('#min-max').textContent = (minB===Infinity || maxB===0) ? '—' : `${(minB/1024).toFixed(1)} KB / ${(maxB/1024).toFixed(1)} KB`;
  $('#est-total').textContent = estTotal ? prettyMB(estTotal) : '—';
  $('#est-zip').textContent   = estZip ? prettyMB(estZip) : '—';
  $('#known-meter').style.width = `${Math.round((totalLinks? knownCount/totalLinks:0)*100)}%`;

  // Histogram
  renderSizeCharts(sampleSizes);
  // Update expected parts KPI
  const estParts = Math.max(partsByCount, partsBySize || 1);
  if(estParts) kpiEstParts.textContent = String(estParts);

  toast(knownCount ? 'تم التقدير' : 'تعذّر جلب Content-Length — تقدير جزئي', knownCount?'success':'warn');
}

// Retry + timeout + jitter
async function fetchWithRetry(url, {tries=3, timeoutMs=15000, backoff=800} = {}){
  for(let attempt=1; attempt<=tries; attempt++){
    const ctrl = new AbortController(); const timer = setTimeout(()=>ctrl.abort(), timeoutMs);
    try{
      const res = await fetch(url, { signal: ctrl.signal, redirect: 'follow' });
      clearTimeout(timer);
      if(!res.ok) throw new Error(`HTTP ${res.status}`);
      return res;
    }catch(e){
      clearTimeout(timer);
      if(attempt === tries) throw e;
      const jitter = Math.floor(Math.random()*400);
      await new Promise(r=>setTimeout(r, backoff*attempt + jitter));
    }
  }
}

// Rate limit per host
const inFlightByHost = new Map();
async function schedulePerHost(url, maxPerHost, taskFn){
  const host = hostOf(url) || 'unknown';
  while((inFlightByHost.get(host)||0) >= maxPerHost){
    await new Promise(r=>setTimeout(r, 100));
  }
  inFlightByHost.set(host, (inFlightByHost.get(host)||0)+1);
  try { return await taskFn(); }
  finally { inFlightByHost.set(host, inFlightByHost.get(host)-1); }
}

// File fetch to blob
async function fetchAsBlob(url){
  const res = await fetchWithRetry(url, { tries:3, timeoutMs:15000, backoff:800 });
  const ct = res.headers.get('content-type') || '';
  const blob = await res.blob();
  return { blob, contentType: ct };
}

// Report export
function exportCSV(rows, fileName){
  const header = ['product_code','url','status','http_status','size_bytes','filename','zip_part','error'];
  const lines = [header.join(',')];
  rows.forEach(r=>{
    const vals = [r.code, r.url, r.status, r.http_status ?? '', r.size ?? '', r.filename ?? '', r.zip_part ?? '', (r.error||'').replace(/[\r\n,]/g,' ')];
    lines.push(vals.map(v=> `"${String(v??'').replace(/"/g,'""')}"`).join(','));
  });
  const csv = lines.join('\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  saveAs(blob, fileName);
}

// Naming
function buildFilePath(pattern, {code, seq, ext}, groupBy){
  let p = pattern;
  p = p.replaceAll('${code}', code).replaceAll('${seq}', seq).replaceAll('${ext}', ext.replace(/^\./,''));
  if(groupBy && !p.includes('/')) p = `${code}/${p}`;
  return p;
}

// ====== Download with batching + controls ======
async function downloadAll(rows, opts){
  const { pool, maxFiles, maxBytes, groupBy, perHost, maxPerProduct, pattern } = opts;

  // Build jobs list with per-product limit
  const jobs=[];
  for(const {code, links} of rows){
    const lim = (maxPerProduct && maxPerProduct>0) ? Math.min(maxPerProduct, links.length) : links.length;
    for(let i=0;i<lim;i++){
      jobs.push({ code, url: links[i], seq: pad2(i+1) });
    }
  }
  // domain filters
  const wl = (whitelistInput.value||'').split(',').map(s=>s.trim()).filter(Boolean);
  const bl = (blacklistInput.value||'').split(',').map(s=>s.trim()).filter(Boolean);
  const passHost = (h)=> (wl.length ? wl.some(x=>h.endsWith(x)) : true) && (bl.length ? !bl.some(x=>h.endsWith(x)) : true);
  const filteredJobs = jobs.filter(j => passHost(hostOf(j.url)));

  totalImages = filteredJobs.length;
  if(!totalImages){ toast('لا توجد روابط مطابقة لمرشحات الدومينات','warn'); return; }

  // Reset report
  reportRows.length = 0; failures.length=0;

  let done=0, filesInZip=0, sizeInZip=0, part=1; setProgress(0,totalImages,'Starting…');
  let zip = new JSZip();

  const startTs = performance.now();
  let lastTs = startTs, lastDone = 0;

  controller = new AbortController(); paused=false; cancelRequested=false;
  pauseBtn.disabled=false; cancelBtn.disabled=false; resumeBtn.disabled=true;
  exportReportBtn.disabled = true; exportFailuresBtn.disabled = true;

  async function finalizeZip(){
    if(!filesInZip) return;
    setProgress(done,totalImages,`Zipping part ${part}…`);
    const blob = await zip.generateAsync({ type:'blob', streamFiles:true, compression:'DEFLATE', compressionOptions:{ level:6 } });
    const d = new Date();
    const name = `images_part${String(part).padStart(2,'0')}_${d.getFullYear()}${pad2(d.getMonth()+1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}.zip`;
    saveAs(blob, name);
    log(`Saved ${name} (${filesInZip} files).`,'success');
    zip = new JSZip(); filesInZip=0; sizeInZip=0; part+=1;
  }

  let i=0;
  async function worker(){
    while(i<filteredJobs.length){
      if(cancelRequested) throw new Error('Cancelled');
      while(paused){ await new Promise(r=>setTimeout(r,150)); if(cancelRequested) throw new Error('Cancelled'); }

      const job = filteredJobs[i++]; const url = job.url;
      try{
        const { blob, contentType } = await schedulePerHost(url, perHost, ()=>fetchAsBlob(url));
        // Content-Type filter: ensure image
        if(!/^image\//i.test(contentType||'')){
          throw new Error(`Not an image (${contentType||'unknown'})`);
        }
        const ext = (inferExt(url, contentType) || '.bin');
        const targetPath = buildFilePath(pattern, {code: job.code, seq: job.seq, ext}, groupBy);

        // rotate zip if thresholds exceed
        if(filesInZip >= maxFiles || (sizeInZip + blob.size) > maxBytes){
          await finalizeZip();
        }
        zip.file(targetPath, blob);
        filesInZip += 1; sizeInZip += blob.size;

        reportRows.push({ code: job.code, url, status:'ok', http_status:200, size:blob.size, filename:targetPath, zip_part:part });
        done++;
        // ETA / speed
        const now = performance.now();
        if(now - lastTs > 1000){
          const elapsedS = (now - startTs)/1000;
          const speed = (done/elapsedS).toFixed(2);
          const remaining = totalImages - done;
          const estRemaining = remaining / Math.max(0.1, (done/elapsedS));
          etaText.textContent = `ETA: ${Math.max(0, Math.round(estRemaining))}ث | السرعة: ${speed} روابط/ث`;
          lastTs = now; lastDone = done;
        }
        setProgress(done,totalImages,`Downloaded ${targetPath}`);
      }catch(err){
        reportRows.push({ code: job.code, url, status:'fail', http_status:'', size:'', filename:'', zip_part:part, error: err.message });
        failures.push({ code: job.code, url, error: err.message });
        done++;
        setProgress(done,totalImages,`Failed: ${job.code}_${job.seq}`);
        log(`❌ ${job.code} [${job.seq}] ${url} → ${err.message}`,'error');
      }
    }
  }

  const workers=[]; for(let w=0; w<pool; w++) workers.push(worker());
  try{
    await Promise.all(workers);
    await finalizeZip();
    setProgress(totalImages,totalImages,'Done');
    toast('اكتمل التحميل','success');
  }catch(e){
    if(e.message === 'Cancelled'){
      toast('تم الإلغاء','warn');
      log('تم إلغاء العملية.','warn');
    } else {
      log(`Unexpected: ${e.message}`,'error');
    }
  }finally{
    pauseBtn.disabled = true; resumeBtn.disabled = true; cancelBtn.disabled = true;
    exportReportBtn.disabled = reportRows.length===0;
    exportFailuresBtn.disabled = failures.length===0;
  }
}

// ====== File Upload UI (drag & drop) ======
['dragenter','dragover','dragleave','drop'].forEach(ev=>{ fileLabel.addEventListener(ev,e=>{ e.preventDefault(); e.stopPropagation(); });});
['dragenter','dragover'].forEach(()=> fileLabel.classList.add);
['dragenter','dragover'].forEach(ev=>{ fileLabel.addEventListener(ev, ()=> fileLabel.classList.add('drag')); });
['dragleave','drop'].forEach(ev=>{ fileLabel.addEventListener(ev, ()=> fileLabel.classList.remove('drag')); });
fileLabel.addEventListener('drop', (e)=>{ const dt=e.dataTransfer; if(!dt||!dt.files||!dt.files[0]) return; fileInput.files=dt.files; fileInput.dispatchEvent(new Event('change')); });
fileInput.addEventListener('change', ()=>{ const f=fileInput.files?.[0]; fileNameEl.textContent = f ? f.name : 'لم يتم اختيار ملف'; });

// ====== Modal (column picker) & preview ======
function openModal(cols){
  codeColSel.innerHTML = ''; linksColSel.innerHTML = '';
  cols.forEach((name, idx)=>{
    const o1=document.createElement('option'); o1.value=idx; o1.textContent=`${idx}: ${name}`; codeColSel.appendChild(o1);
    const o2=document.createElement('option'); o2.value=idx; o2.textContent=`${idx}: ${name}`; linksColSel.appendChild(o2);
  });
  codeColSel.value = codeIdxAuto; linksColSel.value = linksIdxAuto;
  // preview first 50
  const body = headerRow ? rawRows.slice(1, 51) : rawRows.slice(0, 50);
  const colsIndex = [...Array(cols.length).keys()];
  let html = '<table><thead><tr>';
  cols.forEach(c=> html += `<th>${c}</th>`); html += '</tr></thead><tbody>';
  body.forEach(r=>{
    html += '<tr>'; colsIndex.forEach(ci=> html += `<td>${(r[ci]??'')}</td>`); html += '</tr>';
  });
  html += '</tbody></table>';
  previewTable.innerHTML = html;
  modal.classList.remove('hidden');
}
modalClose.addEventListener('click', ()=> modal.classList.add('hidden'));
applyColsBtn.addEventListener('click', ()=>{
  const codeIdx = Number(codeColSel.value), linksIdx = Number(linksColSel.value);
  const body = headerRow ? rawRows.slice(1) : rawRows;
  parsedRows = buildRowsFrom(body, codeIdx, linksIdx);
  renderKpis(parsedRows);
  updateButtons();
  modal.classList.add('hidden');
  toast('تم تطبيق الأعمدة المختارة','success');
});

// ====== Parse file & initial render ======
function updateButtons(){
  const can = parsedRows.length>0;
  startBtn.disabled = !can;
  estimateBtn.disabled = !can;
  previewBtn.disabled = rawRows.length===0;
}
fileInput.addEventListener('change', ()=>{
  logEl.innerHTML=''; setProgress(0,0,'جاري قراءة الملف…');
  parsedRows=[]; totalImages=0; rawRows=[]; headerRow=null;
  const f = fileInput.files?.[0];
  if(!f){ startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; statStatus.textContent='لم يتم اختيار ملف'; return; }
  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type:'array' });
      const sheet = (sheetNameInput.value||'').trim();
      parsedRows = parseWorkbook(wb, sheet);
      renderKpis(parsedRows);
      const totalLinks = parsedRows.reduce((a,r)=>a+r.links.length,0);
      setProgress(0,totalLinks,'Ready');
      log(`Parsed ${parsedRows.length} products with links and ${totalLinks} total links.`,'success');
      log('نصيحة: خليك بين 4–8 تحميلات متوازية للاستقرار.','info');
      updateButtons();
      if(manualSelectInput.checked){
        const cols = (headerRow? headerRow : rawRows[0]).map((v,i)=> (headerRow? String(v||`Col ${i}`): `Col ${i}`));
        openModal(cols);
      }
    }catch(err){
      startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true;
      setProgress(0,0,'Parse error'); log(`Parse error: ${err.message}`,'error');
    }
  };
  reader.onerror = ()=>{ startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; setProgress(0,0,'Read error'); log('Could not read file.','error'); };
  reader.readAsArrayBuffer(f);
});

previewBtn.addEventListener('click', ()=>{
  if(!rawRows.length) return;
  const cols = (headerRow? headerRow : rawRows[0]).map((v,i)=> (headerRow? String(v||`Col ${i}`): `Col ${i}`));
  openModal(cols);
});

estimateBtn.addEventListener('click', async ()=>{ await estimateSizes(parsedRows); });

// ====== Download controls ======
startBtn.addEventListener('click', async ()=>{
  startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; sheetNameInput.disabled=true; fileInput.disabled=true;
  pauseBtn.disabled=false; cancelBtn.disabled=false; resumeBtn.disabled=true;

  const pool = Math.max(1, Math.min(12, Number(concurrencyInput.value||6)));
  const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
  const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
  const perHost = Math.max(1, Math.min(8, Number(perHostInput.value||4)));
  const maxPerProduct = Math.max(0, Number(maxPerProductInput.value||0));
  const groupBy = Boolean(groupByProductInput.checked);
  const pattern = (namingPatternInput.value||'${code}/${code}_${seq}.${ext}').trim();

  try{
    await downloadAll(parsedRows, { pool, maxFiles, maxBytes, groupBy, perHost, maxPerProduct, pattern });
  }catch(err){
    log(`Unexpected error: ${err.message}`,'error');
  }finally{
    startBtn.disabled=false; estimateBtn.disabled=false; previewBtn.disabled=false; sheetNameInput.disabled=false; fileInput.disabled=false;
  }
});
pauseBtn.addEventListener('click', ()=>{ paused=true; pauseBtn.disabled=true; resumeBtn.disabled=false; toast('تم الإيقاف المؤقت','info'); });
resumeBtn.addEventListener('click', ()=>{ paused=false; pauseBtn.disabled=false; resumeBtn.disabled=true; toast('تم الاستكمال','success'); });
cancelBtn.addEventListener('click', ()=>{ cancelRequested=true; pauseBtn.disabled=true; resumeBtn.disabled=true; cancelBtn.disabled=true; toast('جاري الإلغاء…','warn'); });

// Export reports
exportReportBtn.addEventListener('click', ()=> exportCSV(reportRows, 'download_report.csv'));
exportFailuresBtn.addEventListener('click', ()=> exportCSV(failures, 'download_failures.csv'));
