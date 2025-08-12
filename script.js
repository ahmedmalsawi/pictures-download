// ===== helpers =====
const $ = (s)=>document.querySelector(s);
const byId = (id)=>document.getElementById(id);
const fmtMB = (b)=> (b/1024/1024).toFixed(2)+' MB';
const pad2 = (n)=> String(n).padStart(2,'0');
const hostOf = (u)=>{ try{return new URL(u).host;}catch{return ''} };
const sanitizeCode = (s)=> String(s||'').trim().replace(/[\\/:*?"<>|]/g,'-').replace(/\s+/g,'_');
const splitLinks = (raw)=> raw ? [...new Set(String(raw).split(/[\n,;|،]+/).map(s=>s.trim()).filter(Boolean))] : [];
const sleep = (ms)=> new Promise(r=>setTimeout(r, ms));

// ===== DOM =====
const helpBtn = byId('helpBtn'), helpModal = byId('helpModal'), helpClose = byId('helpClose');
const advancedModeChk = byId('advancedMode');

const logEl = byId('log'), toasts = byId('toasts');
const fileInput = byId('fileInput'), fileLabel = byId('fileLabel'), fileNameEl = byId('fileName'), fileMetaEl = byId('fileMeta');
const sheetNameInput = byId('sheetName'), concurrencyInput = byId('concurrency'), perHostInput = byId('perHost');
const maxZipFilesInput = byId('maxZipFiles'), maxZipMBInput = byId('maxZipMB'), maxPerProductInput = byId('maxPerProduct');
const namingPatternInput = byId('namingPattern'), whitelistInput = byId('whitelist'), blacklistInput = byId('blacklist');
const groupByProductInput = byId('groupByProduct'), manualSelectInput = byId('manualSelect');
const previewBtn = byId('previewBtn'), previewBtnSide = byId('previewBtn_side'), estimateBtn = byId('estimateBtn'), startBtn = byId('startBtn');

const taskDownload = byId('taskDownload'), taskDrivePics = byId('taskDrivePics'), taskDriveZips = byId('taskDriveZips');
const driveCard = byId('driveCard');

const remoteEndpointInput = byId('remoteEndpoint');
const remoteFolderLinkInput = byId('remoteFolderLink'), remoteSessionInput = byId('remoteSession');
const remoteConcurrencyInput = byId('remoteConcurrency');
const uploadZipsBtn = byId('uploadZipsBtn'), uploadPicsBtn = byId('uploadPicsBtn'), testUploadBtn = byId('testUploadBtn');

const exportDrivePicsXlsxBtn = byId('exportDrivePicsXlsx'), xlOnePerCodeChk = byId('xlOnePerCode');

const progressBar = byId('progressBar'), statCounts = byId('statCounts'), statStatus = byId('statStatus'), etaText = byId('etaText');
const pauseBtn = byId('pauseBtn'), resumeBtn = byId('resumeBtn'), cancelBtn = byId('cancelBtn');
const exportReportBtn = byId('exportReportBtn'), exportFailuresBtn = byId('exportFailuresBtn');

const kpiTotalProducts = byId('kpi-total-products'), kpiTotalProductsNote = byId('kpi-total-products-note');
const kpiProductsWithLinks = byId('kpi-products-with-links'), kpiCoverage = byId('kpi-coverage');
const kpiTotalLinks = byId('kpi-total-links'), kpiExtTop = byId('kpi-ext-top'), kpiEstParts = byId('kpi-est-parts'), kpiZipLimits = byId('kpi-zip-limits');

let extChart=null, hostChart=null, sizeChart=null;
const extChartCanvas = byId('extChart'), hostChartCanvas = byId('hostChart'), sizeChartCanvas = byId('sizeChart');

// modal + preview
const modal = byId('modal'), modalClose = byId('modalClose');
const codeColSel = byId('codeCol'), linksColSel = byId('linksCol');
const previewRows = byId('previewRows'), previewPics = byId('previewPics');
const tabRows = byId('tabRows'), tabPics = byId('tabPics'), applyColsBtn = byId('applyColsBtn');

// ===== state =====
let parsedRows = [];
let totalRowsInSheet = 0;
let rawRows = [];
let headerRow = null;
let codeIdxAuto = 0, linksIdxAuto = 1;

const reportRows = [];
const failures = [];
const zipParts = [];            // {name, blob}
const uploadedPictureLinks = []; // {code, link, filename}
const uploadedZipLinks = [];     // {name, link}

// ===== UI =====
function toast(msg, type='info', timeout=3000){
  const el = document.createElement('div'); el.className = `toast ${type}`; el.textContent = msg;
  toasts.appendChild(el); setTimeout(()=> el.classList.add('show'), 10);
  setTimeout(()=>{ el.classList.remove('show'); setTimeout(()=> toasts.removeChild(el), 300); }, timeout);
}
function log(msg, type='info', asHtml=false){
  const d=document.createElement('div'); d.className=`log-line ${type}`;
  if(asHtml) d.innerHTML = msg; else d.textContent = msg;
  logEl.appendChild(d); logEl.scrollTop=logEl.scrollHeight;
}
function setProgress(done, total, status=''){
  const pct = total? Math.round((done/total)*100) : 0;
  progressBar.style.width = `${pct}%`; statCounts.textContent = `${done} / ${total}`; statStatus.textContent = status || `${pct}%`;
}
function showGuideToasts(){
  const steps = [
    'تم قراءة الملف. لو الاعمدة غلط افتح المعاينة وحددها.',
    'تقدر تعمل حساب المساحة قبل البدء.',
    'اختر المهام: تحميل أو رفع إلى درايف.',
    'اضغط ابدأ. بعد الانتهاء تقدر تصدر Excel بروابط درايف.'
  ];
  let delay=200; steps.forEach((t,i)=> setTimeout(()=> toast(`${i+1}/${steps.length} - ${t}`,'info',4000), delay+=600));
}
function toggleDriveCard(){ driveCard.classList.toggle('hidden', !(taskDrivePics.checked || taskDriveZips.checked)); }
[taskDrivePics, taskDriveZips].forEach(el=> el.addEventListener('change', toggleDriveCard));
toggleDriveCard();

// simple/advanced
const settingsCard = byId('settingsCard');
const simpleBlock = byId('simpleBlock');
const advancedBlock = byId('advancedBlock');
advancedModeChk.addEventListener('change', ()=> {
  const adv = advancedModeChk.checked;
  advancedBlock.classList.toggle('hidden', !adv);
});
advancedBlock.classList.add('hidden');

// ===== parsing =====
function detectColumns(headerRow, rows){
  if(headerRow){
    const h = headerRow.map(x=>String(x||'').toLowerCase().trim());
    const codeIdx = h.findIndex(v=>/(code|sku|product|كود|رمز|الصنف|المنتج)/.test(v));
    const linksIdx = h.findIndex(v=>/(links|images|image[_ ]?urls?|الرابط|الروابط|الصور|لينكات|لينك)/.test(v));
    return { codeIdx: codeIdx!==-1? codeIdx:0, linksIdx: linksIdx!==-1? linksIdx:1 };
  }
  const limit = Math.min(rows.length, 200);
  const colCount = Math.max(...rows.slice(0,limit).map(r=>r.length), 0);
  const httpScores = new Array(colCount).fill(0);
  const codeScores = new Array(colCount).fill(0);
  const codeSetPerCol = Array.from({length:colCount}, ()=>new Set());
  for(let i=0;i<limit;i++){
    const r=rows[i]||[];
    for(let c=0;c<colCount;c++){
      const v=String(r[c]??'').trim(); if(!v) continue;
      if(/https?:\/\//i.test(v)) httpScores[c]++; else codeScores[c]++;
      codeSetPerCol[c].add(v);
    }
  }
  let linksIdx = httpScores.indexOf(Math.max(...httpScores)); if(linksIdx<0) linksIdx=1;
  let bestCode=-1, bestUnique=-1;
  for(let c=0;c<colCount;c++){
    if(c===linksIdx) continue;
    const unique = codeSetPerCol[c].size + codeScores[c];
    if(unique>bestUnique){ bestUnique=unique; bestCode=c; }
  }
  if(bestCode<0) bestCode=0;
  return { codeIdx: bestCode, linksIdx };
}
function buildRowsFrom(bodyRows, codeIdx, linksIdx){
  const out=[];
  for(const row of bodyRows){
    const code = sanitizeCode(row[codeIdx]);
    const links = splitLinks(row[linksIdx]);
    if(!code) continue;
    if(links.length) out.push({ code, links });
  }
  return out;
}
function parseWorkbook(wb, sheetName=''){
  const target = sheetName && wb.Sheets[sheetName] ? sheetName : wb.SheetNames[0];
  const sheet = wb.Sheets[target]; if(!sheet) throw new Error('Sheet not found.');
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
  if(!rows.length) throw new Error('Sheet is empty.');
  rawRows = rows;
  const first = rows[0]||[];
  const looksHeader = first.some(cell => typeof cell==='string' && /(code|sku|product|links|images|image|كود|رمز|الصنف|المنتج|الرابط|الروابط|الصور|لينكات|لينك)/i.test(cell));
  headerRow = looksHeader ? first : null;
  const body = looksHeader ? rows.slice(1) : rows;
  totalRowsInSheet = body.length;
  const det = detectColumns(headerRow, body); codeIdxAuto=det.codeIdx; linksIdxAuto=det.linksIdx;
  return buildRowsFrom(body, det.codeIdx, det.linksIdx);
}

// ===== charts & KPIs =====
function inferExt(url, ct){
  if(ct){
    if(ct.includes('jpeg')) return '.jpg';
    if(ct.includes('png'))  return '.png';
    if(ct.includes('gif'))  return '.gif';
    if(ct.includes('webp')) return '.webp';
    if(ct.includes('bmp'))  return '.bmp';
    if(ct.includes('svg'))  return '.svg';
  }
  try{ const u=new URL(url); const m=u.pathname.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp|bmp|svg)(?=$|\?)/i); if(m) return `.${m[1].toLowerCase()}`; }catch{}
  return '(unknown)';
}
function summarizeExt(rows){
  const m=new Map();
  for(const {links} of rows){ for(const url of links){ const e=(inferExt(url,'')||'(unknown)').toLowerCase(); m.set(e,(m.get(e)||0)+1); } }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]);
}
function summarizeHosts(rows){
  const m=new Map();
  for(const {links} of rows){ for(const url of links){ const h=hostOf(url)||'(invalid)'; m.set(h,(m.get(h)||0)+1); } }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
}
function renderKpis(rows){
  const totalLinks = rows.reduce((a,r)=>a+r.links.length,0);
  const dist = summarizeExt(rows);
  const top = dist[0] ? `${dist[0][0]} (${dist[0][1]})` : '—';

  kpiTotalProducts.textContent = totalRowsInSheet.toLocaleString('en');
  kpiTotalProductsNote.textContent = 'يشمل المنتجات بدون روابط';
  kpiProductsWithLinks.textContent = rows.length.toLocaleString('en');
  kpiCoverage.textContent = `تغطية ${totalRowsInSheet? Math.round((rows.length/totalRowsInSheet)*100):0}%`;
  kpiTotalLinks.textContent = totalLinks.toLocaleString('en');
  kpiExtTop.textContent = `اكثر امتداد: ${top}`;

  const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
  const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
  kpiZipLimits.textContent = `حد التقسيم: ${maxFiles} ملف / ${fmtMB(maxBytes)}`;
  kpiEstParts.textContent = totalLinks ? String(Math.ceil(totalLinks / maxFiles)) : '—';

  const extLabels = dist.slice(0,8).map(p=>p[0]);
  const extData   = dist.slice(0,8).map(p=>p[1]);
  if(extChart) extChart.destroy();
  extChart = new Chart(extChartCanvas,{type:'doughnut', data:{labels:extLabels, datasets:[{data:extData}]}, options:{plugins:{legend:{position:'bottom', labels:{boxWidth:12}}, tooltip:{rtl:true, textDirection:'rtl'}}, cutout:'60%'}});
  const hosts = summarizeHosts(rows);
  if(hostChart) hostChart.destroy();
  hostChart = new Chart(hostChartCanvas,{type:'bar', data:{labels:hosts.map(p=>p[0]), datasets:[{data:hosts.map(p=>p[1])}]}, options:{plugins:{legend:{display:false}}, scales:{x:{ticks:{autoSkip:false,maxRotation:0,minRotation:0}}, y:{beginAtZero:true}}}});
}

// ===== estimation =====
async function headSize(url){
  try{
    const res = await fetch(url, { method:'HEAD', redirect:'follow' });
    if(!res.ok) throw 0;
    const len = res.headers.get('content-length'); const ct = res.headers.get('content-type')||'';
    return { size: len? Number(len):null, contentType: ct, ok:true };
  }catch{ return { size:null, contentType:'', ok:false }; }
}
function renderSizeHistogram(sizes){
  if(sizeChart) sizeChart.destroy();
  if(!sizes.length) return;
  const kb = sizes.map(b=>Math.max(1, Math.round(b/1024))).sort((a,b)=>a-b);
  const step=64; const maxKB = kb[kb.length-1]; const buckets=[]; for(let i=0;i<=maxKB;i+=step) buckets.push(i);
  const counts=new Array(buckets.length).fill(0);
  kb.forEach(v=> counts[Math.min(Math.floor(v/step), counts.length-1)]++);
  const labels = buckets.map(v=>`${v}-${v+step}KB`);
  sizeChart = new Chart(sizeChartCanvas, { type:'bar', data:{labels, datasets:[{data:counts}]}, options:{plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}}} });
}
async function estimateSizes(rows){
  try{
    const jobs=[]; for(const r of rows) for(const u of r.links) jobs.push(u);
    const conc = Math.max(2, Math.min(10, Number(concurrencyInput.value||6)));
    let i=0, knownBytes=0, knownCount=0, unknown=0, minB=Infinity, maxB=0; const sample=[];
    async function worker(){
      while(i<jobs.length){
        const info = await headSize(jobs[i++]);
        if(info.size!=null){ knownBytes += info.size; knownCount++; sample.push(info.size); minB=Math.min(minB,info.size); maxB=Math.max(maxB,info.size); }
        else unknown++;
      }
    }
    await Promise.all(Array.from({length:conc}, worker));
    const total=jobs.length, avg = knownCount? Math.round(knownBytes/knownCount):0;
    const estTotal = knownBytes + (unknown*avg), zipFactor=1.02, estZip = estTotal? Math.round(estTotal*zipFactor):0;
    const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
    const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
    const partsByCount = Math.ceil(total / maxFiles);
    const partsBySize  = estTotal? Math.ceil(estTotal / maxBytes) : 0;
    byId('known-count').textContent = `${knownCount} / ${total}`;
    byId('known-total').textContent = knownBytes? fmtMB(knownBytes):'—';
    byId('avg-size').textContent = avg? (avg/1024).toFixed(1)+' KB':'—';
    byId('min-max').textContent = (minB===Infinity||maxB===0)? '—' : `${(minB/1024).toFixed(1)} KB / ${(maxB/1024).toFixed(1)} KB`;
    byId('est-total').textContent = estTotal? fmtMB(estTotal):'—';
    byId('est-zip').textContent = estZip? fmtMB(estZip):'—';
    byId('known-meter').style.width = `${Math.round((total? knownCount/total:0)*100)}%`;
    renderSizeHistogram(sample);
    const estParts = Math.max(partsByCount, partsBySize||1);
    if(estParts) kpiEstParts.textContent = String(estParts);
    toast(knownCount? 'تم حساب المساحة':'تعذر جلب Content-Length. حساب جزئي', knownCount?'success':'warn');
  }catch(err){ toast(`خطأ في حساب المساحة: ${err.message}`,'error'); }
}

// ===== preview =====
function openModal(cols){
  codeColSel.innerHTML = ''; linksColSel.innerHTML = '';
  cols.forEach((name, idx)=>{
    const o1=document.createElement('option'); o1.value=idx; o1.textContent=`${idx}: ${name}`; codeColSel.appendChild(o1);
    const o2=document.createElement('option'); o2.value=idx; o2.textContent=`${idx}: ${name}`; linksColSel.appendChild(o2);
  });
  codeColSel.value = codeIdxAuto; linksColSel.value = linksIdxAuto;

  const body = headerRow ? rawRows.slice(1, 51) : rawRows.slice(0, 50);
  const colsIndex = [...Array(cols.length).keys()];
  let html = '<table><thead><tr>'; cols.forEach(c=> html += `<th>${c}</th>`); html += '</tr></thead><tbody>';
  body.forEach(r=>{ html+='<tr>'; colsIndex.forEach(ci=> html+=`<td>${(r[ci]??'')}</td>`); html+='</tr>'; }); html += '</tbody></table>';
  previewRows.innerHTML = html;

  const temp = buildRowsFrom(headerRow? rawRows.slice(1):rawRows, Number(codeColSel.value), Number(linksColSel.value));
  previewPics.innerHTML = '';
  let shown=0;
  for(const pr of temp){
    for(const url of pr.links){
      if(shown>=100) break;
      const card=document.createElement('div'); card.className='pic-card';
      const img=document.createElement('img'); img.loading='lazy'; img.referrerPolicy='no-referrer'; img.src=url; img.onerror=()=>card.classList.add('broken');
      const cap=document.createElement('div'); cap.className='pic-cap'; cap.textContent=pr.code;
      card.appendChild(img); card.appendChild(cap); previewPics.appendChild(card); shown++;
    }
    if(shown>=100) break;
  }
  tabRows.onclick=()=>{ tabRows.classList.add('active'); tabPics.classList.remove('active'); previewRows.classList.remove('hidden'); previewPics.classList.add('hidden'); };
  tabPics.onclick=()=>{ tabPics.classList.add('active'); tabRows.classList.remove('active'); previewRows.classList.add('hidden'); previewPics.classList.remove('hidden'); };

  modal.classList.remove('hidden');
}
byId('modalClose').addEventListener('click', ()=> modal.classList.add('hidden'));
applyColsBtn.addEventListener('click', ()=>{
  const codeIdx = Number(codeColSel.value), linksIdx = Number(linksColSel.value);
  const body = headerRow ? rawRows.slice(1) : rawRows;
  parsedRows = buildRowsFrom(body, codeIdx, linksIdx);
  renderKpis(parsedRows); updateButtons(); modal.classList.add('hidden'); toast('تم تطبيق الاعمدة','success');
});

// ===== help modal =====
helpBtn.addEventListener('click', ()=> helpModal.classList.remove('hidden'));
helpClose.addEventListener('click', ()=> helpModal.classList.add('hidden'));

// ===== remote upload (with JSON fallback) =====
function extractFolderId(link){
  if(!link) return '';
  try{
    const u = new URL(link);
    if(u.pathname.includes('/folders/')){ const parts=u.pathname.split('/'); const id=parts[parts.indexOf('folders')+1]; return id||''; }
    if(u.searchParams.get('id')) return u.searchParams.get('id');
  }catch{} return '';
}
function blobToBase64(blob){
  return new Promise((resolve,reject)=>{
    const r=new FileReader();
    r.onload=()=>{ const s=String(r.result||''); resolve(s.substring(s.indexOf(',')+1)); };
    r.onerror=reject; r.readAsDataURL(blob);
  });
}
async function uploadBlobMultipart(endpoint, blob, filename, {folderId='', sessionName=''}={}){
  const form = new FormData();
  form.append('file', blob, filename);
  form.append('filename', filename);
  if(folderId) form.append('folderId', folderId);
  if(sessionName) form.append('session', sessionName);
  const res = await fetch(endpoint, { method:'POST', body:form });
  const data = await res.json().catch(()=>({ ok: res.ok }));
  if(!res.ok || data?.ok === false) throw new Error(data?.error || `HTTP ${res.status}`);
  return data;
}
async function uploadBlobJSON(endpoint, blob, filename, {folderId='', sessionName=''}={}) {
    const base64 = await blobToBase64(blob);
    const payload = { filename, folderId, session: sessionName, type: blob.type || 'application/octet-stream', data: base64 };
    // بدون Content-Type: يخليه simple request
    const res = await fetch(endpoint, { method:'POST', body: JSON.stringify(payload) });
    const data = await res.json().catch(()=>({ ok: res.ok }));
    if(!res.ok || data?.ok === false) throw new Error(data?.error || `HTTP ${res.status}`);
    return data;
  }
  
  
  async function uploadBlob(endpoint, blob, filename, opts={}){
    try{
      const r = await uploadBlobMultipart(endpoint, blob, filename, opts);
      log(`رفع (multipart): ${filename}`, 'info');
      return r;
    }catch(e){
      log(`التحويل ل JSON بسبب: ${e.message}`, 'warn');
      const r = await uploadBlobJSON(endpoint, blob, filename, opts);
      log(`رفع (json): ${filename}`, 'info');
      return r;
    }
  }
  
function createUploadQueue(concurrency){
  const q=[]; let running=0;
  function run(){ if(running>=concurrency || !q.length) return; const task=q.shift(); running++; (async()=>{ try{ await task(); } finally{ running--; run(); }})(); }
  return { push(task){ q.push(task); run(); }, async drain(){ while(q.length||running) await sleep(200); } };
}

// ===== download pieces =====
async function fetchWithRetry(url, {tries=3, timeoutMs=15000, backoff=800}={}){
  for(let attempt=1; attempt<=tries; attempt++){
    const ctrl=new AbortController(); const timer=setTimeout(()=>ctrl.abort(), timeoutMs);
    try{ const res = await fetch(url, { signal: ctrl.signal, redirect:'follow' }); clearTimeout(timer); if(!res.ok) throw new Error(`HTTP ${res.status}`); return res; }
    catch(e){ clearTimeout(timer); if(attempt===tries) throw e; await sleep(backoff*attempt + Math.random()*400); }
  }
}
function buildFilePath(pattern,{code,seq,ext},groupBy){ let p = pattern.replaceAll('${code}',code).replaceAll('${seq}',seq).replaceAll('${ext}', ext.replace(/^\./,'')); if(groupBy && !p.includes('/')) p = `${code}/${p}`; return p; }
function exportCSV(rows, name){
  const header = ['product_code','url','status','http_status','size_bytes','filename','zip_part','error'];
  const lines=[header.join(',')];
  rows.forEach(r=>{
    const vals=[r.code,r.url,r.status,r.http_status??'',r.size??'',r.filename??'',r.zip_part??'',(r.error||'').replace(/[\r\n,]/g,' ')];
    lines.push(vals.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(','));
  });
  const csv=lines.join('\n'); saveAs(new Blob([csv],{type:'text/csv;charset=utf-8;'}), name);
}

// ===== Excel export =====
function exportDrivePicsExcel(onePerCode){
  if(uploadedPictureLinks.length===0){ toast('لا توجد روابط صور مرفوعة','warn'); return; }
  let rows = uploadedPictureLinks;
  if(onePerCode){
    const seen=new Set(), tmp=[];
    for(const r of uploadedPictureLinks){ if(!r.link) continue; if(seen.has(r.code)) continue; seen.add(r.code); tmp.push(r); }
    rows = tmp;
  }
  const data=[['product_code','drive_link','filename']]; rows.forEach(r=> data.push([r.code, r.link||'', r.filename||'']));
  const wb=XLSX.utils.book_new(); const ws=XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Pictures'); XLSX.writeFile(wb, onePerCode? 'drive_links_pictures_first.xlsx' : 'drive_links_pictures_all.xlsx');
  toast('تم تصدير Excel لروابط الصور','success');
}
function exportDriveZipsExcel(){
  if(uploadedZipLinks.length===0){ toast('لا توجد روابط ZIP مرفوعة','warn'); return; }
  const data=[['zip_name','drive_link']]; uploadedZipLinks.forEach(r=> data.push([r.name, r.link||'']));
  const wb=XLSX.utils.book_new(); const ws=XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'ZIPs'); XLSX.writeFile(wb, 'drive_links_zips.xlsx');
  toast('تم تصدير Excel لروابط ZIP','success');
}

// ===== main download (local) =====
async function downloadAllLocal(rows = parsedRows){
  try{
    const pool = Math.max(1, Math.min(12, Number(concurrencyInput.value||6)));
    const maxFiles = Math.max(50, Number(maxZipFilesInput.value||300));
    const maxBytes = Math.max(50, Number(maxZipMBInput.value||200))*1024*1024;
    const perHost = Math.max(1, Math.min(8, Number(perHostInput.value||4)));
    const maxPerProduct = Math.max(0, Number((maxPerProductInput?.value ?? 0)));
    const groupBy = Boolean(groupByProductInput.checked);
    const pattern = (namingPatternInput.value||'${code}/${code}_${seq}.${ext}').trim();

    const jobs=[];
    for(const {code,links} of rows){ const lim = maxPerProduct? Math.min(maxPerProduct, links.length) : links.length; for(let i=0;i<lim;i++) jobs.push({code, url:links[i], seq: pad2(i+1)}); }
    const wl = whitelistInput.value.split(',').map(s=>s.trim()).filter(Boolean);
    const bl = blacklistInput.value.split(',').map(s=>s.trim()).filter(Boolean);
    const allowHost = (h)=> (wl.length? wl.some(x=>h.endsWith(x)) : true) && (bl.length? !bl.some(x=>h.endsWith(x)) : true);
    const filteredJobs = jobs.filter(j=> allowHost(hostOf(j.url)));
    const totalImages = filteredJobs.length; if(!totalImages){ toast('لا توجد روابط مطابقة للفلترة','warn'); return; }

    reportRows.length=0; failures.length=0; zipParts.length=0; byId('zipList').textContent=''; uploadedZipLinks.length=0; exportDriveZipsXlsxBtn.disabled = true;

    let done=0, filesInZip=0, sizeInZip=0, part=1; setProgress(0,totalImages,'بدء التحميل');
    let zip = new JSZip();

    pauseBtn.disabled=false; resumeBtn.disabled=true; cancelBtn.disabled=false;
    let paused=false, cancel=false;
    pauseBtn.onclick=()=>{ paused=true; pauseBtn.disabled=true; resumeBtn.disabled=false; toast('تم الايقاف المؤقت','info'); };
    resumeBtn.onclick=()=>{ paused=false; pauseBtn.disabled=false; resumeBtn.disabled=true; toast('تم الاستكمال','success'); };
    cancelBtn.onclick=()=>{ cancel=true; pauseBtn.disabled=true; resumeBtn.disabled=true; cancelBtn.disabled=true; toast('جاري الالغاء','warn'); };

    const startTs=performance.now(); let lastTs=startTs;

    const inFlightByHost = new Map();
    async function schedulePerHost(url){
      const h = hostOf(url)||'unknown';
      while((inFlightByHost.get(h)||0) >= perHost) await sleep(100);
      inFlightByHost.set(h, (inFlightByHost.get(h)||0)+1);
      try{ return await fetchWithRetry(url); } finally{ inFlightByHost.set(h, inFlightByHost.get(h)-1); }
    }
    function ensureImage(ct){ if(!/^image\//i.test(ct||'')) throw new Error(`Not an image (${ct||'unknown'})`); }

    async function finalizeZip(){
      if(!filesInZip) return;
      setProgress(done,totalImages,`تجهيز ZIP ${part}`);
      const blob = await zip.generateAsync({type:'blob', streamFiles:true, compression:'DEFLATE', compressionOptions:{level:6}});
      const d=new Date(); const name=`images_part${String(part).padStart(2,'0')}_${d.getFullYear()}${pad2(d.getMonth()+1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}.zip`;
      saveAs(blob, name); zipParts.push({name, blob}); byId('zipList').textContent = `ZIP جاهزة: ${zipParts.map(z=>z.name).join(' , ')}`;
      log(`Saved ${name} (${filesInZip} files).`,'success');
      zip = new JSZip(); filesInZip=0; sizeInZip=0; part+=1; uploadZipsBtn.disabled = zipParts.length === 0;
    }

    let i=0;
    async function worker(){
      while(i<filteredJobs.length){
        if(cancel) throw new Error('Cancelled'); while(paused){ await sleep(150); if(cancel) throw new Error('Cancelled'); }
        const job = filteredJobs[i++];
        try{
          const res = await schedulePerHost(job.url);
          const ct = res.headers.get('content-type')||''; ensureImage(ct);
          const blob = await res.blob();
          const ext = (inferExt(job.url, ct) === '(unknown)') ? '.bin' : inferExt(job.url, ct);
          const path = buildFilePath(pattern, {code:job.code, seq:job.seq, ext}, groupBy);

          if(filesInZip>=maxFiles || (sizeInZip+blob.size)>maxBytes) await finalizeZip();
          zip.file(path, blob); filesInZip++; sizeInZip += blob.size;

          reportRows.push({code:job.code, url:job.url, status:'ok', http_status:200, size:blob.size, filename:path, zip_part:part});
          done++; const now=performance.now(); if(now-lastTs>1000){ const elapsedS=(now-startTs)/1000, speed=(done/elapsedS).toFixed(2); const remain=totalImages-done, eta = Math.max(0, Math.round(remain / Math.max(0.1, done/elapsedS))); etaText.textContent = `ETA: ${eta}ث | السرعة: ${speed} روابط/ث`; lastTs=now; }
          setProgress(done,totalImages,`تم تنزيل ${path}`);
        }catch(err){
          reportRows.push({code:job.code, url:job.url, status:'fail', http_status:'', size:'', filename:'', zip_part:part, error:err.message});
          failures.push({code:job.code, url:job.url, error:err.message});
          done++; setProgress(done,totalImages,`فشل: ${job.code}_${job.seq}`); log(`فشل ${job.code} ${job.url} → ${err.message}`,'error');
        }
      }
    }

    await Promise.all(Array.from({length:pool}, worker));
    await finalizeZip(); setProgress(done, done, 'انتهى التحميل'); toast('انتهى التحميل','success');
  }catch(e){ toast(`خطأ غير متوقع: ${e.message}`,'error'); }
  finally{
    pauseBtn.disabled=true; resumeBtn.disabled=true; cancelBtn.disabled=true;
    exportReportBtn.disabled = reportRows.length===0; exportFailuresBtn.disabled = failures.length===0;
    uploadZipsBtn.disabled = zipParts.length===0; uploadPicsBtn.disabled = parsedRows.length===0 ? true : false;
  }
}

// ===== uploads =====
function extractFolderAndEndpoint(){
  const endpoint = (remoteEndpointInput.value||'').trim();
  const folderId = extractFolderId((remoteFolderLinkInput.value||'').trim());
  const sessionName = (remoteSessionInput.value||'').trim();
  const upConc = Math.max(1, Math.min(4, Number(remoteConcurrencyInput.value||2)));
  if(!endpoint){ toast('ادخل رابط Web App اولا','warn'); }
  else if(endpoint.includes('script.google.com') && !folderId){
    toast('تنبيه: سيتم الرفع الى المجلد الافتراضي عند صاحب الويب آب. يفضل نشر ويب آب خاص بك او وضع رابط مجلدك.', 'warn', 6000);
  }
  return { endpoint, folderId, sessionName, upConc };
}
function createUploadQueueSimple(conc){ return createUploadQueue(conc); }

async function uploadZipPartsToDrive(){
  const { endpoint, folderId, sessionName, upConc } = extractFolderAndEndpoint(); if(!endpoint) return;
  if(zipParts.length===0){ toast('لا يوجد ZIP. نفّذ التحميل أولاً أو اختر تحميل محلي ضمن المهام.','warn'); return; }

  uploadedZipLinks.length = 0;
  const q = createUploadQueueSimple(upConc); let done=0, total=zipParts.length; setProgress(0,total,'رفع ZIP...');
  zipParts.forEach(z=>{
    q.push(async ()=>{
      try{
        const resp = await uploadBlob(endpoint, z.blob, z.name, {folderId, sessionName});
        const link = resp?.webViewLink || resp?.url || ''; uploadedZipLinks.push({name:z.name, link});
        log(`Uploaded ZIP: <a href="${link}" target="_blank" rel="noopener">${z.name}</a>`, 'success', true);
      }catch(e){ log(`Upload ZIP failed: ${z.name} → ${e.message}`,'error'); }
      done++; setProgress(done,total,'رفع ZIP...');
    });
  });
  await q.drain(); exportDriveZipsXlsxBtn.disabled = uploadedZipLinks.length===0; toast('انتهى رفع ZIP','success');
}

async function uploadPicturesOnly(rows = parsedRows){
  const { endpoint, folderId, sessionName } = extractFolderAndEndpoint(); if(!endpoint) return;

  const pool = Math.max(1, Math.min(12, Number(concurrencyInput.value||6)));
  const perHost = Math.max(1, Math.min(8, Number(perHostInput.value||4)));
  const maxPerProduct = Math.max(0, Number(maxPerProductInput.value||0));
  const pattern = (namingPatternInput.value||'${code}_${seq}.${ext}').trim();
  const groupBy = groupByProductInput.checked;

  const jobs=[]; for(const {code,links} of rows){ const lim = maxPerProduct? Math.min(maxPerProduct, links.length) : links.length; for(let i=0;i<lim;i++) jobs.push({code, url:links[i], seq: pad2(i+1)}); }
  const wl = whitelistInput.value.split(',').map(s=>s.trim()).filter(Boolean);
  const bl = blacklistInput.value.split(',').map(s=>s.trim()).filter(Boolean);
  const allowHost = (h)=> (wl.length? wl.some(x=>h.endsWith(x)) : true) && (bl.length? !bl.some(x=>h.endsWith(x)) : true);
  const filteredJobs = jobs.filter(j=> allowHost(hostOf(j.url)));
  const total = filteredJobs.length; if(!total){ toast('لا توجد روابط مطابقة للفلترة','warn'); return; }

  uploadedPictureLinks.length = 0;

  let paused=false, cancel=false;
  pauseBtn.disabled=false; resumeBtn.disabled=true; cancelBtn.disabled=false;
  pauseBtn.onclick=()=>{ paused=true; pauseBtn.disabled=true; resumeBtn.disabled=false; };
  resumeBtn.onclick=()=>{ paused=false; pauseBtn.disabled=false; resumeBtn.disabled=true; };
  cancelBtn.onclick=()=>{ cancel=true; pauseBtn.disabled=true; resumeBtn.disabled=true; cancelBtn.disabled=true; };

  const inFlightByHost = new Map();
  async function schedulePerHost(url){
    const h = hostOf(url)||'unknown';
    while((inFlightByHost.get(h)||0) >= perHost) await sleep(100);
    inFlightByHost.set(h, (inFlightByHost.get(h)||0)+1);
    try{ return await fetchWithRetry(url); } finally{ inFlightByHost.set(h, inFlightByHost.get(h)-1); }
  }

  let done=0; setProgress(0,total,'رفع الصور...'); let i=0;
  async function worker(){
    while(i<filteredJobs.length){
      if(cancel) throw new Error('Cancelled'); while(paused){ await sleep(150); if(cancel) throw new Error('Cancelled'); }
      const job = filteredJobs[i++];
      try{
        const res = await schedulePerHost(job.url);
        const ct = res.headers.get('content-type')||''; if(!/^image\//i.test(ct)) throw new Error(`Not an image (${ct||'unknown'})`);
        const blob = await res.blob();
        const ext = (inferExt(job.url, ct) === '(unknown)') ? '.bin' : inferExt(job.url, ct);
        const path = buildFilePath(pattern, {code:job.code, seq:job.seq, ext}, groupBy);

        const resp = await uploadBlob(endpoint, blob, path.replace(/\//g,'__'), {folderId, sessionName});
        const link = resp?.webViewLink || resp?.url || '';
        uploadedPictureLinks.push({ code: job.code, link, filename: path });
        log(`Uploaded: <a href="${link}" target="_blank" rel="noopener">${path}</a>`,'success', true);

        done++; setProgress(done,total,'رفع الصور...');
      }catch(e){ toast(`فشل رفع صورة: ${e.message}`,'error',4000); log(`Upload failed: ${job.code} ${job.url} → ${e.message}`,'error'); done++; setProgress(done,total,'رفع الصور...'); }
    }
  }

  try{ await Promise.all(Array.from({length:pool}, worker)); toast('انتهى رفع الصور','success'); }
  catch(e){ if(e.message==='Cancelled') toast('تم الالغاء','warn'); else toast(`خطأ: ${e.message}`,'error'); }
  finally{ pauseBtn.disabled=true; resumeBtn.disabled=true; cancelBtn.disabled=true; exportDrivePicsXlsxBtn.disabled = uploadedPictureLinks.length===0; }
}

// ===== events =====
function updateButtons(){
  const can = parsedRows.length>0;
  startBtn.disabled = !can; estimateBtn.disabled = !can; previewBtn.disabled = rawRows.length===0; previewBtnSide.disabled = rawRows.length===0; uploadPicsBtn.disabled = !can;
}
['dragenter','dragover','dragleave','drop'].forEach(ev=> fileLabel.addEventListener(ev,(e)=>{e.preventDefault(); e.stopPropagation();}));
['dragenter','dragover'].forEach(()=> fileLabel.classList.add('drag')); ['dragleave','drop'].forEach(()=> fileLabel.classList.remove('drag'));
fileLabel.addEventListener('drop', (e)=>{ const dt=e.dataTransfer; if(dt?.files?.[0]){ fileInput.files=dt.files; fileInput.dispatchEvent(new Event('change')); }});

fileInput.addEventListener('change', ()=>{
  const f=fileInput.files?.[0]; fileNameEl.textContent = f ? f.name : 'لا يوجد ملف'; fileMetaEl.textContent = f ? `(${(f.size/1024/1024).toFixed(2)} MB)` : '';
  logEl.innerHTML=''; setProgress(0,0,'قراءة الملف'); parsedRows=[]; rawRows=[]; headerRow=null;
  if(!f){ startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; statStatus.textContent='لا يوجد ملف'; return; }
  if(!/\.(xlsx|xls|csv)$/i.test(f.name)){ toast('نوع الملف غير مدعوم. استخدم xlsx او xls او csv','error'); fileInput.value=''; return; }

  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type:'array' });
      const sheet = (sheetNameInput.value||'').trim();
      parsedRows = parseWorkbook(wb, sheet);
      renderKpis(parsedRows);
      const totalLinks = parsedRows.reduce((a,r)=>a+r.links.length,0);
      setProgress(0,totalLinks,'جاهز');
      log(`تم قراءة ${parsedRows.length} منتج فيه روابط و ${totalLinks} رابط.`,'success');
      log('نصيحة: خليك بين 4-8 تحميلات متوازية لاستقرار افضل.','info');
      updateButtons(); showGuideToasts();
      if(manualSelectInput.checked){
        const cols = (headerRow? headerRow : rawRows[0]).map((v,i)=> (headerRow? String(v||`Col ${i}`): `Col ${i}`));
        openModal(cols);
      }
    }catch(err){ startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; setProgress(0,0,'خطأ في القراءة'); log(`Parse error: ${err.message}`,'error'); toast(`خطأ في قراءة الملف: ${err.message}`,'error',5000); }
  };
  reader.onerror = ()=>{ startBtn.disabled=true; estimateBtn.disabled=true; previewBtn.disabled=true; setProgress(0,0,'Read error'); log('Could not read file.','error'); toast('تعذر قراءة الملف','error'); };
  reader.readAsArrayBuffer(f);
});
// وصلة زر تصدير روابط الصور إلى Excel
exportDrivePicsXlsxBtn.addEventListener('click', () => {
    exportDrivePicsExcel(!!xlOnePerCodeChk.checked);
  });

  
function openPreviewIfReady(){ if(!rawRows.length) return; const cols = (headerRow? headerRow : rawRows[0]).map((v,i)=> (headerRow? String(v||`Col ${i}`): `Col ${i}`)); openModal(cols); }
previewBtn.addEventListener('click', openPreviewIfReady);
previewBtnSide.addEventListener('click', openPreviewIfReady);
estimateBtn.addEventListener('click', ()=> estimateSizes(parsedRows));

startBtn.addEventListener('click', async ()=>{
  const wantDownload = taskDownload.checked;
  const wantPics = taskDrivePics.checked;
  const wantZips = taskDriveZips.checked;
  if(!wantDownload && !wantPics && !wantZips){ toast('اختر مهمة واحدة على الاقل','warn'); return; }

  if(wantDownload){ await downloadAllLocal(parsedRows); }
  if(wantPics){ await uploadPicturesOnly(parsedRows); }
  if(wantZips){ await uploadZipPartsToDrive(); }
});

exportReportBtn.addEventListener('click', ()=> exportCSV(reportRows, 'download_report.csv'));
exportFailuresBtn.addEventListener('click', ()=> exportCSV(failures, 'download_failures.csv'));

uploadZipsBtn.addEventListener('click', uploadZipPartsToDrive);
uploadPicsBtn.addEventListener('click', ()=> uploadPicturesOnly(parsedRows));

testUploadBtn.addEventListener('click', async ()=>{
  const { endpoint, folderId, sessionName } = extractFolderAndEndpoint(); if(!endpoint) return;
  try{
    const blob = new Blob([`test ${new Date().toISOString()}`], {type:'text/plain'});
    const resp = await uploadBlob(endpoint, blob, 'test.txt', {folderId, sessionName});
    const link = resp?.webViewLink || resp?.url || '';
    log(`Test OK: <a href="${link}" target="_blank" rel="noopener">${link||'open file'}</a>`, 'success', true);
    toast('تم اختبار الرفع بنجاح','success');
  }catch(e){ log(`Test failed: ${e.message}`, 'error'); toast(`فشل الاختبار: ${e.message}`, 'error', 6000); }
});
