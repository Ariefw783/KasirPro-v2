/* KasirPro AT-08 — Final invoice interaction + professional UI cleanup */
import { STORE_KEYS, readStore } from '../modules/database/database-store.js';

const text=v=>String(v??'').trim();
const norm=v=>text(v).toLowerCase();
const num=v=>Number(String(v??0).replace(/[^0-9.-]/g,''))||0;
const escapeHtml=value=>String(value??'').replace(/[&<>"']/g,char=>({
  '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
}[char]));
const rupiah=v=>new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',maximumFractionDigits:0}).format(num(v));
const invoices=()=>readStore(STORE_KEYS.invoices,[])||[];

function invoiceSubtotal(item){
  return Math.max(0,num(item.qty)*num(item.buyPrice)-num(item.discount));
}
function invoiceTotal(inv){
  const itemTotal=(inv.items||[]).reduce((sum,item)=>sum+invoiceSubtotal(item),0);
  const base=Math.max(0,itemTotal-num(inv.discount));
  const perItem=norm(inv.taxMethod)==='per item'||norm(inv.taxMethod)==='per_product';
  const tax=perItem
    ?(inv.items||[]).reduce((sum,item)=>sum+invoiceSubtotal(item)*num(item.taxPercent)/100,0)
    :base*num(inv.taxPercent)/100;
  return base+tax+num(inv.otherCost);
}
function discountLabel(method,value,amount){
  if(method==='percentage')return `${num(value)}% (${rupiah(amount)})`;
  return rupiah(amount);
}
function dateLabel(v){if(!v)return '—';const d=new Date(v);return Number.isNaN(d.getTime())?text(v):d.toLocaleDateString('id-ID',{day:'2-digit',month:'2-digit',year:'numeric'});}
function section(){return document.querySelector('[data-view-section="purchase-invoices"]');}
function fileInput(){return document.getElementById('invoice-import-file');}

function installStyle(){
  if(document.getElementById('kp-invoice-final-style'))return;
  const style=document.createElement('style');
  style.id='kp-invoice-final-style';
  style.textContent=`
  .kp-final-hidden{display:none!important}
  .kp-invoice-final{max-width:1380px;margin:0 auto;padding:2px 0 28px;display:grid;gap:20px;color:#0f172a}
  .kp-invoice-final *{box-sizing:border-box}
  .kp-final-head{display:flex;justify-content:space-between;gap:24px;align-items:flex-start}
  .kp-final-head h2{font-size:24px;line-height:1.2;margin:0 0 7px;font-weight:800;letter-spacing:-.02em;color:#0f172a}
  .kp-final-head p{margin:0;max-width:760px;color:#64748b;font-size:13px;line-height:1.6}
  .kp-final-source{font-size:12px;font-weight:700;color:#0f766e;background:#f0fdfa;border:1px solid #ccfbf1;padding:7px 10px;border-radius:8px;white-space:nowrap}
  .kp-final-card{background:#fff;border:1px solid #e2e8f0;border-radius:14px;overflow:hidden}
  .kp-final-card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;padding:17px 18px;border-bottom:1px solid #eef2f7}
  .kp-final-card-head h3{margin:0 0 4px;font-size:15px;font-weight:800;color:#0f172a}.kp-final-card-head p{margin:0;color:#64748b;font-size:12px;line-height:1.5}
  .kp-final-import{padding:18px;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:16px;align-items:center}
  .kp-final-filebox{min-width:0;display:flex;align-items:center;gap:13px;padding:13px 14px;border:1px solid #dbe3ea;border-radius:11px;background:#f8fafc}
  .kp-final-fileicon{width:38px;height:38px;border-radius:9px;background:#ecfdf5;color:#047857;display:grid;place-items:center;flex:0 0 auto}
  .kp-final-filemeta{min-width:0;flex:1}.kp-final-filemeta strong{display:block;font-size:13px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.kp-final-filemeta span{display:block;margin-top:3px;font-size:11px;color:#94a3b8}
  .kp-file-trigger{display:inline-flex;align-items:center;justify-content:center;gap:7px;min-height:38px;padding:0 12px;border:1px solid #cbd5e1;border-radius:9px;background:#fff;color:#334155;font-size:12px;font-weight:750;cursor:pointer;white-space:nowrap}
  .kp-file-trigger:hover{background:#f8fafc;border-color:#94a3b8}
  #invoice-import-file{position:absolute!important;width:1px!important;height:1px!important;opacity:0!important;overflow:hidden!important;pointer-events:none!important;clip:rect(0 0 0 0)!important;clip-path:inset(50%)!important}
  .kp-final-actions{display:flex;gap:9px;flex-wrap:wrap;justify-content:flex-end}
  .kp-final-btn{min-height:40px;padding:0 13px;border-radius:9px;border:1px solid #cbd5e1;background:#fff;color:#334155;font:inherit;font-size:12px;font-weight:750;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:7px}
  .kp-final-btn:hover{background:#f8fafc}.kp-final-btn.primary{border-color:#0f766e;background:#0f766e;color:#fff}.kp-final-btn.primary:hover{background:#115e59}.kp-final-btn:disabled{opacity:.46;cursor:not-allowed}
  .kp-final-preview{margin:0 18px 18px;border:1px solid #dbe3ea;border-radius:11px;overflow:hidden;background:#fff}.kp-final-preview[hidden]{display:none!important}
  .kp-final-preview-head{display:flex;justify-content:space-between;gap:12px;align-items:center;background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:11px 13px}.kp-final-preview-head strong{font-size:13px}.kp-final-preview-head span{font-size:11px;color:#64748b}
  .kp-final-preview-body{padding:13px}.kp-final-preview-body #invoice-import-summary{font-size:12px;line-height:1.55;color:#475569;margin:0 0 10px}.kp-final-preview-body #invoice-import-preview{margin:0!important}.kp-final-preview-footer{display:flex;justify-content:flex-end;padding-top:12px;border-top:1px solid #eef2f7;margin-top:12px}
  .kp-final-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.kp-final-kpi{padding:15px 16px;background:#fff;border:1px solid #e2e8f0;border-radius:12px}.kp-final-kpi span{display:block;font-size:11px;font-weight:700;color:#64748b}.kp-final-kpi strong{display:block;margin-top:6px;font-size:21px;font-weight:800;letter-spacing:-.02em;color:#0f172a}.kp-final-kpi small{display:block;margin-top:3px;font-size:10.5px;color:#94a3b8}
  .kp-final-toolbar{display:flex;gap:9px;flex-wrap:wrap;padding:13px 16px;border-bottom:1px solid #eef2f7;background:#fbfcfd}.kp-final-toolbar input,.kp-final-toolbar select{height:38px;border:1px solid #cbd5e1;border-radius:9px;background:#fff;color:#334155;padding:0 10px;font-size:12px;outline:none}.kp-final-toolbar input:focus,.kp-final-toolbar select:focus{border-color:#0f766e;box-shadow:0 0 0 3px rgba(15,118,110,.08)}.kp-final-toolbar input{flex:1;min-width:220px}.kp-final-toolbar select{min-width:155px}
  .kp-final-tablewrap{overflow:auto}.kp-final-table{width:100%;border-collapse:collapse}.kp-final-table th{padding:10px 13px;background:#f8fafc;border-bottom:1px solid #e2e8f0;text-align:left;font-size:10.5px;letter-spacing:.035em;text-transform:uppercase;color:#64748b;white-space:nowrap}.kp-final-table td{padding:12px 13px;border-bottom:1px solid #eef2f7;font-size:12px;color:#334155;vertical-align:middle}.kp-final-table tbody tr:last-child td{border-bottom:0}.kp-final-table tbody tr:hover{background:#fbfdfd}.kp-final-num{font-weight:800;color:#0f172a}.kp-final-money{font-weight:750;color:#0f172a;font-variant-numeric:tabular-nums}.kp-final-status{display:inline-flex;padding:5px 8px;border-radius:999px;font-size:10.5px;font-weight:800}.kp-final-status.draft{background:#fff7ed;color:#c2410c}.kp-final-status.done{background:#ecfdf5;color:#15803d}.kp-final-row-actions{display:flex;gap:7px;justify-content:flex-end}.kp-final-row-actions button{height:32px;padding:0 9px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-size:11px;font-weight:750;cursor:pointer}.kp-final-row-actions button.confirm{border-color:#99f6e4;background:#f0fdfa;color:#0f766e}.kp-final-empty{text-align:center!important;color:#94a3b8!important;padding:30px 16px!important}
  @media(max-width:980px){.kp-final-import{grid-template-columns:1fr}.kp-final-actions{justify-content:flex-start}.kp-final-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.kp-final-head{flex-direction:column;gap:10px}}
  @media(max-width:640px){.kp-invoice-final{gap:15px}.kp-final-kpis{grid-template-columns:1fr 1fr}.kp-final-filebox{align-items:flex-start;flex-wrap:wrap}.kp-file-trigger{width:100%}.kp-final-actions{display:grid;grid-template-columns:1fr 1fr}.kp-final-btn{width:100%}.kp-final-toolbar input,.kp-final-toolbar select{width:100%;min-width:0}.kp-final-row-actions{min-width:170px}}
  `;
  document.head.appendChild(style);
}

function build(){
  const target=section();if(!target)return null;
  let root=document.getElementById('kp-invoice-final');if(root)return root;
  installStyle();
  [...target.children].forEach(child=>child.classList.add('kp-final-hidden'));
  document.querySelectorAll('[data-view="import-faktur"]').forEach(el=>el.classList.add('kp-final-hidden'));
  const oldImport=document.querySelector('[data-view-section="import-faktur"]');if(oldImport)oldImport.classList.add('kp-final-hidden');

  root=document.createElement('div');root.id='kp-invoice-final';root.className='kp-invoice-final';
  root.innerHTML=`
    <div class="kp-final-head">
      <div><h2>Faktur Pembelian</h2><p>Impor, periksa, dan simpan faktur dari aplikasi Input Faktur Offline. Stok hanya berubah setelah Barang Masuk dikonfirmasi.</p></div>
      <div class="kp-final-source"><i class="fa-solid fa-file-excel"></i> Impor Excel</div>
    </div>
    <section class="kp-final-card">
       <div class="kp-final-card-head"><div><h3>Impor Faktur Excel</h3><p>Pilih file hasil aplikasi Input Faktur Offline, lalu periksa datanya sebelum disimpan sebagai Draft.</p></div></div>
      <div class="kp-final-import">
        <div class="kp-final-filebox">
          <div class="kp-final-fileicon"><i class="fa-solid fa-file-excel"></i></div>
          <div class="kp-final-filemeta"><strong id="kp-final-file-name">Belum ada file dipilih</strong><span id="kp-final-file-info">Format XLSX / XLS</span></div>
          <label class="kp-file-trigger" for="invoice-import-file"><i class="fa-solid fa-folder-open"></i> Pilih File</label>
        </div>
        <div class="kp-final-actions">
          <button type="button" class="kp-final-btn primary" id="kp-final-preview" disabled><i class="fa-solid fa-magnifying-glass"></i> Baca & Preview</button>
        </div>
      </div>
      <div class="kp-final-preview" id="kp-final-preview-box" hidden>
        <div class="kp-final-preview-head"><strong>Preview Faktur</strong><span>Periksa data sebelum disimpan</span></div>
        <div class="kp-final-preview-body" id="kp-final-preview-content"></div>
      </div>
    </section>
    <div class="kp-final-kpis">
      <div class="kp-final-kpi"><span>Total Faktur</span><strong id="kp-final-total">0</strong><small>Semua faktur pembelian</small></div>
      <div class="kp-final-kpi"><span>Draft</span><strong id="kp-final-draft">0</strong><small>Menunggu Barang Masuk</small></div>
      <div class="kp-final-kpi"><span>Barang Masuk</span><strong id="kp-final-done">0</strong><small>Sudah menambah stok</small></div>
      <div class="kp-final-kpi"><span>Nilai Faktur</span><strong id="kp-final-value">Rp0</strong><small>Total nilai pembelian</small></div>
    </div>
    <section class="kp-final-card">
       <div class="kp-final-card-head"><div><h3>Daftar Faktur Pembelian</h3><p>Perbaikan Draft dilakukan di aplikasi offline, kemudian file Excel diimpor ulang untuk mengganti Draft lama.</p></div></div>
      <div class="kp-final-toolbar">
        <input id="kp-final-search" type="search" placeholder="Cari nomor faktur atau supplier">
        <select id="kp-final-status"><option value="all">Semua Status</option><option value="draft">Draft</option><option value="done">Barang Masuk</option></select>
        <select id="kp-final-supplier"><option value="all">Semua Perusahaan Supplier</option></select>
      </div>
      <div class="kp-final-tablewrap"><table class="kp-final-table"><thead><tr><th>Nomor Faktur</th><th>Tanggal</th><th>Supplier</th><th>Item</th><th>Total</th><th>Status</th><th style="text-align:right">Aksi</th></tr></thead><tbody id="kp-final-table-body"></tbody></table></div>
    </section>`;
  target.appendChild(root);
  return root;
}

function adoptSingleFileInput(){
  const root=build();if(!root)return;
  const input=fileInput();if(!input)return;
  if(!root.contains(input)) root.appendChild(input);
  input.removeAttribute('hidden');
  input.accept='.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel';
}

function adoptPreviewNodes(){
  const host=document.getElementById('kp-final-preview-content');if(!host)return;
  const summary=document.getElementById('invoice-import-summary');
  const preview=document.getElementById('invoice-import-preview');
  if(summary&&!host.contains(summary))host.appendChild(summary);
  if(preview&&!host.contains(preview))host.appendChild(preview);
  let footer=document.getElementById('kp-final-preview-footer');
  if(!footer){footer=document.createElement('div');footer.id='kp-final-preview-footer';footer.className='kp-final-preview-footer';footer.innerHTML='<button type="button" class="kp-final-btn primary" id="kp-final-save-draft"><i class="fa-solid fa-floppy-disk"></i> Simpan Draft</button>';host.appendChild(footer);}
  const legacyActions=document.getElementById('invoice-import-actions');if(legacyActions)legacyActions.classList.add('kp-final-hidden');
}

function fileChanged(){
  const input=fileInput(),file=input?.files?.[0];
  const name=document.getElementById('kp-final-file-name'),info=document.getElementById('kp-final-file-info'),preview=document.getElementById('kp-final-preview');
  if(name)name.textContent=file?.name||'Belum ada file dipilih';
  if(info)info.textContent=file?`${(file.size/1024).toFixed(1)} KB • ${file.name.toLowerCase().endsWith('.xls')?'XLS':'XLSX'}`:'Format XLSX / XLS';
  if(preview)preview.disabled=!file;
  const box=document.getElementById('kp-final-preview-box');if(box)box.hidden=true;
}

function filteredInvoices(){
  const q=norm(document.getElementById('kp-final-search')?.value),status=document.getElementById('kp-final-status')?.value||'all',supplier=norm(document.getElementById('kp-final-supplier')?.value||'all');
  return invoices().filter(inv=>{
    const done=!!inv.stockApplied||norm(inv.status)==='confirmed';
    if(status==='draft'&&done)return false;if(status==='done'&&!done)return false;
    if(supplier!=='all'&&norm(inv.supplierCode)!==supplier)return false;
    if(q&&!norm(`${inv.number} ${inv.supplierCode} ${inv.supplierName}`).includes(q))return false;
    return true;
  });
}

function render(){
  const all=invoices(),draft=all.filter(x=>!x.stockApplied&&norm(x.status)!=='confirmed'),done=all.length-draft.length;
  const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
  set('kp-final-total',all.length);set('kp-final-draft',draft.length);set('kp-final-done',done);set('kp-final-value',rupiah(all.reduce((s,x)=>s+invoiceTotal(x),0)));
  const supplierFilter=document.getElementById('kp-final-supplier');
  if(supplierFilter){
    const current=supplierFilter.value||'all';
    const supplierRows=[...new Map(all.map(inv=>[norm(inv.supplierCode||inv.supplierName),{value:text(inv.supplierCode||inv.supplierName),label:text(inv.supplierName||inv.supplierCode)}])).values()].filter(x=>x.value).sort((a,b)=>a.label.localeCompare(b.label,'id'));
    supplierFilter.innerHTML='<option value="all">Semua Perusahaan Supplier</option>'+supplierRows.map(x=>`<option value="${escapeHtml(x.value)}">${escapeHtml(x.label)}</option>`).join('');
    if([...supplierFilter.options].some(option=>option.value===current))supplierFilter.value=current;
  }
  const body=document.getElementById('kp-final-table-body');if(!body)return;
  const rows=filteredInvoices();
  body.innerHTML=rows.length?rows.map(inv=>{const isDone=!!inv.stockApplied||norm(inv.status)==='confirmed',id=escapeHtml(inv.id);return `<tr><td><span class="kp-final-num">${escapeHtml(text(inv.number)||'—')}</span></td><td>${escapeHtml(dateLabel(inv.date))}</td><td>${escapeHtml(text(inv.supplierName||inv.supplierCode)||'—')}</td><td>${(inv.items||[]).length}</td><td><span class="kp-final-money">${rupiah(invoiceTotal(inv))}</span></td><td><span class="kp-final-status ${isDone?'done':'draft'}">${isDone?'Barang Masuk':'Draft'}</span></td><td><div class="kp-final-row-actions"><button type="button" data-kp-view-invoice="${id}"><i class="fa-solid fa-eye"></i> Lihat</button>${isDone?'':`<button type="button" class="confirm" data-invoice-confirm="${id}"><i class="fa-solid fa-box-open"></i> Barang Masuk</button>`}</div></td></tr>`;}).join(''):'<tr><td colspan="7" class="kp-final-empty">Belum ada faktur yang sesuai.</td></tr>';
}

function openReadOnly(id){
  const inv=invoices().find(item=>String(item.id)===String(id));
  if(!inv)return window.KasirProDialog?.error?.('Faktur Tidak Ditemukan','Data faktur tidak tersedia.');
  const status=inv.stockApplied||norm(inv.status)==='confirmed'?'Barang Masuk':'Draft';
  const items=(inv.items||[]).map((item,index)=>
    `${index+1}. ${text(item.code)||'Tanpa kode'} — ${text(item.name)||'Produk'}\n`+
    `   ${num(item.qty)} × ${rupiah(item.buyPrice)} · Diskon ${discountLabel(item.discountMethod,item.discountValue,item.discount)} = ${rupiah(invoiceSubtotal(item))}`
  ).join('\n');
  return window.KasirProDialog?.info?.(
    `Faktur ${text(inv.number)||'—'}`,
    `Tanggal: ${dateLabel(inv.date)}\nSupplier: ${text(inv.supplierName||inv.supplierCode)||'—'}\nStatus: ${status}\nDiskon Faktur: ${discountLabel(inv.discountMethod,inv.discountValue,inv.discount)}\nTotal: ${rupiah(invoiceTotal(inv))}\n\n${items||'Tidak ada item.'}`,
    {confirmText:'Tutup'}
  );
}

async function runPreview(){
  const input=fileInput();if(!input?.files?.[0])return window.KasirProDialog?.warning?.('File Belum Dipilih','Pilih file Faktur Excel terlebih dahulu.');
  const api=window.KasirProInvoiceAT08;
  if(!api?.preview) return window.KasirProDialog?.error?.('Modul Import Belum Siap','Muat ulang aplikasi lalu coba kembali.');
  const btn=document.getElementById('kp-final-preview');if(btn){btn.disabled=true;btn.innerHTML='<i class="fa-solid fa-spinner fa-spin"></i> Membaca...';}
  try{
    adoptPreviewNodes();
    await api.preview();
    const box=document.getElementById('kp-final-preview-box');if(box)box.hidden=false;
    document.getElementById('kp-final-preview-box')?.scrollIntoView?.({behavior:'smooth',block:'nearest'});
  }catch(err){console.error('AT-08 Preview:',err);await window.KasirProDialog?.error?.('Preview Faktur Gagal',err?.message||String(err));}
  finally{if(btn){btn.disabled=false;btn.innerHTML='<i class="fa-solid fa-magnifying-glass"></i> Baca & Preview';}}
}

async function saveDraft(){
  const api=window.KasirProInvoiceAT08;
  if(!api?.saveDraft)return window.KasirProDialog?.error?.('Modul Faktur Belum Siap','Muat ulang aplikasi lalu coba kembali.');
  const btn=document.getElementById('kp-final-save-draft');if(btn){btn.disabled=true;btn.innerHTML='<i class="fa-solid fa-spinner fa-spin"></i> Menyimpan...';}
  try{await api.saveDraft();const box=document.getElementById('kp-final-preview-box');if(box)box.hidden=true;render();fileChanged();}
  catch(err){console.error('AT-08 Save Draft:',err);await window.KasirProDialog?.error?.('Faktur Tidak Dapat Disimpan',err?.message||String(err));}
  finally{if(btn){btn.disabled=false;btn.innerHTML='<i class="fa-solid fa-floppy-disk"></i> Simpan Draft';}}
}

function capture(event){
  const action=event.target.closest('button,a,label');if(!action)return;
  if(action.id==='kp-final-preview'){event.preventDefault();event.stopPropagation();event.stopImmediatePropagation();runPreview();return;}
  if(action.id==='kp-final-save-draft'){event.preventDefault();event.stopPropagation();event.stopImmediatePropagation();saveDraft();return;}
   const view=action.closest('[data-kp-view-invoice]');if(view){event.preventDefault();event.stopPropagation();event.stopImmediatePropagation();openReadOnly(view.dataset.kpViewInvoice);return;}
}

function bind(){
  const input=fileInput();if(input&&!input.dataset.kpFinalBound){input.dataset.kpFinalBound='1';input.addEventListener('change',fileChanged);}
  ['kp-final-search','kp-final-status','kp-final-supplier'].forEach(id=>{const el=document.getElementById(id);if(!el||el.dataset.kpFinalBound)return;el.dataset.kpFinalBound='1';el.addEventListener(id==='kp-final-search'?'input':'change',render);});
  if(!document.documentElement.dataset.kpInvoiceFinalCapture){document.documentElement.dataset.kpInvoiceFinalCapture='1';document.addEventListener('click',capture,true);}
}

function install(){
  build();adoptSingleFileInput();adoptPreviewNodes();bind();fileChanged();render();
}

if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',install,{once:true});else install();
window.addEventListener('kasirpro:operational-ready',()=>{install();render();});
window.KasirProInvoiceFinalAT08=Object.freeze({install,render,runPreview});
