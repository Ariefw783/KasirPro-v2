/* KasirPro AT-08 — Faktur Pembelian V2, single-sheet + formula template */
import { STORE_KEYS, readStore, writeStore } from '../modules/database/database-store.js';

const text = v => String(v ?? '').trim();
const norm = v => text(v).toLowerCase();
const num = v => Number(String(v ?? 0).replace(/[^0-9.-]/g,'')) || 0;
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({
  '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
}[char]));
const uid = prefix => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
const nowIso = () => new Date().toISOString();
let preview = [];
let issues = [];

function dialog(){ return window.KasirProDialog; }
function master(){ return readStore(STORE_KEYS.master,{produk:[],supplier:[],kategori:[],pengaturan_toko:[]}); }
function invoices(){ return readStore(STORE_KEYS.invoices,[]); }
function products(m=master()){ return Array.isArray(m.produk) ? m.produk : []; }
function suppliers(m=master()){ return Array.isArray(m.supplier) ? m.supplier : []; }
function canonicalSupplier(value,m=master()){
  const key=norm(value);
  if(!key) return '';
  const row=suppliers(m).find(x=>
    norm(x?.['Nama Supplier'])===key ||
    norm(x?.['Nama Perusahaan'])===key ||
    norm(x?.Supplier)===key ||
    norm(x?.['Kode Supplier'])===key
  );
  return text(row?.Supplier || row?.['Kode Supplier']);
}
function supplierName(label,m=master()){
  const row=suppliers(m).find(x=>norm(x?.Supplier)===norm(label)||norm(x?.['Kode Supplier'])===norm(label));
  return text(row?.['Nama Supplier'] || row?.['Nama Perusahaan']) || label;
}
function invoiceKey(invoice){return `${norm(invoice?.supplierCode||invoice?.supplierName)}::${norm(invoice?.number)}`;}
function byCode(code,m=master()){
  const key=norm(code); if(!key) return null;
  return products(m).find(x=>norm(x?.['Kode Produk'])===key) || null;
}
function dateOnly(value){
  if(!value) return '';
  if(value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0,10);
  const d=new Date(value); return Number.isNaN(d.getTime()) ? text(value) : d.toISOString().slice(0,10);
}
function discountMethod(value){
  const key=norm(value);
  if(['persen','persentase','%'].includes(key)) return 'percentage';
  if(['rupiah','rp','nominal'].includes(key)) return 'nominal';
  return '';
}
function discountAmount(method,value,base){
  const gross=Math.max(0,num(base));
  const entered=Math.max(0,num(value));
  return Math.min(gross,method==='percentage'?gross*entered/100:entered);
}
function paymentType(value){
  const key=norm(value);
  if(key==='cash'||key==='tunai') return 'Cash';
  if(['tempo','kredit','kredit/tempo'].includes(key)) return 'Tempo';
  return '';
}
function taxMethod(value){
  const key=norm(value);
  if(key==='global'||key==='ppn global') return 'Global';
  if(['per item','per produk','item','per_product'].includes(key)) return 'Per Item';
  return '';
}
function subtotal(item){ return Math.max(0,num(item.qty)*num(item.buyPrice)-num(item.discount)); }
function total(inv){
  const itemTotal=(inv.items||[]).reduce((sum,item)=>sum+subtotal(item),0);
  const base=Math.max(0,itemTotal-num(inv.discount));
  const tax=norm(inv.taxMethod)==='per item'
    ?(inv.items||[]).reduce((sum,item)=>sum+subtotal(item)*num(item.taxPercent)/100,0)
    :base*num(inv.taxPercent)/100;
  return base+tax+num(inv.otherCost);
}
function rupiah(v){ return new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',maximumFractionDigits:0}).format(num(v)); }

function fileInput(){ return document.getElementById('invoice-import-file'); }
function findImportSection(){
  const input=fileInput();
  return input?.closest('[data-view-section="import-faktur"]') || null;
}
function inInvoiceSection(el){ const root=findImportSection(); return !!(root && el && root.contains(el)); }

function makeTemplate(){
  const link=document.createElement('a');
  link.href='../data/templates/Template_Faktur_KasirPro.xlsx';
  link.download='Template_Faktur_KasirPro.xlsx';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function readWorkbook(){
  const file=fileInput()?.files?.[0];
  if(!file) throw new Error('Pilih file Excel faktur terlebih dahulu.');
  if(!window.XLSX?.read) throw new Error('SheetJS/XLSX belum tersedia.');
  const buffer=await file.arrayBuffer();
  return window.XLSX.read(buffer,{type:'array',cellDates:true});
}

async function buildPreview(){
  const wb=await readWorkbook();
  const sheetName=wb.SheetNames.find(n=>norm(n).replace(/\s+/g,'_')==='faktur_pembelian');
  if(!sheetName) throw new Error('Sheet FAKTUR_PEMBELIAN tidak ditemukan. Gunakan template Faktur V2 terbaru.');
  const raw=window.XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{defval:'',blankrows:false});
  const rows=raw.map((row,index)=>({row,line:Number.isInteger(row.__rowNum__)?row.__rowNum__+1:index+2}))
    .filter(({row})=>['Nomor Faktur','Supplier','Kode Produk','Nama Produk','Qty'].some(k=>text(row[k])!==''));
  const m=master(),map=new Map(),found=[];
  for(const {row,line} of rows){
    const number=text(row['Nomor Faktur']),date=dateOnly(row['Tanggal Faktur']);
    const supplier=canonicalSupplier(row.Supplier,m),payment=paymentType(row['Jenis Pembayaran']);
    const dueDate=dateOnly(row['Tanggal Jatuh Tempo']),tax=taxMethod(row['Metode PPN']);
    const qty=num(row.Qty),buyPrice=num(row['Harga Beli']),sellPrice=num(row['Harga Jual']);
    const code=text(row['Kode Produk']),existing=byCode(code,m);
    const name=text(existing?.['Nama Produk'])||text(row['Nama Produk']);
    const invoiceMethod=Object.prototype.hasOwnProperty.call(row,'Metode Diskon Faktur')
      ?discountMethod(row['Metode Diskon Faktur']):'nominal';
    const itemMethod=Object.prototype.hasOwnProperty.call(row,'Metode Diskon Item')
      ?discountMethod(row['Metode Diskon Item']):'nominal';
    const invoiceDiscountValue=num(row['Diskon Faktur']),itemDiscountValue=num(row['Diskon Item']);
    const statedScope=norm(row['Lingkup Diskon']);
    const discountScope=statedScope==='diskon global'?'global'
      :statedScope==='diskon per item'?'per_item'
      :!statedScope?(invoiceDiscountValue>0?'global':'per_item'):'';
    const globalTax=num(row['PPN Global (%)']),itemTax=num(row['PPN Item (%)']);
    if(!number){found.push(`Baris ${line}: Nomor Faktur wajib diisi.`);continue;}
    if(!date){found.push(`Baris ${line}: Tanggal Faktur wajib diisi.`);continue;}
    if(!supplier){found.push(`Baris ${line}: Nama perusahaan supplier tidak ditemukan pada Master Supplier.`);continue;}
    if(!payment){found.push(`Baris ${line}: Jenis Pembayaran harus Cash atau Tempo.`);continue;}
    if(payment==='Tempo'&&!dueDate){found.push(`Baris ${line}: Tanggal Jatuh Tempo wajib jika pembayaran Tempo.`);continue;}
    if(!tax){found.push(`Baris ${line}: Metode PPN harus Global atau Per Item.`);continue;}
    if(!discountScope){found.push(`Baris ${line}: Lingkup Diskon harus Diskon Global atau Diskon Per Item.`);continue;}
    if(discountScope==='global'&&itemDiscountValue>0){found.push(`Baris ${line}: Diskon Item harus kosong atau 0 saat Diskon Global dipilih.`);continue;}
    if(discountScope==='per_item'&&invoiceDiscountValue>0){found.push(`Baris ${line}: Diskon Faktur harus kosong atau 0 saat Diskon Per Item dipilih.`);continue;}
    if(qty<=0){found.push(`Baris ${line}: Qty harus lebih dari 0.`);continue;}
    if(!existing&&!name){found.push(`Baris ${line}: Produk baru wajib memiliki Nama Produk.`);continue;}
    if(buyPrice<0||sellPrice<0){found.push(`Baris ${line}: Harga tidak boleh negatif.`);continue;}
    if((discountScope==='global'&&!invoiceMethod)||(discountScope==='per_item'&&!itemMethod)){
      found.push(`Baris ${line}: Metode Diskon yang dipilih harus Persen atau Rupiah.`);continue;
    }
    if(invoiceDiscountValue<0||itemDiscountValue<0||
       (invoiceMethod==='percentage'&&invoiceDiscountValue>100)||
       (itemMethod==='percentage'&&itemDiscountValue>100)){
      found.push(`Baris ${line}: Diskon harus nonnegatif; Persen maksimal 100.`);continue;
    }
    if(globalTax<0||globalTax>100||itemTax<0||itemTax>100){
      found.push(`Baris ${line}: Tarif PPN harus antara 0 dan 100.`);continue;
    }
    const key=norm(number);
    if(!map.has(key)) map.set(key,{
      id:uid('INV'),number,date,supplierCode:supplier,supplierName:supplierName(supplier,m),
      paymentType:payment,dueDate,taxMethod:tax,taxPercent:globalTax,discount:0,discountScope,
      discountMethod:invoiceMethod,discountValue:invoiceDiscountValue,
      otherCost:num(row['Biaya Lain']),notes:text(row.Catatan),
      status:'draft',stockApplied:false,source:'excel-v2-single-sheet',items:[],createdAt:nowIso()
    });
    const inv=map.get(key);
    if(norm(inv.supplierCode)!==norm(supplier)) found.push(`Faktur ${number}: Supplier berbeda pada baris ${line}.`);
    if(inv.date!==date||inv.paymentType!==payment||inv.dueDate!==dueDate||
       inv.taxMethod!==tax||inv.taxPercent!==globalTax||
       inv.discountScope!==discountScope||inv.discountMethod!==invoiceMethod||inv.discountValue!==invoiceDiscountValue||
       inv.otherCost!==num(row['Biaya Lain'])){
      found.push(`Faktur ${number}: Data header berbeda pada baris ${line}.`);
    }
    inv.items.push({
      code:existing?text(existing['Kode Produk']):code,originalCode:code,name,qty,
      unit:text(row.Satuan)||text(existing?.Satuan),buyPrice,
      discount:discountScope==='per_item'?discountAmount(itemMethod,itemDiscountValue,qty*buyPrice):0,
      discountMethod:itemMethod,discountValue:itemDiscountValue,taxPercent:itemTax,
      sellPrice,batch:text(row.Batch),expired:dateOnly(row['Tanggal Expired']),
      notes:text(row.Catatan),supplier:inv.supplierCode,
      productStatus:existing?'existing':'new',reviewRequired:!existing
    });
  }
  for(const inv of map.values()){
    const itemSubtotal=inv.items.reduce((sum,item)=>sum+subtotal(item),0);
    inv.discount=inv.discountScope==='global'
      ?discountAmount(inv.discountMethod,inv.discountValue,itemSubtotal):0;
  }
  preview=[...map.values()];issues=found;
  const body=document.getElementById('invoice-import-preview-body');
  if(body) body.innerHTML=preview.length?preview.map(inv=>`<tr><td>${escapeHtml(inv.number)}</td><td>${escapeHtml(inv.date||'—')}</td><td>${escapeHtml(inv.supplierName||inv.supplierCode)}</td><td>${inv.items.length}</td><td>${rupiah(total(inv))}</td><td><span class="ops-badge warn">Draft</span>${inv.items.some(i=>i.productStatus==='new')?' <span class="ops-badge neutral">Produk Baru</span>':''}</td></tr>`).join(''):'<tr><td colspan="6" class="empty-table-state">Tidak ada faktur valid.</td></tr>';
  const summary=document.getElementById('invoice-import-summary');
  if(summary){summary.hidden=false;summary.innerHTML=`<strong>${preview.length} faktur</strong> terbaca dari ${rows.length} baris. Semua akan disimpan sebagai Draft.${issues.length?`<br><strong>${issues.length} masalah validasi perlu diperbaiki.</strong>`:''}`;}
  const previewBox=document.getElementById('invoice-import-preview');if(previewBox)previewBox.hidden=false;
  const actions=document.getElementById('invoice-import-actions');if(actions)actions.hidden=false;
  if(issues.length) await dialog().warning('Periksa Data Faktur',issues.slice(0,10).join('\n')+(issues.length>10?`\nDan ${issues.length-10} masalah lainnya.`:''),{confirmText:'Perbaiki'});
}

async function saveDraft(){
  if(!preview.length) return dialog().warning('Belum Ada Preview','Baca dan preview file Faktur V2 terlebih dahulu.');
  if(issues.length) return dialog().warning('Import Belum Dapat Dilakukan','Masih ada masalah validasi pada preview. Perbaiki file Excel lalu baca ulang.');
  const list=invoices(); let added=0,updated=0,skipped=0;
  for(const incoming of preview){
    const idx=list.findIndex(x=>invoiceKey(x)===invoiceKey(incoming));
    if(idx<0){ list.unshift(incoming); added++; continue; }
    const old=list[idx];
    if(old.stockApplied || norm(old.status)==='confirmed'){
      await dialog().error(
        'Faktur Sudah Dikonfirmasi',
        `Faktur ${incoming.number} dari ${incoming.supplierName||incoming.supplierCode} sudah pernah diproses sebagai Barang Masuk dan tidak dapat diimpor ulang.`
      );
      skipped++;
      continue;
    }
    const replace=await dialog().warning(
      'Ganti Draft Lama?',
      `Draft faktur ${incoming.number} dari ${incoming.supplierName||incoming.supplierCode} sudah ada. Seluruh isi Draft lama akan diganti oleh file terbaru; item tidak akan digabungkan.`,
      {confirmText:'Ganti Draft',cancelText:'Batalkan',showCancel:true}
    );
    if(!replace){skipped++;continue;}
    incoming.id=old.id;
    incoming.createdAt=old.createdAt;
    incoming.updatedAt=nowIso();
    list[idx]=incoming;
    updated++;
  }
  await writeStore(STORE_KEYS.invoices,list);
  await dialog().success('Faktur Draft Berhasil Disimpan',`Baru: ${added}\nDiperbarui: ${updated}\nDilewati: ${skipped}\n\nStok belum berubah. Stok baru berubah setelah Barang Masuk dikonfirmasi.`,{confirmText:'Selesai'});
  preview=[]; issues=[];
  if(fileInput()) fileInput().value='';
  const body=document.getElementById('invoice-import-preview-body'); if(body) body.innerHTML='<tr><td colspan="6" class="empty-table-state">Belum ada preview faktur.</td></tr>';
  const summary=document.getElementById('invoice-import-summary'); if(summary){ summary.hidden=true; summary.textContent=''; }
  document.querySelector('[data-view="purchase-invoices"]')?.click();
}

function labelOf(el){ return norm(el?.textContent); }
document.addEventListener('click',async event=>{
  const action=event.target.closest('button,a');
  if(!action || !inInvoiceSection(action)) return;
  const label=labelOf(action);
  try{
    if(label.includes('baca & preview')){ event.preventDefault(); event.stopImmediatePropagation(); await buildPreview(); return; }
    if(label.includes('simpan draft import')){ event.preventDefault(); event.stopImmediatePropagation(); await saveDraft(); return; }
  }catch(err){ console.error('AT-08 Faktur:',err); await dialog().error('Faktur Tidak Dapat Diproses',err?.message||String(err),{confirmText:'Tutup'}); }
},true);

window.KasirProInvoiceAT08=Object.freeze({downloadTemplate:makeTemplate,preview:buildPreview,saveDraft});
